import sys
import tempfile
import unittest
import io
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


APP_DIR = Path(__file__).resolve().parents[1] / "outputs" / "c6_streamlit_app"
sys.path.insert(0, str(APP_DIR))

from c6_processor import (  # noqa: E402
    ARGUS_DISCADORA_COLUMNS,
    FINAL_COLUMNS,
    build_argus_discadora_dataframe,
    build_sms_dataframe,
    build_sms_enviados_dataframe,
    dataframe_to_xlsx_bytes,
    enrich_with_ddb,
    enrich_with_phone_sources,
    missing_phone_cpfs_to_csv_bytes,
    sms_enviados_xlsx_bytes,
    sms_xlsx_bytes,
    _normalize_cpf_key,
    _normalize_phone,
    load_ddb_lookup_for_uf,
    transform_c6_dataframe,
)


class DdbLookupTests(unittest.TestCase):
    def test_numeric_excel_identifiers_do_not_gain_trailing_zero(self):
        self.assertEqual(_normalize_cpf_key("12345678901.0"), "12345678901")
        self.assertEqual(_normalize_phone("11999999999.0"), "11999999999")
        self.assertEqual(_normalize_phone("5511999999999"), "11999999999")

    def test_phone_lookup_accepts_numeric_excel_like_cpf_and_phone(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            (base / "hub.csv").write_text(
                "CPF;TELEFONE;DATA DA ULTIMA MODIFICACAO\n12345678901.0;11999999999.0;2026-07-01\n",
                encoding="utf-8",
            )
            frame = pd.DataFrame({"CPF": ["12345678901"]})
            result = enrich_with_phone_sources(frame, base)
            self.assertEqual(result.loc[0, "TELEFONE BASE"], "11999999999")

    def test_phone_period_accepts_alternate_hub_date_header(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            (base / "hub.csv").write_text(
                "CPF;TELEFONE;DATA_ULTIMA_MODIFICACAO_ARQUIVO_ORIGINAL\n"
                "12345678901;11999999999;2025-12-31\n"
                "10987654321;11888888888;2026-07-01\n",
                encoding="utf-8",
            )
            frame = pd.DataFrame({"CPF": ["12345678901", "10987654321"]})
            result = enrich_with_phone_sources(frame, base, telefone_data_inicial=date(2026, 1, 1))
            self.assertEqual(result["TELEFONE BASE"].tolist(), ["", "11888888888"])

    def test_lemit_only_fills_cpfs_that_stayed_without_base_phone(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory) / "hub"
            base.mkdir()
            (base / "hub.csv").write_text(
                "CPF;TELEFONE;DATA DA ULTIMA MODIFICACAO\n"
                "12345678901;11999999999;2026-07-01\n",
                encoding="utf-8",
            )
            lemit = Path(directory) / "retorno_lemit.csv"
            lemit.write_text(
                "CPF;DDD;TELEFONE\n"
                "12345678901;11;888888888\n"
                "10987654321;21;977777777\n",
                encoding="utf-8",
            )
            frame = pd.DataFrame({"CPF": ["12345678901", "10987654321"]})
            result = enrich_with_phone_sources(frame, base, lemit)
            self.assertEqual(result["TELEFONE BASE"].tolist(), ["11999999999", "21977777777"])

    def test_new_dtf_beneficio_header_is_accepted(self):
        raw = pd.DataFrame(
            {
                "ovc_cpf": ["123"],
                "Benefício": ["111.334.308-4"],
                "cr_parcelas_contrato": ["10"],
                "cr_parcelas_em_aberto": ["20"],
                "cr_valor_parcela": ["100,00"],
                "cr_saldo_refin": ["50,00"],
                "cr_taxa": ["1,5"],
                "dc_nome_1": ["Pessoa"],
                "dc_data_nascimento_1": ["01/01/1980"],
                "tf_ddd_1": ["11"],
                "tf_telefone_1": ["999999999"],
            }
        )
        result = transform_c6_dataframe(raw)
        self.assertEqual(result.loc[0, "BENEFÍCIO"], "1113343084")

    def test_enriches_ddb_using_uf_and_normalized_benefit(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            (base / "base_AP feito c6.csv").write_text(
                "NB;DDB\n111.334.308-4;24/02/1999\n", encoding="utf-8"
            )
            frame = pd.DataFrame({"BENEFÍCIO": ["1113343084", "999"]})
            result = enrich_with_ddb(frame, "EXTRACAO C6 AP.csv", base)
            self.assertEqual(result["BDD"].tolist(), ["24/02/1999", ""])

    def test_merges_multiple_ddb_bases_for_same_uf(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            (base / "MACICA NOV25 ESTADO SP 1 feito c6.csv").write_text(
                "NB;DDB\n111;24/02/1999\n", encoding="utf-8"
            )
            (base / "MACICA NOV25 ESTADO SP 2 feito c6.csv").write_text(
                "NB;DDB\n222;03/04/2005\n", encoding="utf-8"
            )
            frame = pd.DataFrame({FINAL_COLUMNS[1]: ["111", "222", "333"]})
            result = enrich_with_ddb(frame, "EXTRACAO C6 SP.csv", base)
            self.assertEqual(result["BDD"].tolist(), ["24/02/1999", "03/04/2005", ""])

    def test_ignores_empty_ddb_file_when_other_uf_bases_have_data(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            (base / "MACICA NOV25 ESTADO SP 1 feito c6.csv").write_text(
                "NB;DDB\n111;24/02/1999\n", encoding="utf-8"
            )
            (base / "MACICA NOV25 ESTADO SP 2 feito c6.csv").write_text("", encoding="utf-8")
            (base / "MACICA NOV25 ESTADO SP 3 feito c6.csv").write_text(
                "NB;DDB\n222;03/04/2005\n", encoding="utf-8"
            )
            frame = pd.DataFrame({FINAL_COLUMNS[1]: ["111", "222"]})
            result = enrich_with_ddb(frame, "EXTRACAO C6 SP.csv", base)
            self.assertEqual(result["BDD"].tolist(), ["24/02/1999", "03/04/2005"])

    def test_rejects_conflicting_ddb_for_same_benefit(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            (base / "base_AC feito c6.csv").write_text(
                "NB;DDB\n123;01/01/2000\n123;02/02/2000\n", encoding="utf-8"
            )
            with self.assertRaisesRegex(ValueError, "conflitante"):
                load_ddb_lookup_for_uf("AC", base)


class ArgusLayoutTests(unittest.TestCase):
    def _frame(self):
        return pd.DataFrame(
            {
                "CPF": ["12345678901"],
                "TAXA": [1.5],
                "FEZ EM": [10],
                "FALTAM": [20],
                "PARCELA": [100.0],
                "SALDO DEV": [50.0],
                "TROCO": [200.0],
                "TELEFONE BASE": ["11999999999"],
                "NOME": ["Pessoa Teste"],
                "DATA DE NASC": [pd.Timestamp("1980-02-01")],
                "DDD C6": ["11"],
                "TEL C6": ["888888888"],
                "TELEFONE C6": ["11888888888"],
                "COEF": [0.0228],
                "BRUTO": [250.0],
                "BDD": ["03/04/2005"],
                "BENEFÍCIO": ["1234567890"],
            }
        )

    def test_layout_padrao_c6_has_exact_columns_and_two_phones(self):
        frame = self._frame()
        result = build_argus_discadora_dataframe(frame, "AP 1 REFIN INSS C6")
        self.assertEqual(result.columns.tolist(), ARGUS_DISCADORA_COLUMNS)
        self.assertEqual(result.columns.tolist().count("Telefone"), 2)
        self.assertEqual(result.iloc[0, 7], "11999999999")
        self.assertEqual(result.iloc[0, 10], "11888888888")
        self.assertEqual(result.iloc[0, 11], "03/04/2005")
        self.assertEqual(result.iloc[0, 12], "1234567890")

    def test_ddb_is_present_in_every_generated_spreadsheet_and_csv(self):
        frame = self._frame()
        workbook_outputs = (
            dataframe_to_xlsx_bytes(frame),
            sms_xlsx_bytes(frame),
            sms_enviados_xlsx_bytes(frame),
        )
        for content in workbook_outputs:
            sheet = load_workbook(io.BytesIO(content), read_only=True).active
            headers = [cell.value for cell in next(sheet.iter_rows())]
            self.assertIn("BDD", headers)

        self.assertIn("BDD", build_sms_dataframe(frame).columns)
        self.assertIn("BDD", build_sms_enviados_dataframe(frame).columns)

        frame.loc[0, "TELEFONE BASE"] = ""
        csv_text = missing_phone_cpfs_to_csv_bytes(frame).decode("utf-8-sig")
        csv_lines = csv_text.splitlines()
        self.assertEqual(csv_lines[0], "cpf")
        self.assertEqual(len(csv_lines), 2)
        self.assertEqual(csv_lines[1], "12345678901")

    def test_missing_phone_csv_ignores_cpfs_that_already_have_base_phone(self):
        frame = pd.DataFrame(
            {
                "CPF": ["12345678901", "10987654321"],
                FINAL_COLUMNS[1]: ["111", "222"],
                "BDD": ["01/01/2000", "02/02/2000"],
                "TELEFONE BASE": ["11999999999", ""],
            }
        )
        csv_text = missing_phone_cpfs_to_csv_bytes(frame).decode("utf-8-sig")
        self.assertEqual(csv_text.splitlines(), ["cpf", "10987654321"])

    def test_final_xlsx_has_requested_bdd_column_order(self):
        content = dataframe_to_xlsx_bytes(self._frame())
        sheet = load_workbook(io.BytesIO(content), read_only=True).active
        headers = [cell.value for cell in next(sheet.iter_rows())]
        self.assertEqual(headers, FINAL_COLUMNS)
        self.assertEqual(headers[-4:], ["COEF", "BRUTO", "TROCO", "BDD"])


if __name__ == "__main__":
    unittest.main()
