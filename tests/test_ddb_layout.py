import sys
import tempfile
import unittest
import io
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
    missing_phone_cpfs_to_csv_bytes,
    sms_enviados_xlsx_bytes,
    sms_xlsx_bytes,
    load_ddb_lookup_for_uf,
    transform_c6_dataframe,
)


class DdbLookupTests(unittest.TestCase):
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
        self.assertEqual(csv_text.splitlines()[0], "cpf,beneficio,BDD")
        self.assertIn("03/04/2005", csv_text)

    def test_final_xlsx_has_requested_bdd_column_order(self):
        content = dataframe_to_xlsx_bytes(self._frame())
        sheet = load_workbook(io.BytesIO(content), read_only=True).active
        headers = [cell.value for cell in next(sheet.iter_rows())]
        self.assertEqual(headers, FINAL_COLUMNS)
        self.assertEqual(headers[-4:], ["COEF", "BRUTO", "TROCO", "BDD"])


if __name__ == "__main__":
    unittest.main()
