import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

import pandas as pd


APP_DIR = Path(__file__).resolve().parents[1] / "outputs" / "c6_streamlit_app"
sys.path.insert(0, str(APP_DIR))

from c6_processor import (  # noqa: E402
    BANK_NOT_IDENTIFIED_MESSAGE,
    build_argus_campaign_name,
    build_lemit_processed_stem,
    extract_argus_campaign_from_filename,
    extract_uf_from_filename,
    find_next_argus_campaign_number,
    identify_bank_from_filename,
    transform_c6_dataframe,
)


def _raw_dtf_frame() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "ovc_cpf": ["12345678901"],
            "cr_matricula": ["111.334.308-4"],
            "cr_parcelas_contrato": ["84"],
            "cr_parcelas_em_aberto": ["72"],
            "cr_valor_parcela": ["100,00"],
            "cr_saldo_refin": ["1000,00"],
            "cr_taxa": ["1,66"],
            "dc_nome_1": ["Pessoa Teste"],
            "dc_data_nascimento_1": ["01/02/1980"],
            "tf_ddd_1": ["11"],
            "tf_telefone_1": ["999999999"],
        }
    )


class BankDetectionTests(unittest.TestCase):
    def test_identifies_c6_and_digio_from_filename(self):
        self.assertEqual(identify_bank_from_filename("2022 refin c6 SP.csv").code, "c6")
        self.assertEqual(identify_bank_from_filename("45k_DIGIO_SP_pro.csv").code, "digio")
        self.assertEqual(identify_bank_from_filename("45k_digio_SP_pro.csv").code, "digio")
        self.assertEqual(identify_bank_from_filename("45k_dígio_SP_pro.csv").code, "digio")

    def test_blocks_unknown_bank(self):
        with self.assertRaisesRegex(ValueError, "Banco nao identificado"):
            identify_bank_from_filename("base_sp_sem_banco.csv")

        self.assertIn('"C6"', BANK_NOT_IDENTIFIED_MESSAGE)
        self.assertIn('"Digio"', BANK_NOT_IDENTIFIED_MESSAGE)

    def test_digio_uses_fixed_coefficient_and_beneficio_as_matricula(self):
        bank = identify_bank_from_filename("exportacao_dados_digio_SP.csv")
        result = transform_c6_dataframe(_raw_dtf_frame(), bank)

        self.assertEqual(result.loc[0, "BENEFÍCIO"], "1113343084")
        self.assertEqual(result.loc[0, "COEF"], 0.022)
        self.assertAlmostEqual(result.loc[0, "BRUTO"], 4545.454545, places=5)

    def test_c6_coefficient_is_preserved(self):
        bank = identify_bank_from_filename("exportacao_dados_c6_SP.csv")
        result = transform_c6_dataframe(_raw_dtf_frame(), bank)

        self.assertEqual(result.loc[0, "COEF"], 0.0228)
        self.assertAlmostEqual(result.loc[0, "BRUTO"], 4385.964912, places=5)

    def test_dtf_without_c6_phone_columns_is_accepted(self):
        raw = _raw_dtf_frame().drop(columns=["tf_ddd_1", "tf_telefone_1"])
        result = transform_c6_dataframe(raw, identify_bank_from_filename("exportacao_dados_c6_SP.csv"))

        self.assertEqual(result.loc[0, "DDD C6"], "")
        self.assertEqual(result.loc[0, "TEL C6"], "")
        self.assertEqual(result.loc[0, "TELEFONE C6"], "")

    def test_bank_changes_processed_filename_stem(self):
        self.assertEqual(extract_uf_from_filename("exportacao_45k_digio_SP_pro"), "SP")
        self.assertEqual(
            build_lemit_processed_stem("exportacao_45k_digio_SP_pro.csv", 2575),
            "2575 refin digio SP (1 ou 0 cartões) sem RL pra disc",
        )

    def test_argus_campaign_number_uses_current_month_state_folder(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            old_month = base / "JULHO 2026" / "SP"
            old_month.mkdir(parents=True)
            (old_month / "SP 118 REFIN INSS C6 lista antiga.csv").write_text("", encoding="utf-8")
            current_month = base / "AGOSTO 2026" / "SP"
            current_month.mkdir(parents=True)
            (current_month / "SP 17 REFIN INSS C6 lista atual.csv").write_text("", encoding="utf-8")

            self.assertEqual(
                find_next_argus_campaign_number("SP", base, date(2026, 8, 10)),
                18,
            )
            self.assertEqual(
                build_argus_campaign_name("2575 refin c6 SP sem RL.xlsx", base, date(2026, 8, 10)),
                "SP 18 REFIN INSS C6",
            )

    def test_extracts_campaign_from_generated_argus_filename(self):
        self.assertEqual(
            extract_argus_campaign_from_filename(
                r"G:\Meu Drive\GRUPO IMPERIO\LEADS\ARGUS\DISC\AGOSTO 2026\SP\SP 18 REFIN INSS C6 lista.csv"
            ),
            "SP 18 REFIN INSS C6",
        )


if __name__ == "__main__":
    unittest.main()
