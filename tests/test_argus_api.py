import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


APP_DIR = Path(__file__).resolve().parents[1] / "outputs" / "c6_streamlit_app"
sys.path.insert(0, str(APP_DIR))

from argus_api import (  # noqa: E402
    ArgusApiError,
    load_argus_config,
    resolve_skill_hash,
    save_argus_config,
    upload_mailing_csv_bytes,
)


class FakeResponse:
    def __init__(self, payload):
        self.payload = payload

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def read(self):
        return json.dumps(self.payload).encode("utf-8")


class QueueOpener:
    def __init__(self, *payloads):
        self.payloads = list(payloads)
        self.requests = []

    def __call__(self, request, timeout):
        self.requests.append((request, timeout))
        return FakeResponse(self.payloads.pop(0))


class ArgusApiTests(unittest.TestCase):
    def test_resolve_padrao_ignores_accent_and_case(self):
        skills = [{"descricaoSkill": "Padrão", "hashEndpointSkill": "hash-real"}]
        self.assertEqual(resolve_skill_hash(skills, "PADRAO"), "hash-real")

    def test_duplicate_skill_name_with_different_hashes_is_rejected(self):
        skills = [
            {"descricaoSkill": "PADRÃO", "hashEndpointSkill": "hash-1"},
            {"descricaoSkill": "PADRÃO", "hashEndpointSkill": "hash-2"},
        ]
        with self.assertRaises(ArgusApiError):
            resolve_skill_hash(skills, "PADRÃO")

    def test_upload_lists_skills_then_posts_expected_multipart(self):
        opener = QueueOpener(
            {
                "codStatus": 1,
                "retornoGetSkillsItens": [
                    {"descricaoSkill": "PADRÃO", "hashEndpointSkill": "abc123"}
                ],
            },
            {"idArquivo": 987, "codStatus": 1, "descStatus": "Sucesso"},
        )
        result = upload_mailing_csv_bytes(
            b"CPF;Nome\n123;Teste\n",
            "mailing.csv",
            "token-de-teste",
            "PADRÃO",
            opener=opener,
        )

        self.assertEqual(result.id_arquivo, 987)
        self.assertEqual(len(opener.requests), 2)
        skills_request = opener.requests[0][0]
        upload_request = opener.requests[1][0]
        self.assertTrue(skills_request.full_url.endswith("/cmd/skills"))
        self.assertTrue(upload_request.full_url.endswith("/abc123/uploadmailing"))
        self.assertEqual(upload_request.get_header("Token-signature"), "token-de-teste")
        self.assertIn('name=""; filename="mailing.csv"', upload_request.data.decode("utf-8"))

    def test_upload_with_real_hash_skips_skill_lookup(self):
        opener = QueueOpener({"idArquivo": 321, "codStatus": 1, "descStatus": "Sucesso"})
        result = upload_mailing_csv_bytes(
            b"CPF;Nome\n123;Teste\n",
            "mailing.csv",
            "token-de-teste",
            "5ni7te02zzh1ka8",
            opener=opener,
        )

        self.assertEqual(result.id_arquivo, 321)
        self.assertEqual(len(opener.requests), 1)
        upload_request = opener.requests[0][0]
        self.assertTrue(upload_request.full_url.endswith("/5ni7te02zzh1ka8/uploadmailing"))

    @unittest.skipUnless(sys.platform == "win32", "DPAPI existe apenas no Windows")
    def test_saved_token_is_encrypted_and_can_be_loaded(self):
        with tempfile.TemporaryDirectory() as directory:
            with mock.patch.dict("os.environ", {"LOCALAPPDATA": directory}, clear=False):
                path = save_argus_config("token-super-secreto", "5ni7te02zzh1ka8")
                self.assertNotIn("token-super-secreto", path.read_text(encoding="utf-8"))
                loaded = load_argus_config()
        self.assertIsNotNone(loaded)
        self.assertEqual(loaded.token_signature, "token-super-secreto")
        self.assertEqual(loaded.skill_name, "5ni7te02zzh1ka8")


if __name__ == "__main__":
    unittest.main()
