import json
import sys
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest import mock
from urllib.error import HTTPError


APP_DIR = Path(__file__).resolve().parents[1] / "outputs" / "c6_streamlit_app"
sys.path.insert(0, str(APP_DIR))

from argus_api import (  # noqa: E402
    ArgusApiError,
    DEFAULT_SKILL_REFERENCE,
    DEFAULT_TOKEN_SIGNATURE,
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


class MixedOpener:
    def __init__(self, *items):
        self.items = list(items)
        self.requests = []

    def __call__(self, request, timeout):
        self.requests.append((request, timeout))
        item = self.items.pop(0)
        if isinstance(item, BaseException):
            raise item
        return FakeResponse(item)


def _http_error(status: int, payload: dict) -> HTTPError:
    body = json.dumps(payload).encode("utf-8")
    return HTTPError("https://argus.test/upload", status, "Erro", {}, BytesIO(body))


class ArgusApiTests(unittest.TestCase):
    def test_default_argus_access_does_not_expose_token(self):
        with tempfile.TemporaryDirectory() as directory:
            with mock.patch.dict("os.environ", {"LOCALAPPDATA": directory}, clear=False):
                loaded = load_argus_config()

        self.assertIsNone(loaded)
        self.assertEqual(DEFAULT_TOKEN_SIGNATURE, "")
        self.assertEqual(DEFAULT_SKILL_REFERENCE, "5ni7te02zzh1ka8")

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

    def test_upload_recovers_when_saved_hash_endpoint_is_invalid(self):
        opener = MixedOpener(
            _http_error(400, {"codStatus": -1, "descStatus": "EndPoint Inválido."}),
            {
                "codStatus": 1,
                "retornoGetSkillsItens": [
                    {"descricaoSkill": "LAYOUT PADRÃO C6", "hashEndpointSkill": "hash-atual"}
                ],
            },
            {"idArquivo": 654, "codStatus": 1, "descStatus": "Sucesso"},
        )
        result = upload_mailing_csv_bytes(
            b"CPF;Nome\n123;Teste\n",
            "mailing.csv",
            "token-de-teste",
            "hashantigo123",
            opener=opener,
        )

        self.assertEqual(result.id_arquivo, 654)
        self.assertEqual(result.skill_hash, "hash-atual")
        self.assertEqual(len(opener.requests), 3)
        self.assertTrue(opener.requests[0][0].full_url.endswith("/hashantigo123/uploadmailing"))
        self.assertTrue(opener.requests[1][0].full_url.endswith("/cmd/skills"))
        self.assertTrue(opener.requests[2][0].full_url.endswith("/hash-atual/uploadmailing"))

    @unittest.skipUnless(sys.platform == "win32", "DPAPI existe apenas no Windows")
    def test_saved_token_is_encrypted_and_can_be_loaded(self):
        with tempfile.TemporaryDirectory() as directory:
            with mock.patch.dict("os.environ", {"LOCALAPPDATA": directory}, clear=False):
                path = save_argus_config("token-super-secreto", "5ni7te02zzh1ka8")
                self.assertNotIn("token-super-secreto", path.read_text(encoding="utf-8"))
                loaded = load_argus_config()
        self.assertIsNotNone(loaded)
        self.assertEqual(loaded.token_signature, "token-super-secreto")
        self.assertEqual(loaded.skill_name, DEFAULT_SKILL_REFERENCE)


if __name__ == "__main__":
    unittest.main()
