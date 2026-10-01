import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock
from urllib.error import URLError
from urllib.request import Request


APP_DIR = Path(__file__).resolve().parents[1] / "outputs" / "c6_streamlit_app"
sys.path.insert(0, str(APP_DIR))

from argus_api import (  # noqa: E402
    ARGUS_ROOMS,
    ArgusApiError,
    ArgusUploadResult,
    _request_json,
    file_was_uploaded,
    load_argus_config,
    record_upload,
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
    def test_only_the_three_named_rooms_are_valid(self):
        self.assertEqual(ARGUS_ROOMS, ("PADRÃO", "SALA POLLY", "SALA CLEITON"))
        for room in ("", "SALA DESCONHECIDA", "hash-antiga"):
            with self.subTest(room=room), self.assertRaisesRegex(ArgusApiError, "Sala da Argus"):
                upload_mailing_csv_bytes(
                    b"CPF;Nome\n123;Teste\n",
                    "mailing.csv",
                    room,
                    relay_secret="segredo-relay",
                    opener=QueueOpener(),
                )

    def test_upload_sends_room_to_relay_and_requires_argus_confirmation(self):
        opener = QueueOpener(
            {
                "idArquivo": 777,
                "codStatus": 1,
                "descStatus": "Sucesso",
                "argusConfirmed": True,
                "room": "SALA POLLY",
                "skillHash": "61vm1mnizndyeou",
                "relay": "scagi",
            }
        )
        result = upload_mailing_csv_bytes(
            b"CPF;Nome\n123;Teste\n",
            "mailing.csv",
            "SALA POLLY",
            relay_url="https://scagi.test/api/argus/upload",
            relay_secret="segredo-relay",
            opener=opener,
        )

        self.assertTrue(result.argus_confirmed)
        self.assertEqual(result.id_arquivo, 777)
        self.assertEqual(result.room, "SALA POLLY")
        request = opener.requests[0][0]
        self.assertEqual(request.full_url, "https://scagi.test/api/argus/upload")
        self.assertEqual(request.get_header("X-argus-relay-secret"), "segredo-relay")
        self.assertEqual(request.get_header("X-argus-filename"), "mailing.csv")
        self.assertEqual(request.get_header("X-argus-room"), "SALA POLLY")
        self.assertEqual(request.data, b"CPF;Nome\n123;Teste\n")

    def test_server_response_without_argus_confirmation_is_not_success(self):
        opener = QueueOpener(
            {"idArquivo": 777, "codStatus": 1, "descStatus": "Servidor recebeu"}
        )
        with self.assertRaisesRegex(ArgusApiError, "Servidor recebeu"):
            upload_mailing_csv_bytes(
                b"CPF;Nome\n123;Teste\n",
                "mailing.csv",
                "PADRÃO",
                relay_secret="segredo-relay",
                opener=opener,
            )

    def test_success_log_only_marks_confirmed_upload_as_sent(self):
        with tempfile.TemporaryDirectory() as directory:
            csv_path = Path(directory) / "mailing.csv"
            csv_path.write_text("CPF;Nome\n123;Teste\n", encoding="utf-8")
            result = ArgusUploadResult(
                id_arquivo=1,
                cod_status=1,
                desc_status="Sucesso",
                skill_hash="2jovr4cpoyeb9bf",
                room="PADRÃO",
                argus_confirmed=True,
            )
            with mock.patch.dict("os.environ", {"LOCALAPPDATA": directory}, clear=False):
                record_upload(csv_path, result)
                self.assertTrue(file_was_uploaded(csv_path))

    @unittest.skipUnless(sys.platform == "win32", "Fallback com curl.exe existe apenas no Windows")
    def test_url_error_retries_with_windows_curl(self):
        def fake_curl(command, **kwargs):
            output_path = Path(command[command.index("--output") + 1])
            output_path.write_text(json.dumps({"codStatus": 1}), encoding="utf-8")
            return subprocess.CompletedProcess(command, 0, stdout="200", stderr="")

        request = Request("https://scagi.test/api/argus/upload", method="POST", data=b"csv")
        with mock.patch("argus_api.urllib_request.urlopen", side_effect=URLError("getaddrinfo failed")):
            with mock.patch("argus_api.subprocess.run", side_effect=fake_curl) as run:
                payload = _request_json(request, timeout=60, opener=None)

        self.assertEqual(payload["codStatus"], 1)
        self.assertEqual(run.call_count, 1)

    @unittest.skipUnless(sys.platform == "win32", "DPAPI existe apenas no Windows")
    def test_relay_secret_is_encrypted_and_can_be_loaded(self):
        with tempfile.TemporaryDirectory() as directory:
            with (
                mock.patch.dict("os.environ", {"LOCALAPPDATA": directory}, clear=False),
                mock.patch("argus_api._protect_windows_data", side_effect=lambda value: b"protected:" + value),
                mock.patch("argus_api._unprotect_windows_data", side_effect=lambda value: value.removeprefix(b"protected:")),
            ):
                path = save_argus_config(
                    "segredo-relay-local",
                    "https://scagi.test/api/argus/upload",
                )
                contents = path.read_text(encoding="utf-8")
                self.assertNotIn("segredo-relay-local", contents)
                loaded = load_argus_config()
        self.assertIsNotNone(loaded)
        self.assertEqual(loaded.relay_url, "https://scagi.test/api/argus/upload")
        self.assertEqual(loaded.relay_secret, "segredo-relay-local")


if __name__ == "__main__":
    unittest.main()
