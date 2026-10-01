from __future__ import annotations

import base64
import ctypes
import hashlib
import json
import os
import subprocess
import sys
import tempfile
from ctypes import wintypes
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable
from urllib import error as urllib_error
from urllib import request as urllib_request


DEFAULT_ARGUS_RELAY_URL = "https://scagi.grupoimperiofinanceira.com.br/api/argus/upload"
ARGUS_ROOMS = ("PADRÃO", "SALA POLLY", "SALA CLEITON")
CONFIG_DIR_NAME = "Auto C6BANK - Datafast for Argus"


class ArgusApiError(RuntimeError):
    pass


@dataclass(frozen=True)
class ArgusConfig:
    relay_url: str
    relay_secret: str = ""


@dataclass(frozen=True)
class ArgusUploadResult:
    id_arquivo: int
    cod_status: int
    desc_status: str
    skill_hash: str
    room: str
    argus_confirmed: bool


def upload_mailing_csv(
    csv_path: str | Path,
    room: str,
    *,
    relay_url: str = "",
    relay_secret: str = "",
    timeout: int = 600,
    opener: Callable | None = None,
) -> ArgusUploadResult:
    path = Path(csv_path)
    if not path.is_file():
        raise FileNotFoundError(f"CSV da Argus não encontrado: {path}")
    if path.suffix.lower() != ".csv":
        raise ValueError(f"O arquivo para upload precisa ser CSV: {path.name}")
    return upload_mailing_csv_bytes(
        path.read_bytes(),
        path.name,
        room,
        relay_url=relay_url,
        relay_secret=relay_secret,
        timeout=timeout,
        opener=opener,
    )


def upload_mailing_csv_bytes(
    csv_bytes: bytes,
    filename: str,
    room: str,
    *,
    relay_url: str = "",
    relay_secret: str = "",
    timeout: int = 600,
    opener: Callable | None = None,
) -> ArgusUploadResult:
    if not csv_bytes:
        raise ValueError("O CSV da Argus está vazio.")
    if not str(filename).lower().endswith(".csv"):
        raise ValueError("O arquivo para upload precisa ter extensão .csv.")
    room_name = _required_room(room)
    return _upload_mailing_csv_bytes_via_relay(
        csv_bytes,
        filename,
        room_name,
        relay_url=relay_url,
        relay_secret=relay_secret,
        timeout=timeout,
        opener=opener,
    )


def _upload_mailing_csv_bytes_via_relay(
    csv_bytes: bytes,
    filename: str,
    room: str,
    *,
    relay_url: str = "",
    relay_secret: str = "",
    timeout: int,
    opener: Callable | None,
) -> ArgusUploadResult:
    relay_secret = (relay_secret or os.environ.get("ARGUS_RELAY_SECRET", "")).strip()
    if not relay_secret:
        raise ArgusApiError("A chave de autenticação do relay não está configurada nesta máquina.")

    relay_url = (relay_url or os.environ.get("ARGUS_RELAY_URL", DEFAULT_ARGUS_RELAY_URL)).strip()
    if not relay_url:
        raise ArgusApiError("A URL do relay do SCAGI não está configurada nesta máquina.")

    request = urllib_request.Request(
        relay_url,
        data=csv_bytes,
        headers={
            "Accept": "application/json",
            "Content-Type": "text/csv; charset=utf-8",
            "X-Argus-Relay-Secret": relay_secret,
            "X-Argus-Filename": filename,
            "X-Argus-Room": room,
        },
        method="POST",
    )
    payload = _request_json(request, timeout=timeout, opener=opener)
    cod_status = _as_int(payload.get("codStatus"), default=-1)
    desc_status = str(payload.get("descStatus") or "").strip()
    error = payload.get("error") if isinstance(payload.get("error"), dict) else {}
    error_message = str(error.get("message") or "").strip()
    id_arquivo = _as_int(payload.get("idArquivo"), default=0)
    confirmed = payload.get("argusConfirmed") is True and cod_status >= 1 and id_arquivo > 0
    if not confirmed:
        raise ArgusApiError(
            error_message
            or desc_status
            or "O servidor respondeu, mas a Argus não confirmou o recebimento do CSV."
        )
    return ArgusUploadResult(
        id_arquivo=id_arquivo,
        cod_status=cod_status,
        desc_status=desc_status or "Recebimento confirmado pela Argus",
        skill_hash=str(payload.get("skillHash") or ""),
        room=str(payload.get("room") or room),
        argus_confirmed=True,
    )


def load_argus_config() -> ArgusConfig | None:
    environment_relay_url = os.environ.get("ARGUS_RELAY_URL", "").strip()
    environment_relay_secret = os.environ.get("ARGUS_RELAY_SECRET", "").strip()
    if environment_relay_secret:
        return ArgusConfig(
            environment_relay_url or DEFAULT_ARGUS_RELAY_URL,
            environment_relay_secret,
        )

    path = argus_config_path()
    if not path.exists():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        relay_secret = ""
        if payload.get("protected_relay_secret"):
            protected_relay_secret = base64.b64decode(payload["protected_relay_secret"])
            relay_secret = _unprotect_windows_data(protected_relay_secret).decode("utf-8")
        relay_url = str(payload.get("relay_url") or DEFAULT_ARGUS_RELAY_URL).strip()
        if not relay_secret:
            return None
    except Exception as exc:
        raise ArgusApiError(
            "Não consegui abrir a configuração protegida do relay. Configure o acesso novamente."
        ) from exc
    return ArgusConfig(relay_url, relay_secret)


def save_argus_config(
    relay_secret: str,
    relay_url: str = DEFAULT_ARGUS_RELAY_URL,
) -> Path:
    secret = str(relay_secret or "").strip()
    if not secret:
        raise ArgusApiError("Informe a chave de autenticação do relay.")
    protected_relay_secret = _protect_windows_data(secret.encode("utf-8"))
    path = argus_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(
            {
                "relay_url": str(relay_url or DEFAULT_ARGUS_RELAY_URL).strip(),
                "protected_relay_secret": base64.b64encode(protected_relay_secret).decode("ascii"),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return path


def file_was_uploaded(csv_path: str | Path) -> bool:
    digest = _file_sha256(Path(csv_path))
    log_path = argus_upload_log_path()
    if not log_path.exists():
        return False
    for line in log_path.read_text(encoding="utf-8").splitlines():
        try:
            record = json.loads(line)
        except json.JSONDecodeError:
            continue
        if (
            record.get("sha256") == digest
            and _as_int(record.get("cod_status"), 0) >= 1
            and record.get("argus_confirmed") is True
        ):
            return True
    return False


def record_upload(csv_path: str | Path, result: ArgusUploadResult) -> Path:
    path = Path(csv_path)
    log_path = argus_upload_log_path()
    log_path.parent.mkdir(parents=True, exist_ok=True)
    record = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "csv_path": str(path.resolve()),
        "filename": path.name,
        "sha256": _file_sha256(path),
        **asdict(result),
    }
    with log_path.open("a", encoding="utf-8") as stream:
        stream.write(json.dumps(record, ensure_ascii=False) + "\n")
    return log_path


def argus_config_dir() -> Path:
    local_appdata = os.environ.get("LOCALAPPDATA")
    if local_appdata:
        return Path(local_appdata) / CONFIG_DIR_NAME
    return Path.home() / ".auto_c6bank"


def argus_config_path() -> Path:
    return argus_config_dir() / "argus_config.json"


def argus_upload_log_path() -> Path:
    return argus_config_dir() / "argus_uploads.jsonl"


def _request_json(request: urllib_request.Request, *, timeout: int, opener: Callable | None) -> dict:
    open_request = opener or urllib_request.urlopen
    try:
        with open_request(request, timeout=timeout) as response:
            body = response.read().decode("utf-8-sig", errors="replace")
    except urllib_error.HTTPError as exc:
        detail = exc.read().decode("utf-8-sig", errors="replace")
        try:
            payload = json.loads(detail) if detail else {}
        except json.JSONDecodeError:
            payload = None
        if isinstance(payload, dict):
            return payload
        raise ArgusApiError(f"Argus respondeu HTTP {exc.code}: {detail}") from exc
    except urllib_error.URLError as exc:
        if opener is None and sys.platform == "win32":
            body = _request_with_windows_curl(request, timeout=timeout)
        else:
            raise ArgusApiError(f"Não foi possível conectar à Argus: {exc.reason}") from exc
    try:
        payload = json.loads(body) if body else {}
    except json.JSONDecodeError as exc:
        raise ArgusApiError(f"A Argus retornou uma resposta inválida: {body[:300]}") from exc
    if not isinstance(payload, dict):
        raise ArgusApiError("A Argus retornou um JSON em formato inesperado.")
    return payload


def _is_ssl_handshake_failure(exc: BaseException) -> bool:
    detail = str(getattr(exc, "reason", exc)).casefold()
    return "ssl" in detail and "handshake failure" in detail


def _request_with_windows_curl(request: urllib_request.Request, *, timeout: int) -> str:
    """Usa o TLS nativo do Windows quando o OpenSSL do Python é recusado pela Argus."""
    request_body_path: Path | None = None
    response_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(prefix="auto-c6-argus-response-", delete=False) as response_file:
            response_path = Path(response_file.name)

        command = [
            "curl.exe",
            "--silent",
            "--show-error",
            "--max-time",
            str(max(int(timeout), 1)),
            "--request",
            request.get_method(),
            "--url",
            request.full_url,
            "--output",
            str(response_path),
            "--write-out",
            "%{http_code}",
        ]
        for name, value in request.header_items():
            command.extend(("--header", f"{name}: {value}"))

        if request.data is not None:
            with tempfile.NamedTemporaryFile(prefix="auto-c6-argus-request-", delete=False) as request_file:
                request_file.write(request.data)
                request_body_path = Path(request_file.name)
            command.extend(("--data-binary", f"@{request_body_path}"))

        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=max(int(timeout), 1) + 5,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            check=False,
        )
        response_body = response_path.read_text(encoding="utf-8-sig", errors="replace")
        if completed.returncode != 0:
            detail = completed.stderr.strip() or f"curl retornou código {completed.returncode}"
            raise ArgusApiError(f"Não foi possível conectar à Argus pelo TLS do Windows: {detail}")
        try:
            status = int(completed.stdout.strip()[-3:])
        except ValueError as exc:
            raise ArgusApiError("A conexão com a Argus não informou o status HTTP.") from exc
        if status >= 400:
            raise ArgusApiError(f"Argus respondeu HTTP {status}: {response_body}")
        return response_body
    except subprocess.TimeoutExpired as exc:
        raise ArgusApiError("A conexão com a Argus excedeu o tempo limite.") from exc
    except FileNotFoundError as exc:
        raise ArgusApiError("O cliente HTTPS do Windows (curl.exe) não foi encontrado.") from exc
    finally:
        for path in (request_body_path, response_path):
            if path is not None:
                try:
                    path.unlink(missing_ok=True)
                except OSError:
                    pass


def _required_room(value: str) -> str:
    room = str(value or "").strip()
    if room not in ARGUS_ROOMS:
        raise ArgusApiError(
            "Sala da Argus ausente ou inválida. Selecione PADRÃO, SALA POLLY ou SALA CLEITON."
        )
    return room


def _as_int(value: object, default: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def _protect_windows_data(data: bytes) -> bytes:
    if sys.platform != "win32":
        raise ArgusApiError("O armazenamento protegido da chave do relay requer Windows.")
    return _windows_crypt(data, protect=True)


def _unprotect_windows_data(data: bytes) -> bytes:
    if sys.platform != "win32":
        raise ArgusApiError("O armazenamento protegido da chave do relay requer Windows.")
    return _windows_crypt(data, protect=False)


def _windows_crypt(data: bytes, *, protect: bool) -> bytes:
    buffer = ctypes.create_string_buffer(data)
    input_blob = _DataBlob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
    output_blob = _DataBlob()
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    crypt32.CryptProtectData.argtypes = (
        ctypes.POINTER(_DataBlob),
        wintypes.LPCWSTR,
        ctypes.POINTER(_DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DataBlob),
    )
    crypt32.CryptProtectData.restype = wintypes.BOOL
    crypt32.CryptUnprotectData.argtypes = (
        ctypes.POINTER(_DataBlob),
        ctypes.POINTER(wintypes.LPWSTR),
        ctypes.POINTER(_DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DataBlob),
    )
    crypt32.CryptUnprotectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p
    if protect:
        success = crypt32.CryptProtectData(
            ctypes.byref(input_blob),
            "Auto C6BANK Argus",
            None,
            None,
            None,
            0,
            ctypes.byref(output_blob),
        )
    else:
        success = crypt32.CryptUnprotectData(
            ctypes.byref(input_blob),
            None,
            None,
            None,
            None,
            0,
            ctypes.byref(output_blob),
        )
    if not success:
        raise ctypes.WinError()
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(output_blob.pbData, ctypes.c_void_p))
