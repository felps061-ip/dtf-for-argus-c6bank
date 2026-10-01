from __future__ import annotations

import base64
import ctypes
import hashlib
import json
import os
import subprocess
import sys
import tempfile
import unicodedata
import uuid
from ctypes import wintypes
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable
from urllib import error as urllib_error
from urllib import parse as urllib_parse
from urllib import request as urllib_request


ARGUS_BASE_URL = "https://argus.app.br/apiargus"
DEFAULT_ARGUS_RELAY_URL = "https://scagi.grupoimperiofinanceira.com.br/api/argus/upload"
DEFAULT_SKILL_NAME = "LAYOUT PADRÃO C6"
DEFAULT_SKILL_REFERENCE = "5ni7te02zzh1ka8"
DEFAULT_TOKEN_SIGNATURE = os.environ.get("ARGUS_TOKEN_SIGNATURE_DEFAULT", "").strip()
CONFIG_DIR_NAME = "Auto C6BANK - Datafast for Argus"


class ArgusApiError(RuntimeError):
    pass


@dataclass(frozen=True)
class ArgusConfig:
    token_signature: str
    skill_name: str = DEFAULT_SKILL_REFERENCE
    relay_url: str = ""
    relay_secret: str = ""


@dataclass(frozen=True)
class ArgusUploadResult:
    id_arquivo: int
    cod_status: int
    desc_status: str
    skill_hash: str


def list_skills(
    token_signature: str,
    *,
    base_url: str = ARGUS_BASE_URL,
    timeout: int = 60,
    opener: Callable | None = None,
) -> list[dict]:
    token = _required_token(token_signature)
    request = urllib_request.Request(
        f"{base_url.rstrip('/')}/cmd/skills",
        headers={"Token-Signature": token, "Accept": "application/json"},
        method="GET",
    )
    payload = _request_json(request, timeout=timeout, opener=opener)
    if _as_int(payload.get("codStatus"), default=1) < 1:
        raise ArgusApiError(str(payload.get("descStatus") or "Falha ao listar Skills da Argus."))

    items = payload.get("retornoGetSkillsItens") or payload.get("items") or []
    if not isinstance(items, list):
        raise ArgusApiError("A Argus retornou uma lista de Skills em formato inesperado.")
    return [{str(key).strip(): value for key, value in item.items()} for item in items if isinstance(item, dict)]


def resolve_skill_hash(skills: list[dict], skill_reference: str) -> str:
    reference = str(skill_reference or "").strip()
    if not reference:
        raise ArgusApiError("Informe o Skill da Argus.")

    for item in skills:
        skill_hash = str(item.get("hashEndpointSkill") or "").strip()
        if skill_hash and skill_hash == reference:
            return skill_hash

    normalized_reference = _normalize_text(reference)
    matches: list[str] = []
    for item in skills:
        values = (
            item.get("descricaoSkill"),
            item.get("descricaoCompleta"),
            item.get("descricaoCampanha"),
        )
        if normalized_reference not in {_normalize_text(value) for value in values if value is not None}:
            continue
        skill_hash = str(item.get("hashEndpointSkill") or "").strip()
        if skill_hash and skill_hash not in matches:
            matches.append(skill_hash)

    if not matches:
        raise ArgusApiError(f'Não encontrei o Skill "{reference}" na conta Argus.')
    if len(matches) > 1:
        raise ArgusApiError(
            f'O nome de Skill "{reference}" corresponde a mais de um hash. '
            "Informe o hashEndpointSkill exato nas configurações."
        )
    return matches[0]


def upload_mailing_csv(
    csv_path: str | Path,
    token_signature: str,
    skill_reference: str = DEFAULT_SKILL_REFERENCE,
    *,
    base_url: str = ARGUS_BASE_URL,
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
        token_signature,
        skill_reference,
        base_url=base_url,
        relay_url=relay_url,
        relay_secret=relay_secret,
        timeout=timeout,
        opener=opener,
    )


def upload_mailing_csv_bytes(
    csv_bytes: bytes,
    filename: str,
    token_signature: str,
    skill_reference: str = DEFAULT_SKILL_REFERENCE,
    *,
    base_url: str = ARGUS_BASE_URL,
    relay_url: str = "",
    relay_secret: str = "",
    timeout: int = 600,
    opener: Callable | None = None,
) -> ArgusUploadResult:
    token = _required_token(token_signature)
    if not csv_bytes:
        raise ValueError("O CSV da Argus está vazio.")
    if not str(filename).lower().endswith(".csv"):
        raise ValueError("O arquivo para upload precisa ter extensão .csv.")

    reference = str(skill_reference or "").strip()
    resolved_by_hash = _looks_like_skill_hash(reference)
    fallback_skill_name = DEFAULT_SKILL_NAME
    if _looks_like_skill_hash(reference):
        skill_hash = reference
    else:
        skills = list_skills(
            token,
            base_url=base_url,
            timeout=min(timeout, 60),
            opener=opener,
        )
        skill_hash = resolve_skill_hash(skills, reference)

    try:
        return _post_mailing_csv_bytes(
            csv_bytes,
            filename,
            token,
            skill_hash,
            base_url=base_url,
            timeout=timeout,
            opener=opener,
        )
    except ArgusApiError as exc:
        direct_error = exc
        if resolved_by_hash and _is_invalid_endpoint_error(exc):
            try:
                skills = list_skills(
                    token,
                    base_url=base_url,
                    timeout=min(timeout, 60),
                    opener=opener,
                )
                fallback_hash = resolve_skill_hash(skills, fallback_skill_name)
                if fallback_hash != skill_hash:
                    return _post_mailing_csv_bytes(
                        csv_bytes,
                        filename,
                        token,
                        fallback_hash,
                        base_url=base_url,
                        timeout=timeout,
                        opener=opener,
                    )
            except ArgusApiError as fallback_exc:
                direct_error = fallback_exc

        relay_result = _upload_mailing_csv_bytes_via_relay(
            csv_bytes,
            filename,
            skill_hash,
            relay_url=relay_url,
            relay_secret=relay_secret,
            timeout=timeout,
            opener=opener,
        )
        if relay_result is not None:
            return relay_result
        raise direct_error


def _upload_mailing_csv_bytes_via_relay(
    csv_bytes: bytes,
    filename: str,
    skill_hash: str,
    *,
    relay_url: str = "",
    relay_secret: str = "",
    timeout: int,
    opener: Callable | None,
) -> ArgusUploadResult | None:
    relay_secret = (relay_secret or os.environ.get("ARGUS_RELAY_SECRET", "")).strip()
    if not relay_secret:
        return None

    relay_url = (relay_url or os.environ.get("ARGUS_RELAY_URL", DEFAULT_ARGUS_RELAY_URL)).strip()
    if not relay_url:
        return None

    request = urllib_request.Request(
        relay_url,
        data=csv_bytes,
        headers={
            "Accept": "application/json",
            "Content-Type": "text/csv; charset=utf-8",
            "X-Argus-Relay-Secret": relay_secret,
            "X-Argus-Filename": filename,
            "X-Argus-Skill": skill_hash,
        },
        method="POST",
    )
    payload = _request_json(request, timeout=timeout, opener=opener)
    cod_status = _as_int(payload.get("codStatus"), default=-1)
    desc_status = str(payload.get("descStatus") or "").strip()
    if cod_status < 1:
        raise ArgusApiError(desc_status or "O relay do SCAGI recusou o upload do mailing.")
    return ArgusUploadResult(
        id_arquivo=_as_int(payload.get("idArquivo"), default=0),
        cod_status=cod_status,
        desc_status=desc_status or "Sucesso via SCAGI",
        skill_hash=str(payload.get("skillHash") or skill_hash),
    )


def _post_mailing_csv_bytes(
    csv_bytes: bytes,
    filename: str,
    token: str,
    skill_hash: str,
    *,
    base_url: str,
    timeout: int,
    opener: Callable | None,
) -> ArgusUploadResult:
    body, content_type = _multipart_file_body(csv_bytes, filename)
    encoded_hash = urllib_parse.quote(skill_hash, safe="")
    request = urllib_request.Request(
        f"{base_url.rstrip('/')}/{encoded_hash}/uploadmailing",
        data=body,
        headers={
            "Token-Signature": token,
            "Accept": "application/json",
            "Content-Type": content_type,
        },
        method="POST",
    )
    payload = _request_json(request, timeout=timeout, opener=opener)
    cod_status = _as_int(payload.get("codStatus"), default=-1)
    desc_status = str(payload.get("descStatus") or "").strip()
    if cod_status < 1:
        raise ArgusApiError(desc_status or "A Argus recusou o upload do mailing.")
    return ArgusUploadResult(
        id_arquivo=_as_int(payload.get("idArquivo"), default=0),
        cod_status=cod_status,
        desc_status=desc_status or "Sucesso",
        skill_hash=skill_hash,
    )


def load_argus_config() -> ArgusConfig | None:
    environment_token = os.environ.get("ARGUS_TOKEN_SIGNATURE", "").strip()
    environment_skill = os.environ.get("ARGUS_SKILL", DEFAULT_SKILL_REFERENCE).strip()
    environment_relay_url = os.environ.get("ARGUS_RELAY_URL", "").strip()
    environment_relay_secret = os.environ.get("ARGUS_RELAY_SECRET", "").strip()
    if environment_token:
        return ArgusConfig(
            environment_token,
            environment_skill or DEFAULT_SKILL_REFERENCE,
            environment_relay_url,
            environment_relay_secret,
        )

    path = argus_config_path()
    if not path.exists():
        if DEFAULT_TOKEN_SIGNATURE:
            return ArgusConfig(DEFAULT_TOKEN_SIGNATURE, DEFAULT_SKILL_REFERENCE)
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        protected_token = base64.b64decode(payload["protected_token"])
        token = _unprotect_windows_data(protected_token).decode("utf-8")
        relay_secret = ""
        if payload.get("protected_relay_secret"):
            protected_relay_secret = base64.b64decode(payload["protected_relay_secret"])
            relay_secret = _unprotect_windows_data(protected_relay_secret).decode("utf-8")
        skill_name = str(payload.get("skill_name") or DEFAULT_SKILL_REFERENCE).strip()
        relay_url = str(payload.get("relay_url") or "").strip()
        if skill_name == DEFAULT_SKILL_NAME:
            skill_name = DEFAULT_SKILL_REFERENCE
    except Exception as exc:
        raise ArgusApiError(
            "Não consegui abrir a configuração protegida da Argus. Configure o acesso novamente."
        ) from exc
    return ArgusConfig(token, skill_name, relay_url, relay_secret)


def save_argus_config(
    token_signature: str,
    skill_name: str = DEFAULT_SKILL_REFERENCE,
    relay_url: str = "",
    relay_secret: str = "",
) -> Path:
    token = _required_token(token_signature)
    skill = str(skill_name or DEFAULT_SKILL_REFERENCE).strip()
    if skill == DEFAULT_SKILL_NAME:
        skill = DEFAULT_SKILL_REFERENCE
    protected = _protect_windows_data(token.encode("utf-8"))
    protected_relay_secret = (
        _protect_windows_data(relay_secret.strip().encode("utf-8"))
        if relay_secret.strip()
        else b""
    )
    path = argus_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(
            {
                "protected_token": base64.b64encode(protected).decode("ascii"),
                "skill_name": skill,
                "relay_url": relay_url.strip(),
                "protected_relay_secret": (
                    base64.b64encode(protected_relay_secret).decode("ascii")
                    if protected_relay_secret
                    else ""
                ),
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
        if record.get("sha256") == digest and _as_int(record.get("cod_status"), 0) >= 1:
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


def _multipart_file_body(csv_bytes: bytes, filename: str) -> tuple[bytes, str]:
    boundary = f"----AutoC6Bank{uuid.uuid4().hex}"
    safe_filename = str(filename).replace('"', "'").replace("\r", "").replace("\n", "")
    prefix = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name=""; filename="{safe_filename}"\r\n'
        "Content-Type: text/csv\r\n\r\n"
    ).encode("utf-8")
    suffix = f"\r\n--{boundary}--\r\n".encode("ascii")
    return prefix + csv_bytes + suffix, f"multipart/form-data; boundary={boundary}"


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


def _is_invalid_endpoint_error(exc: BaseException) -> bool:
    normalized = _normalize_text(str(exc))
    return "endpoint" in normalized and ("invalid" in normalized or "invalido" in normalized)


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


def _required_token(value: str) -> str:
    token = str(value or "").strip()
    if not token:
        raise ArgusApiError("Informe a Token-Signature da Argus.")
    return token


def _normalize_text(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return "".join(char for char in text if not unicodedata.combining(char)).strip().casefold()


def _looks_like_skill_hash(value: str) -> bool:
    return bool(value and len(value) >= 12 and value.isascii() and value.isalnum())


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
        raise ArgusApiError("O armazenamento protegido da Token-Signature requer Windows.")
    return _windows_crypt(data, protect=True)


def _unprotect_windows_data(data: bytes) -> bytes:
    if sys.platform != "win32":
        raise ArgusApiError("O armazenamento protegido da Token-Signature requer Windows.")
    return _windows_crypt(data, protect=False)


def _windows_crypt(data: bytes, *, protect: bool) -> bytes:
    buffer = ctypes.create_string_buffer(data)
    input_blob = _DataBlob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
    output_blob = _DataBlob()
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
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
