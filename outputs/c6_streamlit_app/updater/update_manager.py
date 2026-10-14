from __future__ import annotations

import logging
import os
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from urllib.request import Request, urlopen

from packaging.version import Version

from app_version import __version__
from .checksum import read_expected_sha256, verify_sha256
from .github_release import USER_AGENT, ReleaseInfo, get_latest_release


class UpdateState(str, Enum):
    IDLE = "IDLE"
    CHECKING = "CHECKING"
    AVAILABLE = "AVAILABLE"
    DOWNLOADING = "DOWNLOADING"
    VERIFYING = "VERIFYING"
    READY = "READY"
    INSTALLING = "INSTALLING"
    ERROR = "ERROR"


@dataclass(frozen=True)
class UpdateCheck:
    state: UpdateState
    release: ReleaseInfo | None = None
    message: str = ""


def logger() -> logging.Logger:
    log = logging.getLogger("auto_c6bank.updater")
    if log.handlers:
        return log
    folder = update_root().parent / "logs"
    folder.mkdir(parents=True, exist_ok=True)
    handler = logging.FileHandler(folder / "updater.log", encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    log.addHandler(handler)
    log.setLevel(logging.INFO)
    return log


def update_root() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    return base / "GrupoImperio" / "AutoC6Bank" / "updates"


def check_for_update(*, current_version: str = __version__) -> UpdateCheck:
    try:
        release = get_latest_release()
        if release is None or Version(release.version) <= Version(current_version):
            return UpdateCheck(UpdateState.IDLE)
        logger().info("Atualizacao disponivel: %s -> %s", current_version, release.version)
        return UpdateCheck(UpdateState.AVAILABLE, release)
    except Exception as exc:
        logger().warning("Verificacao de atualizacao falhou: %s", exc)
        return UpdateCheck(UpdateState.ERROR, message="Nao foi possivel verificar atualizacoes.")


def _download(url: str, destination: Path, *, progress=None, timeout: int = 45) -> None:
    if not url.startswith("https://"):
        raise ValueError("A atualizacao precisa usar HTTPS.")
    request = Request(url, headers={"User-Agent": USER_AGENT})
    with urlopen(request, timeout=timeout) as response, destination.open("wb") as output:
        total = int(response.headers.get("Content-Length") or 0)
        received = 0
        while True:
            chunk = response.read(1024 * 256)
            if not chunk:
                break
            output.write(chunk)
            received += len(chunk)
            if progress:
                progress(received, total)


def download_verified_update(release: ReleaseInfo, *, progress=None) -> tuple[Path, Path]:
    folder = update_root()
    folder.mkdir(parents=True, exist_ok=True)
    exe_part = folder / "AutoC6Bank.exe.part"
    exe_new = folder / "AutoC6Bank.new.exe"
    checksum = folder / "AutoC6Bank.exe.sha256"
    updater = folder / "AutoC6BankUpdater.exe"
    updater_checksum = folder / "AutoC6BankUpdater.exe.sha256"
    updater_part = folder / "AutoC6BankUpdater.exe.part"
    for path in (exe_part, exe_new, checksum, updater, updater_part, updater_checksum):
        path.unlink(missing_ok=True)
    try:
        _download(release.exe_url, exe_part, progress=progress)
        _download(release.checksum_url, checksum)
        expected = read_expected_sha256(checksum)
        if not verify_sha256(exe_part, expected):
            raise ValueError("Checksum SHA-256 da atualizacao nao confere.")
        exe_part.replace(exe_new)
        _download(release.updater_url, updater_part)
        _download(release.updater_checksum_url, updater_checksum)
        updater_expected = read_expected_sha256(updater_checksum)
        if not verify_sha256(updater_part, updater_expected):
            raise ValueError("Checksum SHA-256 do atualizador nao confere.")
        updater_part.replace(updater)
        logger().info("Download e checksum concluidos para %s", release.version)
        return exe_new, updater
    except Exception:
        exe_part.unlink(missing_ok=True)
        exe_new.unlink(missing_ok=True)
        updater_part.unlink(missing_ok=True)
        updater.unlink(missing_ok=True)
        raise


def current_executable() -> Path:
    return Path(sys.executable if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1] / "desktop_app.py")


def can_replace(path: Path) -> tuple[bool, str]:
    parent = path.parent
    try:
        with tempfile.NamedTemporaryFile(prefix=".update-write-", dir=parent, delete=True):
            pass
    except OSError:
        return False, "Nao foi possivel atualizar porque o aplicativo esta em uma pasta protegida pelo Windows."
    return True, ""


def launch_updater(new_exe: Path, updater_exe: Path) -> None:
    target = current_executable()
    allowed, message = can_replace(target)
    if not allowed:
        raise PermissionError(message)
    backup = target.with_suffix(".old.exe")
    subprocess.Popen(
        [str(updater_exe), "--pid", str(os.getpid()), "--target", str(target), "--new", str(new_exe), "--backup", str(backup)],
        cwd=str(target.parent),
        close_fds=True,
    )
