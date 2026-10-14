from __future__ import annotations

import json
from dataclasses import dataclass
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


REPOSITORY = "felps061-ip/dtf-for-argus-c6bank"
LATEST_RELEASE_URL = f"https://api.github.com/repos/{REPOSITORY}/releases/latest"
USER_AGENT = "AutoC6Bank-Updater/2.7"
EXE_ASSET = "AutoC6Bank.exe"
CHECKSUM_ASSET = "AutoC6Bank.exe.sha256"
UPDATER_ASSET = "AutoC6BankUpdater.exe"
UPDATER_CHECKSUM_ASSET = "AutoC6BankUpdater.exe.sha256"


class ReleaseLookupError(RuntimeError):
    """Falha tratavel ao consultar uma release publica."""


@dataclass(frozen=True)
class ReleaseInfo:
    tag_name: str
    version: str
    notes: str
    exe_url: str
    checksum_url: str
    updater_url: str
    updater_checksum_url: str = ""


def get_latest_release(*, timeout: int = 12, opener=urlopen) -> ReleaseInfo | None:
    request = Request(
        LATEST_RELEASE_URL,
        headers={"Accept": "application/vnd.github+json", "User-Agent": USER_AGENT},
    )
    try:
        with opener(request, timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except (HTTPError, URLError, TimeoutError, OSError, json.JSONDecodeError) as exc:
        raise ReleaseLookupError("Nao foi possivel consultar atualizacoes agora.") from exc

    if not isinstance(payload, dict) or payload.get("draft") or payload.get("prerelease"):
        return None
    tag = str(payload.get("tag_name") or "").strip()
    if not tag:
        return None
    assets = {
        str(asset.get("name")): str(asset.get("browser_download_url"))
        for asset in payload.get("assets", [])
        if isinstance(asset, dict)
    }
    required = (EXE_ASSET, CHECKSUM_ASSET, UPDATER_ASSET, UPDATER_CHECKSUM_ASSET)
    if not all(assets.get(name, "").startswith("https://") for name in required):
        return None
    return ReleaseInfo(
        tag_name=tag,
        version=tag.removeprefix("v").removeprefix("V"),
        notes=str(payload.get("body") or "").strip(),
        exe_url=assets[EXE_ASSET],
        checksum_url=assets[CHECKSUM_ASSET],
        updater_url=assets[UPDATER_ASSET],
        updater_checksum_url=assets[UPDATER_CHECKSUM_ASSET],
    )
