from __future__ import annotations

import hashlib
from pathlib import Path


def sha256_file(path: str | Path, chunk_size: int = 1024 * 1024) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(chunk_size), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_expected_sha256(path: str | Path) -> str:
    value = Path(path).read_text(encoding="utf-8-sig").strip().split()[0].lower()
    if len(value) != 64 or any(char not in "0123456789abcdef" for char in value):
        raise ValueError("Checksum SHA-256 invalido.")
    return value


def verify_sha256(path: str | Path, expected: str) -> bool:
    return sha256_file(path).casefold() == expected.strip().casefold()
