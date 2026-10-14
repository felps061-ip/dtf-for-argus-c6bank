from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path


def _wait_for_exit(pid: int, timeout: int = 45) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            os.kill(pid, 0)
        except OSError:
            return True
        time.sleep(0.4)
    return False


def install_update(target: Path, new: Path, backup: Path, pid: int) -> None:
    if not _wait_for_exit(pid):
        raise RuntimeError("O aplicativo anterior nao foi encerrado a tempo.")
    if not new.is_file():
        raise FileNotFoundError("Arquivo da atualizacao nao encontrado.")
    backup.unlink(missing_ok=True)
    backup_created = False
    try:
        if target.exists():
            target.replace(backup)
            backup_created = True
        new.replace(target)
        subprocess.Popen([str(target)], cwd=str(target.parent), close_fds=True)
    except Exception:
        if backup_created and backup.exists():
            target.unlink(missing_ok=True)
            backup.replace(target)
        raise


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pid", required=True, type=int)
    parser.add_argument("--target", required=True)
    parser.add_argument("--new", required=True)
    parser.add_argument("--backup", required=True)
    args = parser.parse_args()
    try:
        install_update(Path(args.target), Path(args.new), Path(args.backup), args.pid)
    except Exception as exc:
        print(f"Falha ao atualizar: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
