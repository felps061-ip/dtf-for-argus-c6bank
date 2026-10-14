import argparse
import importlib.util
import os
import shutil
import subprocess
import sys
import traceback
import winreg
from pathlib import Path
from tkinter import BooleanVar, StringVar, Tk, filedialog, messagebox, ttk


APP_NAME = "Auto C6BANK - Datafast for Argus"
APP_EXE = f"{APP_NAME}.exe"
PUBLISHER = "Grupo Imperio"


def load_app_version() -> str:
    """Carrega a versão única compartilhada com o aplicativo principal."""
    locations = (
        Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)) / "app_version.py",
        Path(__file__).resolve().parents[1] / "c6_streamlit_app" / "app_version.py",
    )
    for location in locations:
        if not location.exists():
            continue
        spec = importlib.util.spec_from_file_location("auto_c6_app_version", location)
        if spec is None or spec.loader is None:
            continue
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return str(module.__version__)
    return "0.0.0"


APP_VERSION = load_app_version()


def default_install_dir() -> Path:
    local_appdata = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    return local_appdata / "Programs" / APP_NAME


def start_menu_dir() -> Path:
    appdata = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    return appdata / "Microsoft" / "Windows" / "Start Menu" / "Programs" / APP_NAME


def desktop_dir() -> Path:
    return Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Desktop"


def resource_path(*parts: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    candidate = base.joinpath(*parts)
    if candidate.exists():
        return candidate

    # Fallback for running installer.py from source during development.
    source_candidate = Path(__file__).resolve().parents[1] / "c6_executavel" / APP_EXE
    return source_candidate


def ps_quote(value: Path | str) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def create_shortcut(link_path: Path, target_path: Path, working_dir: Path, icon_path: Path | None = None) -> None:
    link_path.parent.mkdir(parents=True, exist_ok=True)
    icon = icon_path or target_path
    script = (
        "$shell = New-Object -ComObject WScript.Shell; "
        f"$shortcut = $shell.CreateShortcut({ps_quote(link_path)}); "
        f"$shortcut.TargetPath = {ps_quote(target_path)}; "
        f"$shortcut.WorkingDirectory = {ps_quote(working_dir)}; "
        f"$shortcut.IconLocation = {ps_quote(str(icon) + ',0')}; "
        "$shortcut.Save()"
    )
    subprocess.run(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        check=True,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )


def write_uninstaller(install_dir: Path) -> Path:
    uninstall_path = install_dir / f"Desinstalar {APP_NAME}.cmd"
    script = f"""@echo off
set "APP_DIR=%~dp0"
del "%USERPROFILE%\\Desktop\\{APP_NAME}.lnk" >nul 2>nul
rmdir /s /q "%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\{APP_NAME}" >nul 2>nul
reg delete "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{APP_NAME}" /f >nul 2>nul
cd /d "%TEMP%"
start "" /min powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Sleep -Milliseconds 700; Remove-Item -LiteralPath '%APP_DIR%' -Recurse -Force"
"""
    uninstall_path.write_text(script, encoding="utf-8")
    return uninstall_path


def register_uninstall(install_dir: Path, app_path: Path, uninstall_path: Path) -> None:
    key_path = rf"Software\Microsoft\Windows\CurrentVersion\Uninstall\{APP_NAME}"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
        winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, APP_NAME)
        winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, APP_VERSION)
        winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, PUBLISHER)
        winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, str(install_dir))
        winreg.SetValueEx(key, "DisplayIcon", 0, winreg.REG_SZ, str(app_path))
        winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, f'"{uninstall_path}"')
        winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)


def install_app(
    install_dir: Path,
    create_shortcuts: bool = True,
    register_in_windows: bool = True,
) -> Path:
    source_exe = resource_path("app", APP_EXE)
    if not source_exe.exists():
        raise FileNotFoundError(f"Executavel embutido nao encontrado: {source_exe}")

    install_dir.mkdir(parents=True, exist_ok=True)
    app_path = install_dir / APP_EXE
    shutil.copy2(source_exe, app_path)

    uninstall_path = write_uninstaller(install_dir)

    if create_shortcuts:
        create_shortcut(desktop_dir() / f"{APP_NAME}.lnk", app_path, install_dir, app_path)
        menu_dir = start_menu_dir()
        create_shortcut(menu_dir / f"{APP_NAME}.lnk", app_path, install_dir, app_path)
        create_shortcut(menu_dir / f"Desinstalar {APP_NAME}.lnk", uninstall_path, install_dir, app_path)

    if register_in_windows:
        register_uninstall(install_dir, app_path, uninstall_path)

    return app_path


class InstallerWindow:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title(f"Instalador - {APP_NAME}")
        self.root.resizable(False, False)

        self.install_dir = StringVar(value=str(default_install_dir()))
        self.shortcuts = BooleanVar(value=True)
        self.status = StringVar(value="Pronto para instalar.")

        frame = ttk.Frame(self.root, padding=18)
        frame.grid(row=0, column=0, sticky="nsew")

        title = ttk.Label(frame, text=APP_NAME, font=("Segoe UI", 14, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        text = ttk.Label(
            frame,
            text="Instala o programa neste usuario do Windows, sem precisar de administrador.",
            wraplength=500,
        )
        text.grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 16))

        ttk.Label(frame, text="Salvar em").grid(row=2, column=0, sticky="w")
        entry = ttk.Entry(frame, textvariable=self.install_dir, width=62)
        entry.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(4, 10))

        ttk.Button(frame, text="Procurar", command=self.choose_dir).grid(row=3, column=2, padx=(8, 0), pady=(4, 10))

        ttk.Checkbutton(frame, text="Criar atalhos na Area de Trabalho e Menu Iniciar", variable=self.shortcuts).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(0, 12)
        )

        ttk.Label(frame, textvariable=self.status).grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 12))

        buttons = ttk.Frame(frame)
        buttons.grid(row=6, column=0, columnspan=3, sticky="e")
        ttk.Button(buttons, text="Cancelar", command=self.root.destroy).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(buttons, text="Instalar", command=self.install).grid(row=0, column=1)

    def choose_dir(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.install_dir.get())
        if selected:
            self.install_dir.set(selected)

    def install(self) -> None:
        try:
            self.status.set("Instalando...")
            self.root.update_idletasks()
            app_path = install_app(Path(self.install_dir.get()), self.shortcuts.get(), True)
        except Exception:
            traceback_text = traceback.format_exc()
            self.status.set("Falha na instalacao.")
            messagebox.showerror("Erro ao instalar", traceback_text)
            return

        self.status.set("Instalado com sucesso.")
        open_now = messagebox.askyesno("Instalacao concluida", "Instalado com sucesso. Abrir o programa agora?")
        if open_now:
            subprocess.Popen([str(app_path)], cwd=str(app_path.parent))
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=f"Instalador do {APP_NAME}")
    parser.add_argument("--silent", action="store_true", help="Instala sem abrir a tela grafica.")
    parser.add_argument("--install-dir", default="", help="Pasta de instalacao.")
    parser.add_argument("--no-shortcuts", action="store_true", help="Nao cria atalhos.")
    parser.add_argument("--no-registry", action="store_true", help="Nao registra no Windows.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    install_dir = Path(args.install_dir) if args.install_dir else default_install_dir()

    if args.silent:
        try:
            app_path = install_app(install_dir, not args.no_shortcuts, not args.no_registry)
        except Exception as exc:
            print(f"Falha na instalacao: {exc}", file=sys.stderr)
            return 1
        print(f"Instalado em: {app_path}")
        return 0

    InstallerWindow().run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
