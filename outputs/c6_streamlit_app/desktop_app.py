from __future__ import annotations

import argparse
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import StringVar, Tk, filedialog, messagebox
from tkinter import ttk

from c6_processor import (
    DEFAULT_PHONE_BASE_DIR,
    FilterOptions,
    build_filtered_from_csv,
    get_missing_phone_cpfs,
    parse_date_value,
    parse_faltam_values,
    write_missing_phone_cpfs_csv,
    write_xlsx,
)

APP_NAME = "Auto C6BANK - Datafast for Argus"
FOOTER_TEXT = "v1.0 by: Equipe TI Grupo Império"


class C6DesktopApp:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title(APP_NAME)
        self.root.geometry("760x520")
        self.root.minsize(720, 480)

        self.csv_path = StringVar()
        self.lemit_path = StringVar()
        self.output_path = StringVar()
        self.faltam = StringVar()
        self.faltam_modo = StringVar(value="Maior que o prazo")
        self.taxa_min = StringVar(value="0,00")
        self.troco_min = StringVar(value="0,00")
        self.data_inicial = StringVar(value="01/01/1954")
        self.data_final = StringVar(value="31/12/1995")
        self.status = StringVar(value="Selecione a Extração DTF CSV.")

        self._build_ui()

    def run(self) -> None:
        self.root.mainloop()

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        frame = ttk.Frame(self.root, padding=18)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, weight=0)

        title = ttk.Label(frame, text=APP_NAME, font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        ttk.Label(frame, text="Extração DTF CSV").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=self.csv_path).grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Button(frame, text="Escolher", command=self._choose_csv).grid(
            row=1, column=2, padx=(8, 0), pady=5
        )

        filters = ttk.LabelFrame(frame, text="Filtros", padding=12)
        filters.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(14, 8))
        for column in range(4):
            filters.columnconfigure(column, weight=1)

        self._add_labeled_entry(filters, "FALTAM", self.faltam, 0, 0)
        self._add_labeled_combo(
            filters,
            "Regra FALTAM",
            self.faltam_modo,
            ["Maior que o prazo", "Ate o prazo", "Exato"],
            0,
            1,
        )
        self._add_labeled_entry(filters, "TAXA minima", self.taxa_min, 0, 2)
        self._add_labeled_entry(filters, "TROCO maior que", self.troco_min, 0, 3)
        self._add_labeled_entry(filters, "Data inicial", self.data_inicial, 1, 0)
        self._add_labeled_entry(filters, "Data final", self.data_final, 1, 1)

        hint = ttk.Label(
            filters,
            text="FALTAM pode ficar vazio ou receber varios prazos, como 67,72.",
            foreground="#555555",
        )
        hint.grid(row=2, column=0, columnspan=4, sticky="w", pady=(8, 0))

        notebook = ttk.Notebook(frame)
        notebook.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(12, 8))

        cpfs_tab = ttk.Frame(notebook, padding=12)
        xlsx_tab = ttk.Frame(notebook, padding=12)
        notebook.add(cpfs_tab, text="1. CPFs para Lemiti")
        notebook.add(xlsx_tab, text="2. Retorno Lemiti e XLSX")

        cpfs_tab.columnconfigure(0, weight=1)
        self.missing_button = ttk.Button(
            cpfs_tab,
            text="Extrair CPFs para Lemiti (.csv)",
            command=self._start_export_missing_cpfs,
        )
        self.missing_button.grid(row=0, column=0, sticky="w")

        xlsx_tab.columnconfigure(1, weight=1)
        ttk.Label(xlsx_tab, text="Retorno Lemiti").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(xlsx_tab, textvariable=self.lemit_path).grid(
            row=0, column=1, sticky="ew", pady=5
        )
        ttk.Button(xlsx_tab, text="Escolher", command=self._choose_lemit).grid(
            row=0, column=2, padx=(8, 0), pady=5
        )

        ttk.Label(xlsx_tab, text="Salvar em").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(xlsx_tab, textvariable=self.output_path).grid(
            row=1, column=1, sticky="ew", pady=5
        )
        ttk.Button(xlsx_tab, text="Escolher", command=self._choose_output).grid(
            row=1, column=2, padx=(8, 0), pady=5
        )

        self.generate_button = ttk.Button(xlsx_tab, text="Gerar XLSX", command=self._start_generate)
        self.generate_button.grid(row=2, column=2, sticky="e", pady=(10, 0))

        ttk.Label(frame, textvariable=self.status, foreground="#1F4E78").grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )
        ttk.Separator(frame).grid(row=5, column=0, columnspan=3, sticky="ew", pady=(14, 6))
        ttk.Label(frame, text=FOOTER_TEXT, foreground="#666666").grid(
            row=6, column=0, columnspan=3, sticky="e"
        )

    def _add_labeled_entry(
        self, parent: ttk.Frame, label: str, variable: StringVar, row: int, column: int
    ) -> None:
        box = ttk.Frame(parent)
        box.grid(row=row, column=column, sticky="ew", padx=(0, 14), pady=4)
        box.columnconfigure(0, weight=1)
        ttk.Label(box, text=label).grid(row=0, column=0, sticky="w")
        ttk.Entry(box, textvariable=variable).grid(row=1, column=0, sticky="ew", pady=(3, 0))

    def _add_labeled_combo(
        self,
        parent: ttk.Frame,
        label: str,
        variable: StringVar,
        values: list[str],
        row: int,
        column: int,
    ) -> None:
        box = ttk.Frame(parent)
        box.grid(row=row, column=column, sticky="ew", padx=(0, 14), pady=4)
        box.columnconfigure(0, weight=1)
        ttk.Label(box, text=label).grid(row=0, column=0, sticky="w")
        ttk.Combobox(box, textvariable=variable, values=values, state="readonly").grid(
            row=1, column=0, sticky="ew", pady=(3, 0)
        )

    def _choose_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="Escolha a Extração DTF CSV",
            filetypes=[("CSV", "*.csv"), ("Todos os arquivos", "*.*")],
        )
        if not path:
            return

        self.csv_path.set(path)
        if not self.output_path.get().strip():
            source = Path(path)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.output_path.set(str(source.with_name(f"auto_c6bank_{stamp}.xlsx")))

    def _choose_lemit(self) -> None:
        path = filedialog.askopenfilename(
            title="Escolha o retorno Lemiti",
            filetypes=[
                ("Arquivos Lemiti", "*.csv *.xlsx *.xls"),
                ("CSV", "*.csv"),
                ("Excel", "*.xlsx *.xls"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if path:
            self.lemit_path.set(path)

    def _choose_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Salvar em",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if path:
            self.output_path.set(path)

    def _start_generate(self) -> None:
        self._set_busy(True)
        self.status.set("Processando e buscando telefones...")
        thread = threading.Thread(target=self._generate, daemon=True)
        thread.start()

    def _start_export_missing_cpfs(self) -> None:
        try:
            csv_path = self._get_csv_path()
        except ValueError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        source = Path(csv_path)
        output_path = filedialog.asksaveasfilename(
            title="Salvar CPFs sem telefone",
            initialdir=str(source.parent),
            initialfile=f"cpfs_sem_telefone_{stamp}.csv",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not output_path:
            return

        self._set_busy(True)
        self.status.set("Buscando telefones e separando CPFs sem telefone...")
        thread = threading.Thread(
            target=self._export_missing_cpfs, args=(output_path,), daemon=True
        )
        thread.start()

    def _generate(self) -> None:
        try:
            csv_path = self._get_csv_path()
            output_path = self.output_path.get().strip()
            if not output_path:
                raise ValueError("Escolha onde salvar o XLSX.")
            lemit_path = self.lemit_path.get().strip()
            if not lemit_path:
                raise ValueError("Escolha o retorno Lemiti antes de gerar o XLSX.")

            transformed, filtered = build_filtered_from_csv(
                csv_path, self._get_options(), DEFAULT_PHONE_BASE_DIR, lemit_path
            )
            write_xlsx(filtered, output_path)
            message = (
                f"Concluido: {len(filtered)} de {len(transformed)} linhas exportadas.\n"
                f"Arquivo: {output_path}"
            )
            self.root.after(0, self._generation_success, message)
        except Exception as exc:
            self.root.after(0, self._generation_error, str(exc))

    def _export_missing_cpfs(self, output_path: str) -> None:
        try:
            transformed, filtered = build_filtered_from_csv(
                self._get_csv_path(), self._get_options(), DEFAULT_PHONE_BASE_DIR
            )
            missing = get_missing_phone_cpfs(filtered)
            write_missing_phone_cpfs_csv(filtered, output_path)
            message = (
                f"Concluido: {len(missing)} CPFs sem telefone exportados.\n"
                f"Base filtrada: {len(filtered)} de {len(transformed)} linhas.\n"
                f"Arquivo: {output_path}"
            )
            self.root.after(0, self._generation_success, message)
        except Exception as exc:
            self.root.after(0, self._generation_error, str(exc))

    def _get_csv_path(self) -> str:
        csv_path = self.csv_path.get().strip()
        if not csv_path:
            raise ValueError("Escolha a Extração DTF CSV.")
        return csv_path

    def _get_options(self) -> FilterOptions:
        return FilterOptions(
            faltam=parse_faltam_values(self.faltam.get()),
            faltam_modo=_faltam_mode_code(self.faltam_modo.get()),
            taxa_min=_parse_decimal(self.taxa_min.get()),
            troco_min=_parse_decimal(self.troco_min.get()),
            data_inicial=parse_date_value(self.data_inicial.get()),
            data_final=parse_date_value(self.data_final.get()),
        )

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.generate_button.configure(state=state)
        self.missing_button.configure(state=state)

    def _generation_success(self, message: str) -> None:
        self._set_busy(False)
        self.status.set(message.replace("\n", " "))
        messagebox.showinfo(APP_NAME, message)

    def _generation_error(self, message: str) -> None:
        self._set_busy(False)
        self.status.set("Erro ao gerar XLSX.")
        messagebox.showerror(APP_NAME, message)


def _parse_decimal(value: str) -> float:
    text = str(value).strip()
    if not text:
        return 0.0
    return float(text.replace(".", "").replace(",", ".") if "," in text else text)


def _faltam_mode_code(label: str) -> str:
    if label == "Exato":
        return "exato"
    if label == "Ate o prazo":
        return "ate"
    return "maior"


def main() -> None:
    if "--smoke-test" in sys.argv:
        _run_smoke_test()
        return

    C6DesktopApp().run()


def _run_smoke_test() -> None:
    parser = argparse.ArgumentParser(description="Smoke test do executavel C6.")
    parser.add_argument("--smoke-test", action="store_true")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--lemit-input", default="")
    parser.add_argument("--faltam", default="")
    parser.add_argument(
        "--faltam-modo",
        default="maior",
        choices=["maior", "a_partir", "ate", "exato"],
    )
    args = parser.parse_args()

    options = FilterOptions(
        faltam=parse_faltam_values(args.faltam),
        faltam_modo=args.faltam_modo,
        taxa_min=0.0,
        troco_min=0.0,
        data_inicial=parse_date_value("01/01/1954"),
        data_final=parse_date_value("31/12/1995"),
    )
    lemit_input = args.lemit_input or None
    transformed, filtered = build_filtered_from_csv(
        args.input, options, DEFAULT_PHONE_BASE_DIR, lemit_input
    )
    write_xlsx(filtered, args.output)
    print(f"OK {len(filtered)} de {len(transformed)} linhas")


if __name__ == "__main__":
    main()
