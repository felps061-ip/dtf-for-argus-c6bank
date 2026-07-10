from __future__ import annotations

import argparse
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import StringVar, Tk, filedialog, messagebox
from tkinter import ttk

from c6_processor import (
    DEFAULT_PHONE_BASE_DIRS,
    FilterOptions,
    GoogleSheetsSetupError,
    build_argus_campaign_name,
    build_argus_output_path,
    build_filtered_from_csv,
    build_lemit_processed_xlsx_path,
    build_sms_enviados_output_path,
    build_sms_enviados_output_filename,
    build_sms_output_path,
    create_sms_enviados_google_sheet,
    get_missing_phone_cpfs,
    parse_date_value,
    parse_faltam_values,
    write_argus_discadora_csv,
    write_missing_phone_cpfs_csv,
    write_sms_enviados_xlsx,
    write_sms_xlsx,
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
        self.argus_output_path = StringVar()
        self.argus_campaign = StringVar()
        self.faltam = StringVar()
        self.faltam_modo = StringVar(value="Maior que o prazo")
        self.taxa_min = StringVar(value="0,00")
        self.troco_min = StringVar(value="0,00")
        self.data_inicial = StringVar(value="01/01/1954")
        self.data_final = StringVar(value="31/12/1995")
        self.telefone_data_inicial = StringVar(value="01/01/2026")
        self.status = StringVar(value="Selecione a Extração DTF CSV.")

        self._build_ui()
        self.root.bind("<F5>", self._reset_form)

    def run(self) -> None:
        self.root.mainloop()

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        frame = ttk.Frame(self.root, padding=18)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, weight=0)

        title = ttk.Label(frame, text=APP_NAME, font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))
        self.reset_button = ttk.Button(frame, text="Reiniciar / F5", command=self._reset_form)
        self.reset_button.grid(row=0, column=2, sticky="e", pady=(0, 12))

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
        self._add_labeled_entry(filters, "Periodo Telefone", self.telefone_data_inicial, 1, 2)
        hint = ttk.Label(
            filters,
            text="FALTAM pode ficar vazio ou receber varios prazos, como 67,72.",
            foreground="#555555",
        )
        hint.grid(row=2, column=0, columnspan=4, sticky="w", pady=(8, 0))

        self.notebook = ttk.Notebook(frame)
        self.notebook.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(12, 8))

        cpfs_tab = ttk.Frame(self.notebook, padding=12)
        xlsx_tab = ttk.Frame(self.notebook, padding=12)
        sms_tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(cpfs_tab, text="1. CPFs para Lemiti")
        self.notebook.add(xlsx_tab, text="2. Retorno Lemiti, XLSX e Argus")
        self.notebook.add(sms_tab, text="3. SMS")

        cpfs_tab.columnconfigure(0, weight=1)
        self.missing_button = ttk.Button(
            cpfs_tab,
            text="Extrair CPFs para Lemiti (.csv)",
            command=self._start_export_missing_cpfs,
        )
        self.missing_button.grid(row=0, column=0, sticky="w")
        self.audit_xlsx_button = ttk.Button(
            cpfs_tab,
            text="Gerar planilha de conferencia (.xlsx)",
            command=self._start_export_audit_xlsx,
        )
        self.audit_xlsx_button.grid(row=1, column=0, sticky="w", pady=(10, 0))

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

        ttk.Label(xlsx_tab, text="Campanha Argus").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(xlsx_tab, textvariable=self.argus_campaign).grid(
            row=2, column=1, sticky="ew", pady=5
        )

        ttk.Label(xlsx_tab, text="Salvar CSV Argus").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Entry(xlsx_tab, textvariable=self.argus_output_path).grid(
            row=3, column=1, sticky="ew", pady=5
        )
        ttk.Button(xlsx_tab, text="Escolher", command=self._choose_argus_output).grid(
            row=3, column=2, padx=(8, 0), pady=5
        )

        ttk.Label(
            xlsx_tab,
            text="O CSV da Argus e gerado depois do XLSX usando a UF e o nome do XLSX final.",
            foreground="#555555",
            wraplength=650,
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))

        self.generate_button = ttk.Button(
            xlsx_tab, text="Gerar XLSX e CSV Argus", command=self._start_generate
        )
        self.generate_button.grid(row=5, column=2, sticky="e", pady=(14, 0))

        sms_tab.columnconfigure(0, weight=1)
        ttk.Label(
            sms_tab,
            text="Gera o XLSX para SMS usando a mesma base filtrada e salva na pasta mensal do Disparo Pro.",
            foreground="#555555",
            wraplength=650,
        ).grid(row=0, column=0, sticky="w")
        self.sms_button = ttk.Button(
            sms_tab,
            text="Gerar XLSX SMS",
            command=self._start_export_sms_xlsx,
        )
        self.sms_button.grid(row=1, column=0, sticky="w", pady=(12, 0))

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

    def _reset_form(self, event: object | None = None) -> None:
        self.csv_path.set("")
        self.lemit_path.set("")
        self.output_path.set("")
        self.argus_output_path.set("")
        self.argus_campaign.set("")
        self.faltam.set("")
        self.faltam_modo.set("Maior que o prazo")
        self.taxa_min.set("0,00")
        self.troco_min.set("0,00")
        self.data_inicial.set("01/01/1954")
        self.data_final.set("31/12/1995")
        self.telefone_data_inicial.set("01/01/2026")
        self.status.set("Tela reiniciada. Selecione a Extracao DTF CSV.")
        self.notebook.select(0)

    def _choose_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="Escolha a Extração DTF CSV",
            filetypes=[("CSV", "*.csv"), ("Todos os arquivos", "*.*")],
        )
        if not path:
            return

        self.csv_path.set(path)
        self._suggest_argus_campaign()

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
            self._suggest_argus_campaign()

    def _suggest_argus_campaign(self) -> None:
        try:
            campanha = build_argus_campaign_name(self._get_csv_path())
            self.argus_campaign.set(campanha)
        except Exception as exc:
            self.argus_campaign.set("")
            messagebox.showerror(APP_NAME, str(exc))

    def _choose_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Salvar em",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if path:
            self.output_path.set(path)

    def _choose_argus_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Salvar CSV Discadora Argus",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if path:
            self.argus_output_path.set(path)

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

    def _start_export_audit_xlsx(self) -> None:
        try:
            csv_path = self._get_csv_path()
        except ValueError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        source = Path(csv_path)
        output_path = filedialog.asksaveasfilename(
            title="Salvar planilha de conferencia",
            initialdir=str(source.parent),
            initialfile=f"conferencia_c6_telefones_{stamp}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not output_path:
            return

        self._set_busy(True)
        self.status.set("Gerando planilha de conferencia com telefones da base...")
        thread = threading.Thread(
            target=self._export_audit_xlsx, args=(output_path,), daemon=True
        )
        thread.start()

    def _start_export_sms_xlsx(self) -> None:
        try:
            csv_path = self._get_csv_path()
        except ValueError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return

        lemit_path = self.lemit_path.get().strip()
        if not lemit_path:
            messagebox.showerror(APP_NAME, "Escolha o retorno Lemiti antes de gerar o SMS.")
            return

        self._set_busy(True)
        self.status.set("Gerando XLSX de SMS na pasta do Disparo Pro...")
        thread = threading.Thread(target=self._export_sms_xlsx, daemon=True)
        thread.start()

    def _generate(self) -> None:
        try:
            csv_path = self._get_csv_path()
            output_path = self.output_path.get().strip()
            argus_output_path = self.argus_output_path.get().strip()
            lemit_path = self.lemit_path.get().strip()
            if not lemit_path:
                raise ValueError("Escolha o retorno Lemiti antes de gerar o XLSX.")

            transformed, filtered = build_filtered_from_csv(
                csv_path, self._get_options(), DEFAULT_PHONE_BASE_DIRS, lemit_path
            )
            if not output_path:
                output_path = str(build_lemit_processed_xlsx_path(csv_path, len(filtered)))
                self.output_path.set(output_path)
            campanha = build_argus_campaign_name(output_path)
            self.argus_campaign.set(campanha)
            if not argus_output_path:
                argus_output_path = str(build_argus_output_path(output_path))
                self.argus_output_path.set(argus_output_path)
            write_xlsx(filtered, output_path)
            write_argus_discadora_csv(filtered, argus_output_path, campanha)
            message = (
                f"Concluido: {len(filtered)} de {len(transformed)} linhas exportadas.\n"
                f"XLSX: {output_path}\n"
                f"CSV Argus: {argus_output_path}"
            )
            self.root.after(0, self._generation_success, message)
        except Exception as exc:
            self.root.after(0, self._generation_error, str(exc))

    def _export_sms_xlsx(self) -> None:
        try:
            csv_path = self._get_csv_path()
            lemit_path = self.lemit_path.get().strip()
            transformed, filtered = build_filtered_from_csv(
                csv_path, self._get_options(), DEFAULT_PHONE_BASE_DIRS, lemit_path
            )
            output_path = build_sms_output_path(csv_path, len(filtered))
            write_sms_xlsx(filtered, output_path)
            enviados_title = build_sms_enviados_output_filename(csv_path, len(filtered)).removesuffix(".xlsx")
            try:
                enviados_result = create_sms_enviados_google_sheet(filtered, enviados_title)
            except Exception as exc:
                enviados_path = build_sms_enviados_output_path(csv_path, len(filtered))
                write_sms_enviados_xlsx(filtered, enviados_path)
                prefix = "Google Sheets" if isinstance(exc, GoogleSheetsSetupError) else "Falha Google Sheets"
                enviados_result = f"{enviados_path}\n{prefix}: {exc}"
            message = (
                f"Concluido: SMS gerado com base em {len(filtered)} de {len(transformed)} linhas.\n"
                f"SMS: {output_path}\n"
                f"Auto-enviados: {enviados_result}"
            )
            self.root.after(0, self._generation_success, message)
        except Exception as exc:
            self.root.after(0, self._generation_error, str(exc))

    def _export_missing_cpfs(self, output_path: str) -> None:
        try:
            transformed, filtered = build_filtered_from_csv(
                self._get_csv_path(), self._get_options(), DEFAULT_PHONE_BASE_DIRS
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

    def _export_audit_xlsx(self, output_path: str) -> None:
        try:
            transformed, filtered = build_filtered_from_csv(
                self._get_csv_path(), self._get_options(), DEFAULT_PHONE_BASE_DIRS
            )
            write_xlsx(filtered, output_path)
            missing = get_missing_phone_cpfs(filtered)
            message = (
                f"Concluido: {len(filtered)} de {len(transformed)} linhas na conferencia.\n"
                f"Sem telefone base: {len(missing)}.\n"
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
            telefone_data_inicial=parse_date_value(self.telefone_data_inicial.get()),
        )

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.reset_button.configure(state=state)
        self.generate_button.configure(state=state)
        self.missing_button.configure(state=state)
        self.audit_xlsx_button.configure(state=state)
        self.sms_button.configure(state=state)

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
        args.input, options, DEFAULT_PHONE_BASE_DIRS, lemit_input
    )
    write_xlsx(filtered, args.output)
    print(f"OK {len(filtered)} de {len(transformed)} linhas")


if __name__ == "__main__":
    main()
