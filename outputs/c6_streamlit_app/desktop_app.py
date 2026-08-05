from __future__ import annotations

import argparse
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import PhotoImage, StringVar, Tk, filedialog, messagebox
from tkinter import ttk

from argus_api import (
    DEFAULT_SKILL_REFERENCE,
    file_was_uploaded,
    load_argus_config,
    record_upload,
    upload_mailing_csv,
)

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
    extract_uf_from_filename,
    format_bank_info,
    get_missing_phone_cpfs,
    identify_bank_from_filename,
    parse_date_value,
    parse_faltam_values,
    read_c6_csv,
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
        icon_path = Path(__file__).with_name("app_icon.ico")
        if icon_path.is_file():
            self.root.iconbitmap(default=str(icon_path))
        self.root.geometry("900x650")
        self.root.minsize(860, 620)
        self.brand_images = self._load_brand_images()

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
        self.bank_info = StringVar(value='Banco identificado: selecione um arquivo com "C6" ou "Digio" no nome.')
        self.status = StringVar(value="Selecione a Extração DTF CSV.")
        self.file_summary = StringVar(value="Nenhum arquivo selecionado")
        self.bank_summary = StringVar(value="-")
        self.uf_summary = StringVar(value="-")
        self.coef_summary = StringVar(value="-")
        self.records_summary = StringVar(value="-")
        self.step_summary = StringVar(value="Etapa atual: Selecionando Base")
        self.progress_summary = StringVar(value="Aguardando arquivo")
        self.result_summary = StringVar(value="")
        self.source_ready = False
        self.last_argus_csv_path: Path | None = None

        self._build_ui()
        self._set_source_ready(False)
        self.root.bind("<F5>", self._reset_form)

    def run(self) -> None:
        self.root.mainloop()

    def _load_brand_images(self) -> dict[str, PhotoImage]:
        images: dict[str, PhotoImage] = {}
        for key, filename, subsample in (
            ("imperio", "brand_imperio.png", 34),
            ("c6", "brand_c6.png", 11),
        ):
            path = Path(__file__).with_name(filename)
            if not path.is_file():
                continue
            try:
                image = PhotoImage(file=str(path))
                images[key] = image.subsample(subsample, subsample)
            except Exception:
                continue
        return images

    def _build_ui(self) -> None:
        self._build_minimal_ui()
        return

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
        ttk.Label(filters, textvariable=self.bank_info, foreground="#1F4E78").grid(
            row=3, column=0, columnspan=4, sticky="w", pady=(6, 0)
        )

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

        ttk.Label(xlsx_tab, text="Acesso Argus usa configuração local segura.").grid(
            row=5, column=0, sticky="w", pady=(14, 0)
        )
        self.argus_upload_button = ttk.Button(
            xlsx_tab,
            text="Enviar este CSV para a Argus",
            command=self._start_argus_upload,
            state="disabled",
        )
        self.argus_upload_button.grid(row=5, column=1, sticky="e", pady=(14, 0), padx=(8, 8))
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

    def _build_minimal_ui(self) -> None:
        self.root.configure(bg="#F7F7F5")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self._configure_minimal_styles()

        frame = ttk.Frame(self.root, padding=(22, 16), style="Minimal.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(4, weight=1)

        header = ttk.Frame(frame, style="Minimal.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.columnconfigure(0, weight=1)
        title_wrap = ttk.Frame(header, style="Minimal.TFrame")
        title_wrap.grid(row=0, column=0, rowspan=2, sticky="w")
        title_wrap.columnconfigure(1, weight=1)
        if "imperio" in self.brand_images:
            logo_tile = ttk.Frame(title_wrap, padding=4, style="LogoTile.TFrame")
            logo_tile.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 10))
            ttk.Label(logo_tile, image=self.brand_images["imperio"], style="LogoTile.TLabel").grid(
                row=0, column=0
            )
        ttk.Label(title_wrap, text="Auto Datafast para Argus", style="MinimalTitle.TLabel").grid(
            row=0, column=1, sticky="w"
        )
        ttk.Label(
            title_wrap,
            text="Prepare bases C6 e Digio automaticamente.",
            style="MinimalSub.TLabel",
        ).grid(row=1, column=1, sticky="w", pady=(2, 0))
        self.brand_badge = ttk.Label(header, text="", style="BrandNeutral.TLabel")
        self.brand_badge.grid(row=0, column=1, rowspan=2, sticky="e", padx=(0, 12))
        self.reset_button = ttk.Button(
            header,
            text="Limpar Projeto",
            command=self._reset_form,
            style="Quiet.TButton",
        )
        self.reset_button.grid(row=0, column=2, rowspan=2, sticky="e")

        ttk.Label(frame, textvariable=self.step_summary, style="Step.TLabel").grid(
            row=1, column=0, sticky="w", pady=(0, 8)
        )
        self.base_summary_label = ttk.Label(frame, textvariable=self.bank_info, style="Success.TLabel")
        self.base_summary_label.grid(row=2, column=0, sticky="w", pady=(0, 8))
        self.base_summary_label.grid_remove()

        source = ttk.Frame(frame, padding=(0, 6), style="Minimal.TFrame")
        source.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        source.columnconfigure(1, weight=1)
        ttk.Label(source, text="Selecionar Base", style="BlockTitle.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 6)
        )
        ttk.Entry(source, textvariable=self.csv_path).grid(row=1, column=0, columnspan=2, sticky="ew")
        ttk.Button(source, text="Selecionar Base", command=self._choose_csv, style="Quiet.TButton").grid(
            row=1, column=2, sticky="e", padx=(10, 0)
        )

        filters = ttk.Frame(frame, padding=(0, 6), style="Minimal.TFrame")
        filters.grid(row=4, column=0, sticky="ew", pady=(0, 8))
        filters.columnconfigure(0, weight=1)
        filters.columnconfigure(1, weight=1)
        ttk.Label(filters, text="Filtros", style="BlockTitle.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 6)
        )
        self._add_labeled_combo(
            filters,
            "Regra dos prazos",
            self.faltam_modo,
            ["Maior que o prazo", "Ate o prazo", "Exato"],
            1,
            0,
        )
        self._add_labeled_entry(filters, "Prazos FALTAM", self.faltam, 1, 1)
        self._add_labeled_entry(filters, "Taxa minima", self.taxa_min, 2, 0)
        self._add_labeled_entry(filters, "Troco maior que", self.troco_min, 2, 1)
        self._add_labeled_entry(filters, "Nascimento inicial", self.data_inicial, 3, 0)
        self._add_labeled_entry(filters, "Nascimento final", self.data_final, 3, 1)
        self._add_labeled_entry(filters, "Periodo telefone", self.telefone_data_inicial, 4, 0)

        self.notebook = ttk.Notebook(frame)
        self.notebook.grid(row=5, column=0, sticky="nsew", pady=(4, 0))

        cpfs_tab = ttk.Frame(self.notebook, padding=(14, 12), style="Minimal.TFrame")
        xlsx_tab = ttk.Frame(self.notebook, padding=(14, 12), style="Minimal.TFrame")
        sms_tab = ttk.Frame(self.notebook, padding=(14, 12), style="Minimal.TFrame")
        self.notebook.add(cpfs_tab, text="Preparar para Lemiti")
        self.notebook.add(xlsx_tab, text="Finalizar Base e Argus")
        self.notebook.add(sms_tab, text="Gerar SMS")

        cpfs_tab.columnconfigure(0, weight=1)
        self.missing_button = ttk.Button(
            cpfs_tab,
            text="Gerar arquivo para Lemiti",
            command=self._start_export_missing_cpfs,
            style="Primary.TButton",
        )
        self.missing_button.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.audit_xlsx_button = ttk.Button(
            cpfs_tab,
            text="Gerar planilha de conferencia",
            command=self._start_export_audit_xlsx,
            style="Quiet.TButton",
        )
        self.audit_xlsx_button.grid(row=1, column=0, sticky="ew")

        xlsx_tab.columnconfigure(1, weight=1)
        ttk.Label(xlsx_tab, text="Retorno Lemiti").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(xlsx_tab, textvariable=self.lemit_path).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(xlsx_tab, text="Selecionar", command=self._choose_lemit, style="Quiet.TButton").grid(
            row=0, column=2, sticky="e", padx=(8, 0), pady=4
        )
        ttk.Label(xlsx_tab, text="Salvar XLSX").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(xlsx_tab, textvariable=self.output_path).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(xlsx_tab, text="Local", command=self._choose_output, style="Quiet.TButton").grid(
            row=1, column=2, sticky="e", padx=(8, 0), pady=4
        )
        ttk.Label(xlsx_tab, text="Campanha").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(xlsx_tab, textvariable=self.argus_campaign).grid(row=2, column=1, columnspan=2, sticky="ew", pady=4)
        ttk.Label(xlsx_tab, text="CSV Argus").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(xlsx_tab, textvariable=self.argus_output_path).grid(row=3, column=1, sticky="ew", pady=4)
        ttk.Button(xlsx_tab, text="Local", command=self._choose_argus_output, style="Quiet.TButton").grid(
            row=3, column=2, sticky="e", padx=(8, 0), pady=4
        )
        self.generate_button = ttk.Button(
            xlsx_tab,
            text="▶ PROCESSAR BASE",
            command=self._start_generate,
            style="Primary.TButton",
        )
        self.generate_button.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(12, 8))
        self.argus_upload_button = ttk.Button(
            xlsx_tab,
            text="Enviar CSV para Argus",
            command=self._start_argus_upload,
            state="disabled",
            style="Quiet.TButton",
        )
        self.argus_upload_button.grid(row=5, column=0, columnspan=3, sticky="ew")

        sms_tab.columnconfigure(0, weight=1)
        self.sms_button = ttk.Button(
            sms_tab,
            text="Gerar arquivos SMS",
            command=self._start_export_sms_xlsx,
            style="Primary.TButton",
        )
        self.sms_button.grid(row=0, column=0, sticky="ew")

        self.result_frame = ttk.Frame(frame, padding=(0, 10), style="Minimal.TFrame")
        self.result_frame.grid(row=6, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(self.result_frame, text="Resultado", style="BlockTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(self.result_frame, textvariable=self.result_summary, style="Result.TLabel").grid(
            row=1, column=0, sticky="w", pady=(4, 0)
        )
        self.result_frame.grid_remove()

        ttk.Label(frame, textvariable=self.status, style="FooterStatus.TLabel").grid(
            row=7, column=0, sticky="ew", pady=(8, 0)
        )

    def _configure_minimal_styles(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Minimal.TFrame", background="#F7F7F5")
        style.configure("Minimal.TLabel", background="#F7F7F5")
        style.configure("LogoTile.TFrame", background="#111111")
        style.configure("LogoTile.TLabel", background="#111111")
        style.configure("MinimalTitle.TLabel", background="#F7F7F5", foreground="#111111", font=("Segoe UI", 16, "bold"))
        style.configure("MinimalSub.TLabel", background="#F7F7F5", foreground="#555555", font=("Segoe UI", 9))
        style.configure("BrandNeutral.TLabel", background="#F7F7F5", foreground="#111111", font=("Segoe UI", 17, "bold"), padding=(10, 0))
        style.configure("DigioBrand.TLabel", background="#F7F7F5", foreground="#23282F", font=("Segoe UI", 19, "bold"), padding=(10, 0))
        style.configure("Step.TLabel", background="#F7F7F5", foreground="#6A6A6A", font=("Segoe UI", 9, "bold"))
        style.configure("BlockTitle.TLabel", background="#F7F7F5", foreground="#111111", font=("Segoe UI", 11, "bold"))
        style.configure("Success.TLabel", background="#F7F7F5", foreground="#167A3A", font=("Segoe UI", 9, "bold"))
        style.configure("Result.TLabel", background="#F7F7F5", foreground="#111111", font=("Segoe UI", 10))
        style.configure("FooterStatus.TLabel", background="#F7F7F5", foreground="#666666", font=("Segoe UI", 9))
        style.configure("Primary.TButton", font=("Segoe UI", 11, "bold"), padding=(14, 10), background="#D7B56D", foreground="#111111")
        style.map("Primary.TButton", background=[("disabled", "#D8D8D8"), ("active", "#E2C57D"), ("!disabled", "#D7B56D")])
        style.configure("Quiet.TButton", font=("Segoe UI", 9), padding=(10, 6), background="#FFFFFF", foreground="#222222")

    def _build_modern_ui(self) -> None:
        self.root.configure(bg="#F5F3EE")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self._configure_styles()

        frame = ttk.Frame(self.root, padding=18, style="App.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)

        header = ttk.Frame(frame, padding=(18, 16), style="Header.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Auto Datafast para Argus", style="HeaderTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text="Prepare bases C6 e Digio com banco, UF e coeficiente automaticos.",
            style="HeaderSub.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.reset_button = ttk.Button(
            header,
            text="Limpar e comecar novamente (F5)",
            command=self._reset_form,
            style="Ghost.TButton",
        )
        self.reset_button.grid(row=0, column=1, rowspan=2, sticky="e")

        summary = ttk.Frame(frame, padding=(0, 14, 0, 8), style="App.TFrame")
        summary.grid(row=1, column=0, sticky="ew")
        for column in range(5):
            summary.columnconfigure(column, weight=1)
        self._summary_card(summary, "Arquivo atual", self.file_summary, 0)
        self._summary_card(summary, "Banco", self.bank_summary, 1)
        self._summary_card(summary, "UF", self.uf_summary, 2)
        self._summary_card(summary, "Coeficiente", self.coef_summary, 3)
        self._summary_card(summary, "Registros", self.records_summary, 4)

        workflow = ttk.Frame(frame, padding=(14, 10), style="Card.TFrame")
        workflow.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        workflow.columnconfigure(1, weight=1)
        ttk.Label(workflow, text="Etapa atual", style="MiniTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(workflow, textvariable=self.step_summary, style="Strong.TLabel").grid(
            row=1, column=0, sticky="w"
        )
        self.progress_bar = ttk.Progressbar(workflow, mode="determinate", maximum=100)
        self.progress_bar.grid(row=0, column=1, rowspan=2, sticky="ew", padx=18)
        ttk.Label(workflow, textvariable=self.progress_summary, style="Muted.TLabel").grid(
            row=0, column=2, rowspan=2, sticky="e"
        )

        body = ttk.Frame(frame, style="App.TFrame")
        body.grid(row=3, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(2, weight=1)

        source_card = ttk.LabelFrame(body, text="1. Selecionar base", padding=14, style="Card.TLabelframe")
        source_card.grid(row=0, column=0, sticky="ew")
        source_card.columnconfigure(1, weight=1)
        ttk.Label(source_card, text="Escolha a exportacao DTF CSV do C6 ou Digio.", style="Strong.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        ttk.Label(
            source_card,
            text="O sistema identifica banco, UF, coeficiente e quantidade de registros.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 8))
        ttk.Label(source_card, text="Base DTF CSV").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(source_card, textvariable=self.csv_path).grid(row=2, column=1, sticky="ew", pady=5)
        ttk.Button(source_card, text="Selecionar base", command=self._choose_csv, style="Accent.TButton").grid(
            row=2, column=2, padx=(10, 0), pady=5
        )
        ttk.Label(source_card, textvariable=self.bank_info, style="Success.TLabel").grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )

        filters = ttk.LabelFrame(body, text="2. Conferir filtros", padding=14, style="Card.TLabelframe")
        filters.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        for column in range(4):
            filters.columnconfigure(column, weight=1)

        self._add_labeled_entry(filters, "Prazos FALTAM", self.faltam, 0, 0)
        self._add_labeled_combo(
            filters,
            "Regra dos prazos",
            self.faltam_modo,
            ["Maior que o prazo", "Ate o prazo", "Exato"],
            0,
            1,
        )
        self._add_labeled_entry(filters, "Taxa minima", self.taxa_min, 0, 2)
        self._add_labeled_entry(filters, "Troco maior que", self.troco_min, 0, 3)
        self._add_labeled_entry(filters, "Nascimento inicial", self.data_inicial, 1, 0)
        self._add_labeled_entry(filters, "Nascimento final", self.data_final, 1, 1)
        self._add_labeled_entry(filters, "Periodo minimo do telefone", self.telefone_data_inicial, 1, 2)
        ttk.Label(
            filters,
            text="Dica: deixe FALTAM vazio para todos, ou use varios prazos separados por virgula, como 67,72.",
            style="Muted.TLabel",
        ).grid(row=2, column=0, columnspan=4, sticky="w", pady=(8, 0))

        self.notebook = ttk.Notebook(body)
        self.notebook.grid(row=2, column=0, sticky="nsew", pady=(12, 0))

        cpfs_tab = ttk.Frame(self.notebook, padding=16)
        xlsx_tab = ttk.Frame(self.notebook, padding=16)
        sms_tab = ttk.Frame(self.notebook, padding=16)
        self.notebook.add(cpfs_tab, text="Preparar para Lemiti")
        self.notebook.add(xlsx_tab, text="Finalizar Base e Argus")
        self.notebook.add(sms_tab, text="Gerar SMS")

        cpfs_tab.columnconfigure(0, weight=1)
        ttk.Label(cpfs_tab, text="3. Gerar CPFs que ainda precisam de consulta", style="Section.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            cpfs_tab,
            text="Use esta etapa para mandar ao Lemiti apenas os CPFs filtrados sem telefone base.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 12))
        self.missing_button = ttk.Button(
            cpfs_tab,
            text="Gerar arquivo para Lemiti",
            command=self._start_export_missing_cpfs,
            style="Accent.TButton",
        )
        self.missing_button.grid(row=2, column=0, sticky="w")
        self.audit_xlsx_button = ttk.Button(
            cpfs_tab,
            text="Gerar planilha de conferencia",
            command=self._start_export_audit_xlsx,
        )
        self.audit_xlsx_button.grid(row=3, column=0, sticky="w", pady=(10, 0))

        xlsx_tab.columnconfigure(1, weight=1)
        ttk.Label(xlsx_tab, text="4. Finalizar telefones e gerar Argus", style="Section.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        ttk.Label(
            xlsx_tab,
            text="Selecione o retorno Lemiti. O XLSX final e o CSV Argus saem juntos.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 12))
        ttk.Label(xlsx_tab, text="Retorno Lemiti").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(xlsx_tab, textvariable=self.lemit_path).grid(row=2, column=1, sticky="ew", pady=5)
        ttk.Button(xlsx_tab, text="Selecionar retorno", command=self._choose_lemit).grid(
            row=2, column=2, padx=(10, 0), pady=5
        )

        ttk.Label(xlsx_tab, text="Salvar XLSX final").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Entry(xlsx_tab, textvariable=self.output_path).grid(row=3, column=1, sticky="ew", pady=5)
        ttk.Button(xlsx_tab, text="Escolher local", command=self._choose_output).grid(
            row=3, column=2, padx=(10, 0), pady=5
        )

        ttk.Label(xlsx_tab, text="Campanha Argus").grid(row=4, column=0, sticky="w", pady=5)
        ttk.Entry(xlsx_tab, textvariable=self.argus_campaign).grid(row=4, column=1, sticky="ew", pady=5)

        ttk.Label(xlsx_tab, text="Salvar CSV Argus").grid(row=5, column=0, sticky="w", pady=5)
        ttk.Entry(xlsx_tab, textvariable=self.argus_output_path).grid(row=5, column=1, sticky="ew", pady=5)
        ttk.Button(xlsx_tab, text="Escolher local", command=self._choose_argus_output).grid(
            row=5, column=2, padx=(10, 0), pady=5
        )
        self.generate_button = ttk.Button(
            xlsx_tab, text="Gerar XLSX final e CSV Argus", command=self._start_generate, style="Accent.TButton"
        )
        self.generate_button.grid(row=6, column=2, sticky="e", pady=(14, 0))
        self.argus_upload_button = ttk.Button(
            xlsx_tab,
            text="Enviar CSV para Argus",
            command=self._start_argus_upload,
            state="disabled",
        )
        self.argus_upload_button.grid(row=6, column=1, sticky="e", pady=(14, 0), padx=(8, 8))
        ttk.Label(xlsx_tab, text="Acesso Argus usa configuracao local segura.", style="Muted.TLabel").grid(
            row=6, column=0, sticky="w", pady=(14, 0)
        )

        sms_tab.columnconfigure(0, weight=1)
        ttk.Label(sms_tab, text="5. Gerar arquivos para SMS", style="Section.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            sms_tab,
            text="Gera o CAMP de SMS e o arquivo auto-enviados usando a mesma base filtrada.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 12))
        self.sms_button = ttk.Button(
            sms_tab,
            text="Gerar arquivos SMS",
            command=self._start_export_sms_xlsx,
            style="Accent.TButton",
        )
        self.sms_button.grid(row=2, column=0, sticky="w")

        status_card = ttk.Frame(frame, padding=(14, 10), style="Card.TFrame")
        status_card.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        status_card.columnconfigure(0, weight=1)
        ttk.Label(status_card, textvariable=self.status, style="Status.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(status_card, text=FOOTER_TEXT, style="Muted.TLabel").grid(row=0, column=1, sticky="e")

    def _configure_styles(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("App.TFrame", background="#F5F3EE")
        style.configure("Header.TFrame", background="#111111")
        style.configure("Card.TFrame", background="#FFFFFF", relief="flat")
        style.configure("Card.TLabelframe", background="#FFFFFF", bordercolor="#D8D2C2")
        style.configure("Card.TLabelframe.Label", background="#F5F3EE", foreground="#111111", font=("Segoe UI", 10, "bold"))
        style.configure("HeaderTitle.TLabel", background="#111111", foreground="#FFFFFF", font=("Segoe UI", 20, "bold"))
        style.configure("HeaderSub.TLabel", background="#111111", foreground="#D7B56D", font=("Segoe UI", 10))
        style.configure("SummaryTitle.TLabel", background="#FFFFFF", foreground="#6B6B6B", font=("Segoe UI", 9))
        style.configure("SummaryValue.TLabel", background="#FFFFFF", foreground="#111111", font=("Segoe UI", 12, "bold"))
        style.configure("MiniTitle.TLabel", background="#FFFFFF", foreground="#6B6B6B", font=("Segoe UI", 9))
        style.configure("Strong.TLabel", background="#FFFFFF", foreground="#111111", font=("Segoe UI", 10, "bold"))
        style.configure("Section.TLabel", background="#FFFFFF", foreground="#111111", font=("Segoe UI", 13, "bold"))
        style.configure("Muted.TLabel", background="#FFFFFF", foreground="#666666", font=("Segoe UI", 9))
        style.configure("Status.TLabel", background="#FFFFFF", foreground="#1F4E78", font=("Segoe UI", 10, "bold"))
        style.configure("Success.TLabel", background="#FFFFFF", foreground="#167A3A", font=("Segoe UI", 10, "bold"))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8))
        style.map("Accent.TButton", foreground=[("disabled", "#888888"), ("!disabled", "#111111")])
        style.configure("Ghost.TButton", font=("Segoe UI", 9), padding=(12, 7))

    def _summary_card(self, parent: ttk.Frame, title: str, variable: StringVar, column: int) -> None:
        card = ttk.Frame(parent, padding=(12, 10), style="Card.TFrame")
        card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
        ttk.Label(card, text=title, style="SummaryTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(card, textvariable=variable, style="SummaryValue.TLabel").grid(row=1, column=0, sticky="w")

    def _add_labeled_entry(
        self, parent: ttk.Frame, label: str, variable: StringVar, row: int, column: int
    ) -> None:
        box = ttk.Frame(parent, style="Minimal.TFrame")
        box.grid(row=row, column=column, sticky="ew", padx=(0, 10), pady=2)
        box.columnconfigure(0, weight=1)
        ttk.Label(box, text=label).grid(row=0, column=0, sticky="w")
        ttk.Entry(box, textvariable=variable).grid(row=1, column=0, sticky="ew", pady=(2, 0))

    def _add_labeled_combo(
        self,
        parent: ttk.Frame,
        label: str,
        variable: StringVar,
        values: list[str],
        row: int,
        column: int,
    ) -> None:
        box = ttk.Frame(parent, style="Minimal.TFrame")
        box.grid(row=row, column=column, sticky="ew", padx=(0, 10), pady=2)
        box.columnconfigure(0, weight=1)
        ttk.Label(box, text=label).grid(row=0, column=0, sticky="w")
        ttk.Combobox(box, textvariable=variable, values=values, state="readonly").grid(
            row=1, column=0, sticky="ew", pady=(2, 0)
        )

    def _reset_form(self, event: object | None = None) -> None:
        self.csv_path.set("")
        self.lemit_path.set("")
        self.output_path.set("")
        self.argus_output_path.set("")
        self.argus_campaign.set("")
        self._reset_filter_fields()
        self.bank_info.set("")
        self.file_summary.set("Nenhum arquivo selecionado")
        self.bank_summary.set("-")
        self.uf_summary.set("-")
        self.coef_summary.set("-")
        self.records_summary.set("-")
        self.step_summary.set("Etapa atual: Selecionando Base")
        self.progress_summary.set("Aguardando arquivo")
        self.result_summary.set("")
        if hasattr(self, "base_summary_label"):
            self.base_summary_label.grid_remove()
        if hasattr(self, "result_frame"):
            self.result_frame.grid_remove()
        if hasattr(self, "progress_bar"):
            self.progress_bar.configure(mode="determinate")
            self.progress_bar["value"] = 0
        self.status.set("Tela limpa. Selecione uma base para comecar.")
        self.last_argus_csv_path = None
        self._set_source_ready(False)
        self.notebook.select(0)

    def _reset_filter_fields(self) -> None:
        self.faltam.set("")
        self.faltam_modo.set("Maior que o prazo")
        self.taxa_min.set("0,00")
        self.troco_min.set("0,00")
        self.data_inicial.set("01/01/1954")
        self.data_final.set("31/12/1995")
        self.telefone_data_inicial.set("01/01/2026")
        self._set_brand_badge("")

    def _set_brand_badge(self, bank_code: str) -> None:
        badge = getattr(self, "brand_badge", None)
        if badge is None:
            return
        if bank_code == "c6" and "c6" in self.brand_images:
            badge.configure(image=self.brand_images["c6"], text="", style="BrandNeutral.TLabel")
            return
        if bank_code == "digio":
            badge.configure(image="", text="digio", style="DigioBrand.TLabel")
            return
        badge.configure(image="", text="", style="BrandNeutral.TLabel")

    def _choose_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="Escolha a Extração DTF CSV",
            filetypes=[("CSV", "*.csv"), ("Todos os arquivos", "*.*")],
        )
        if not path:
            return

        self.csv_path.set(path)
        self._load_source_summary(path)

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

    def _update_bank_info(self) -> None:
        try:
            self.bank_info.set(format_bank_info(self._get_csv_path()))
        except Exception as exc:
            self.bank_info.set(str(exc))

    def _load_source_summary(self, path: str) -> None:
        try:
            bank = identify_bank_from_filename(path)
            uf = extract_uf_from_filename(Path(path).stem)
            rows = len(read_c6_csv(path))
        except Exception as exc:
            self._set_source_ready(False)
            self.file_summary.set(Path(path).name)
            self.bank_summary.set("-")
            self.uf_summary.set("-")
            self.coef_summary.set("-")
            self.records_summary.set("-")
            self.bank_info.set("")
            self._set_brand_badge("")
            if hasattr(self, "base_summary_label"):
                self.base_summary_label.grid_remove()
            self.status.set("Arquivo incorreto. Selecione uma base com C6 ou Digio no nome.")
            self.step_summary.set("Etapa atual: Selecionando Base")
            self.progress_summary.set("Arquivo precisa ser corrigido")
            messagebox.showerror(APP_NAME, self._friendly_error(exc))
            return

        self._reset_filter_fields()
        self._set_brand_badge(bank.code)
        self.file_summary.set(_shorten_filename(Path(path).name))
        self.bank_summary.set(bank.display_name)
        self.uf_summary.set(uf)
        self.coef_summary.set(str(bank.coefficient).replace(".", ","))
        self.records_summary.set(_format_count(rows))
        self.bank_info.set(
            f"✓ Base carregada  Banco: {bank.display_name} | UF: {uf} | Registros: {_format_count(rows)} | Coeficiente: {str(bank.coefficient).replace('.', ',')}"
        )
        if hasattr(self, "base_summary_label"):
            self.base_summary_label.grid()
        self.result_summary.set("")
        if hasattr(self, "result_frame"):
            self.result_frame.grid_remove()
        self.status.set("Base pronta. Confira os filtros e escolha a proxima etapa.")
        self.step_summary.set("Etapa atual: Conferindo filtros")
        self.progress_summary.set("Base validada")
        if hasattr(self, "progress_bar"):
            self.progress_bar.configure(mode="determinate")
            self.progress_bar["value"] = 25
        self._set_source_ready(True)
        self._suggest_argus_campaign()

    def _set_source_ready(self, ready: bool) -> None:
        self.source_ready = ready
        state = "normal" if ready else "disabled"
        for button_name in ("missing_button", "audit_xlsx_button", "generate_button", "sms_button"):
            button = getattr(self, button_name, None)
            if button is not None:
                button.configure(state=state)
        upload_state = "normal" if ready and self.last_argus_csv_path is not None else "disabled"
        button = getattr(self, "argus_upload_button", None)
        if button is not None:
            button.configure(state=upload_state)

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
        self.step_summary.set("Etapa atual: Preparando arquivos")
        self.progress_summary.set("Processando telefones, Lemiti e BDD/DDB")
        self.status.set("Processando base final e preparando arquivos da Argus...")
        thread = threading.Thread(target=self._generate, daemon=True)
        thread.start()

    def _start_argus_upload(self) -> None:
        path = self.last_argus_csv_path
        if path is None or not path.is_file():
            messagebox.showerror(APP_NAME, "Gere o CSV Argus antes de enviar.")
            return
        try:
            config = load_argus_config()
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        if config is None:
            messagebox.showerror(
                APP_NAME,
                "Acesso Argus não configurado nesta máquina. Salve a Token-Signature localmente antes de enviar.",
            )
            return
        if file_was_uploaded(path) and not messagebox.askyesno(
            APP_NAME,
            "Este mesmo conteúdo já foi enviado com sucesso para a Argus. Deseja reenviar?",
        ):
            return

        self._set_busy(True)
        confirmation = (
            "Confirmar envio para a Argus\n\n"
            f"Banco: {self.bank_summary.get()}\n"
            f"UF: {self.uf_summary.get()}\n"
            f"Campanha: {self.argus_campaign.get().strip() or '-'}\n"
            f"Arquivo: {path.name}\n\n"
            "Deseja enviar este CSV agora?"
        )
        if not messagebox.askyesno(APP_NAME, confirmation):
            self._set_busy(False)
            return
        self.step_summary.set("Etapa atual: Enviando para Argus")
        self.progress_summary.set("Enviando arquivo confirmado")
        self.status.set(f"Enviando {path.name} para a Argus...")
        thread = threading.Thread(target=self._upload_argus, args=(path, config), daemon=True)
        thread.start()

    def _upload_argus(self, path: Path, config) -> None:
        try:
            result = upload_mailing_csv(
                path,
                config.token_signature,
                config.skill_name or DEFAULT_SKILL_REFERENCE,
            )
            record_upload(path, result)
            message = (
                "Mailing enviado para a Argus.\n"
                f"Arquivo: {path.name}\n"
                f"idArquivo: {result.id_arquivo}\n"
                f"codStatus: {result.cod_status}\n"
                f"descStatus: {result.desc_status}"
            )
            self.root.after(0, self._argus_upload_success, message)
        except Exception as exc:
            self.root.after(0, self._generation_error, str(exc))

    def _start_export_missing_cpfs(self) -> None:
        try:
            csv_path = self._get_csv_path()
        except ValueError as exc:
            messagebox.showerror(APP_NAME, self._friendly_error(exc))
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
        self.step_summary.set("Etapa atual: Preparando Lemiti")
        self.progress_summary.set("Buscando telefones no HUB")
        self.status.set("Buscando telefones e separando CPFs sem telefone...")
        thread = threading.Thread(
            target=self._export_missing_cpfs, args=(output_path,), daemon=True
        )
        thread.start()

    def _start_export_audit_xlsx(self) -> None:
        try:
            csv_path = self._get_csv_path()
        except ValueError as exc:
            messagebox.showerror(APP_NAME, self._friendly_error(exc))
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
        self.step_summary.set("Etapa atual: Preparando conferencia")
        self.progress_summary.set("Gerando conferencia")
        self.status.set("Gerando planilha de conferencia com telefones da base...")
        thread = threading.Thread(
            target=self._export_audit_xlsx, args=(output_path,), daemon=True
        )
        thread.start()

    def _start_export_sms_xlsx(self) -> None:
        try:
            csv_path = self._get_csv_path()
        except ValueError as exc:
            messagebox.showerror(APP_NAME, self._friendly_error(exc))
            return

        lemit_path = self.lemit_path.get().strip()
        if not lemit_path:
            messagebox.showerror(APP_NAME, self._friendly_error("Escolha o retorno Lemiti antes de gerar o SMS."))
            return

        self._set_busy(True)
        self.step_summary.set("Etapa atual: Gerando SMS")
        self.progress_summary.set("Preparando CAMP e auto-enviados")
        self.status.set("Gerando arquivos de SMS na pasta do Disparo Pro...")
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
            missing_count = (filtered["TELEFONE BASE"].fillna("").astype(str).str.strip() == "").sum()
            found_count = len(filtered) - int(missing_count)
            message = (
                "✓ Arquivos gerados\n"
                f"✓ Telefones encontrados: {_format_count(found_count)}\n"
                f"✓ CPFs sem telefone: {_format_count(int(missing_count))}\n"
                f"✓ Pasta salva: {Path(argus_output_path).parent}"
            )
            self.root.after(0, self._generation_ready, message, argus_output_path)
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
            missing_count = (filtered["TELEFONE BASE"].fillna("").astype(str).str.strip() == "").sum()
            message = (
                "✓ Arquivos gerados\n"
                f"✓ Telefones encontrados: {_format_count(len(filtered) - int(missing_count))}\n"
                f"✓ CPFs sem telefone: {_format_count(int(missing_count))}\n"
                f"✓ Pasta salva: {Path(output_path).parent}"
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
                "✓ Arquivos gerados\n"
                f"✓ Telefones encontrados: {_format_count(len(filtered) - len(missing))}\n"
                f"✓ CPFs sem telefone: {_format_count(len(missing))}\n"
                f"✓ Pasta salva: {Path(output_path).parent}"
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
                "✓ Arquivos gerados\n"
                f"✓ Telefones encontrados: {_format_count(len(filtered) - len(missing))}\n"
                f"✓ CPFs sem telefone: {_format_count(len(missing))}\n"
                f"✓ Pasta salva: {Path(output_path).parent}"
            )
            self.root.after(0, self._generation_success, message)
        except Exception as exc:
            self.root.after(0, self._generation_error, str(exc))

    def _get_csv_path(self) -> str:
        csv_path = self.csv_path.get().strip()
        if not csv_path:
            raise ValueError("Escolha a Extração DTF CSV.")
        identify_bank_from_filename(csv_path)
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
        action_state = "disabled" if busy or not self.source_ready else state
        self.generate_button.configure(state=action_state)
        self.missing_button.configure(state=action_state)
        self.audit_xlsx_button.configure(state=action_state)
        self.sms_button.configure(state=action_state)
        upload_state = "disabled" if busy or self.last_argus_csv_path is None else "normal"
        self.argus_upload_button.configure(state=upload_state)
        if hasattr(self, "progress_bar"):
            if busy:
                self.progress_bar.configure(mode="indeterminate")
                self.progress_bar.start(12)
            else:
                self.progress_bar.stop()
                self.progress_bar.configure(mode="determinate")
                if self.source_ready:
                    self.progress_bar["value"] = max(float(self.progress_bar["value"]), 25)

    def _generation_ready(self, message: str, argus_output_path: str) -> None:
        self.last_argus_csv_path = Path(argus_output_path)
        self._generation_success(message)
        self.status.set("Arquivos prontos. O envio para a Argus foi habilitado.")

    def _argus_upload_success(self, message: str) -> None:
        self._set_busy(False)
        self.status.set("Mailing enviado para a Argus com sucesso.")
        self.progress_summary.set("Envio concluido")
        if hasattr(self, "progress_bar"):
            self.progress_bar["value"] = 100
        messagebox.showinfo(APP_NAME, message)

    def _generation_success(self, message: str) -> None:
        self._set_busy(False)
        self.result_summary.set(message)
        if hasattr(self, "result_frame"):
            self.result_frame.grid()
        self.status.set("Concluido.")
        self.progress_summary.set("Concluido")
        if hasattr(self, "progress_bar"):
            self.progress_bar["value"] = 100
        messagebox.showinfo(APP_NAME, message)

    def _generation_error(self, message: str) -> None:
        self._set_busy(False)
        friendly = self._friendly_error(message)
        self.status.set("Nao foi possivel concluir. Veja a orientacao exibida.")
        self.progress_summary.set("Corrija o ponto indicado e tente novamente")
        messagebox.showerror(APP_NAME, friendly)

    def _friendly_error(self, error: object) -> str:
        text = str(error)
        if "Banco nao identificado" in text:
            return (
                "Arquivo incorreto\n\n"
                "Nao consegui identificar se esta base pertence ao C6 ou ao Digio.\n\n"
                'Selecione um arquivo que contenha "C6" ou "Digio" no nome.'
            )
        if "Cabe" in text and "ausentes" in text:
            columns = text.split(":", 1)[-1].strip()
            return (
                "Arquivo com colunas faltando\n\n"
                f"Nao encontrei estas colunas obrigatorias: {columns}.\n\n"
                "Confira se voce selecionou a exportacao DTF correta."
            )
        if "retorno Lemiti" in text or "Retorno Lemiti" in text:
            return (
                "Retorno Lemiti pendente\n\n"
                "Para finalizar a base, selecione primeiro o arquivo de retorno do Lemiti."
            )
        return (
            "Nao foi possivel concluir esta etapa\n\n"
            f"{text}\n\n"
            "Confira os arquivos selecionados e tente novamente."
        )


def _shorten_filename(name: str, limit: int = 34) -> str:
    if len(name) <= limit:
        return name
    return f"{name[: limit - 3]}..."


def _format_count(value: int) -> str:
    return f"{int(value):,}".replace(",", ".")


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
