from __future__ import annotations

import argparse
import csv
import io
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import BinaryIO, Iterable

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell


COEF_VALUE = 0.0228
DEFAULT_PHONE_BASE_DIR = Path(
    r"G:\Meu Drive\GRUPO IMPÉRIO\LEADS\TELEFONES\HUB CONSOLIDADO"
)
FALLBACK_PHONE_BASE_DIR = Path(
    r"H:\Meu Drive\GRUPO IMPÉRIO\LEADS\TELEFONES\HUB CONSOLIDADO"
)
DEFAULT_PHONE_BASE_DIRS = (DEFAULT_PHONE_BASE_DIR, FALLBACK_PHONE_BASE_DIR)
ARGUS_DISC_BASE_DIRS = (
    Path(r"G:\Meu Drive\GRUPO IMPÉRIO\LEADS\ARGUS\DISC"),
    Path(r"H:\Meu Drive\GRUPO IMPÉRIO\LEADS\ARGUS\DISC"),
)
BRAZIL_UFS = {
    "AC",
    "AL",
    "AM",
    "AP",
    "BA",
    "CE",
    "DF",
    "ES",
    "GO",
    "MA",
    "MG",
    "MS",
    "MT",
    "PA",
    "PB",
    "PE",
    "PI",
    "PR",
    "RJ",
    "RN",
    "RO",
    "RR",
    "RS",
    "SC",
    "SE",
    "SP",
    "TO",
}
MONTH_NAMES_PT_BR = {
    1: "JANEIRO",
    2: "FEVEREIRO",
    3: "MARÇO",
    4: "ABRIL",
    5: "MAIO",
    6: "JUNHO",
    7: "JULHO",
    8: "AGOSTO",
    9: "SETEMBRO",
    10: "OUTUBRO",
    11: "NOVEMBRO",
    12: "DEZEMBRO",
}

SOURCE_COLUMNS = {
    "ovc_cpf": "CPF",
    "cr_matricula": "BENEFÍCIO",
    "cr_parcelas_contrato": "FEZ EM",
    "cr_parcelas_em_aberto": "FALTAM",
    "cr_valor_parcela": "PARCELA",
    "cr_saldo_refin": "SALDO DEV",
    "cr_taxa": "TAXA",
    "dc_nome_1": "NOME",
    "dc_data_nascimento_1": "DATA DE NASC",
    "tf_ddd_1": "DDD C6",
    "tf_telefone_1": "TEL C6",
}

IGNORED_COLUMNS = {"cr_valor_contrato", "cr_observacao"}

FINAL_COLUMNS = [
    "CPF",
    "BENEFÍCIO",
    "FEZ EM",
    "FALTAM",
    "PARCELA",
    "SALDO DEV",
    "TAXA",
    "NOME",
    "DATA DE NASC",
    "DDD C6",
    "TEL C6",
    "TELEFONE C6",
    "TELEFONE BASE",
    "COEF",
    "BRUTO",
    "TROCO",
]

ARGUS_DISCADORA_COLUMNS = [
    "CAMPANHA",
    "CPF",
    "Taxa",
    "FEZ EM",
    "FALTAM",
    " PARCELA ",
    "TROCO",
    "Telefone",
    "Nome",
    "Data Nascimento",
    "TELEFONE C6",
]


@dataclass(frozen=True)
class FilterOptions:
    faltam: tuple[int, ...] = ()
    faltam_modo: str = "maior"
    taxa_min: float = 0.0
    troco_min: float = 0.0
    data_inicial: date | None = date(1954, 1, 1)
    data_final: date | None = date(1995, 12, 31)
    telefone_data_inicial: date | None = None


def read_c6_csv(source: str | Path | bytes | BinaryIO) -> pd.DataFrame:
    data = _read_bytes(source)
    last_error: Exception | None = None

    for encoding in ("utf-8-sig", "cp1252", "latin1"):
        try:
            text = data.decode(encoding)
            df = pd.read_csv(io.StringIO(text), sep=";", dtype=str)
            df.columns = [str(col).strip().lstrip("\ufeff") for col in df.columns]
            return df
        except Exception as exc:
            last_error = exc

    raise ValueError(f"Não consegui ler o CSV. Último erro: {last_error}")


def transform_c6_dataframe(raw_df: pd.DataFrame) -> pd.DataFrame:
    missing = [col for col in SOURCE_COLUMNS if col not in raw_df.columns]
    if missing:
        joined = ", ".join(missing)
        raise ValueError(f"Cabeçalhos obrigatórios ausentes: {joined}")

    output = pd.DataFrame()
    output["CPF"] = raw_df["ovc_cpf"].map(_digits_only)
    output["BENEFÍCIO"] = raw_df["cr_matricula"].map(_digits_only)
    output["FEZ EM"] = _parse_int_series(raw_df["cr_parcelas_contrato"])
    output["FALTAM"] = _parse_int_series(raw_df["cr_parcelas_em_aberto"])
    output["PARCELA"] = raw_df["cr_valor_parcela"].map(_parse_br_number)
    output["SALDO DEV"] = raw_df["cr_saldo_refin"].map(_parse_br_number)
    output["TAXA"] = raw_df["cr_taxa"].map(_parse_br_number)
    output["NOME"] = raw_df["dc_nome_1"].fillna("").astype(str).str.strip()
    output["DATA DE NASC"] = pd.to_datetime(
        raw_df["dc_data_nascimento_1"], dayfirst=True, errors="coerce"
    )
    output["DDD C6"] = raw_df["tf_ddd_1"].map(_digits_only)
    output["TEL C6"] = raw_df["tf_telefone_1"].map(_digits_only)
    output["TELEFONE C6"] = output["DDD C6"] + output["TEL C6"]
    output["TELEFONE BASE"] = ""
    output["COEF"] = COEF_VALUE
    output["BRUTO"] = output["PARCELA"] / output["COEF"]
    output["TROCO"] = output["BRUTO"] - output["SALDO DEV"]

    return output[FINAL_COLUMNS]


def apply_filters(df: pd.DataFrame, options: FilterOptions) -> pd.DataFrame:
    mask = df["SALDO DEV"].fillna(0) > 0

    if options.faltam:
        faltam_values = tuple(int(value) for value in options.faltam)
        if options.faltam_modo == "exato":
            mask &= df["FALTAM"].isin(faltam_values)
        elif options.faltam_modo == "ate":
            mask &= df["FALTAM"] <= max(faltam_values)
        elif options.faltam_modo == "a_partir":
            mask &= df["FALTAM"] >= min(faltam_values)
        else:
            mask &= df["FALTAM"] > min(faltam_values)

    if options.taxa_min is not None:
        mask &= df["TAXA"].fillna(float("-inf")) >= float(options.taxa_min)

    if options.troco_min is not None:
        mask &= df["TROCO"].fillna(float("-inf")) > float(options.troco_min)

    if options.data_inicial is not None:
        start = pd.Timestamp(options.data_inicial)
        mask &= df["DATA DE NASC"].notna() & (df["DATA DE NASC"] >= start)

    if options.data_final is not None:
        end = pd.Timestamp(options.data_final)
        mask &= df["DATA DE NASC"].notna() & (df["DATA DE NASC"] <= end)

    return df.loc[mask].copy().reset_index(drop=True)


def parse_faltam_values(text: str | Iterable[int] | None) -> tuple[int, ...]:
    if text is None:
        return ()

    if not isinstance(text, str):
        return tuple(int(value) for value in text)

    cleaned = text.strip()
    if not cleaned:
        return ()

    values: list[int] = []
    for part in re.split(r"[,\s;/]+", cleaned):
        if part:
            values.append(int(part))

    return tuple(dict.fromkeys(values))


def parse_date_value(value: str | date | None) -> date | None:
    if value is None or value == "":
        return None

    if isinstance(value, date):
        return value

    parsed = pd.to_datetime(str(value), dayfirst=True, errors="raise")
    return parsed.date()


def write_xlsx(df: pd.DataFrame, output_path: str | Path | BinaryIO) -> None:
    if hasattr(output_path, "write"):
        workbook = xlsxwriter.Workbook(output_path, {"in_memory": True})
    else:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        workbook = xlsxwriter.Workbook(str(output_path), {"constant_memory": True})
    worksheet = workbook.add_worksheet("C6 Filtrado")

    header_fmt = workbook.add_format(
        {
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#1F4E78",
            "border": 1,
            "border_color": "#B7C9D6",
            "align": "center",
            "valign": "vcenter",
        }
    )
    text_fmt = workbook.add_format({"num_format": "@", "valign": "vcenter"})
    digit_number_fmt = workbook.add_format({"num_format": "0", "valign": "vcenter"})
    int_fmt = workbook.add_format({"num_format": "#,##0", "valign": "vcenter"})
    number_fmt = workbook.add_format({"num_format": "0.00", "valign": "vcenter"})
    coef_fmt = workbook.add_format({"num_format": "0.0000", "valign": "vcenter"})
    money_fmt = workbook.add_format(
        {"num_format": '[$R$-pt-BR] #,##0.00', "valign": "vcenter"}
    )
    date_fmt = workbook.add_format({"num_format": "dd/mm/yyyy", "valign": "vcenter"})

    worksheet.set_tab_color("#1F4E78")
    worksheet.freeze_panes(1, 0)
    worksheet.set_row(0, 24)

    for col_idx, column in enumerate(FINAL_COLUMNS):
        worksheet.write(0, col_idx, column, header_fmt)

    col_index = {name: idx for idx, name in enumerate(FINAL_COLUMNS)}

    for row_idx, row in df.iterrows():
        excel_row = row_idx + 1
        _write_digits_number(worksheet, excel_row, col_index["CPF"], row["CPF"], digit_number_fmt)
        _write_digits_number(
            worksheet, excel_row, col_index["BENEFÍCIO"], row["BENEFÍCIO"], digit_number_fmt
        )
        _write_number(worksheet, excel_row, col_index["FEZ EM"], row["FEZ EM"], int_fmt)
        _write_number(worksheet, excel_row, col_index["FALTAM"], row["FALTAM"], int_fmt)
        _write_number(worksheet, excel_row, col_index["PARCELA"], row["PARCELA"], money_fmt)
        _write_number(
            worksheet, excel_row, col_index["SALDO DEV"], row["SALDO DEV"], money_fmt
        )
        _write_number(worksheet, excel_row, col_index["TAXA"], row["TAXA"], number_fmt)
        worksheet.write_string(excel_row, col_index["NOME"], _safe_text(row["NOME"]), text_fmt)
        _write_date(worksheet, excel_row, col_index["DATA DE NASC"], row["DATA DE NASC"], date_fmt)
        _write_digits_number(
            worksheet, excel_row, col_index["DDD C6"], row["DDD C6"], digit_number_fmt
        )
        _write_digits_number(
            worksheet, excel_row, col_index["TEL C6"], row["TEL C6"], digit_number_fmt
        )
        _write_digits_number(
            worksheet, excel_row, col_index["TELEFONE C6"], row["TELEFONE C6"], digit_number_fmt
        )
        telefone_base = _safe_text(row["TELEFONE BASE"]).strip()
        if telefone_base:
            worksheet.write_string(
                excel_row, col_index["TELEFONE BASE"], telefone_base, text_fmt
            )
        else:
            worksheet.write_blank(excel_row, col_index["TELEFONE BASE"], None, text_fmt)
        _write_number(worksheet, excel_row, col_index["COEF"], row["COEF"], coef_fmt)

        parcela_cell = xl_rowcol_to_cell(excel_row, col_index["PARCELA"])
        saldo_cell = xl_rowcol_to_cell(excel_row, col_index["SALDO DEV"])
        coef_cell = xl_rowcol_to_cell(excel_row, col_index["COEF"])
        bruto_cell = xl_rowcol_to_cell(excel_row, col_index["BRUTO"])
        worksheet.write_formula(
            excel_row,
            col_index["BRUTO"],
            f"={parcela_cell}/{coef_cell}",
            money_fmt,
            _number_or_blank(row["BRUTO"]),
        )
        worksheet.write_formula(
            excel_row,
            col_index["TROCO"],
            f"={bruto_cell}-{saldo_cell}",
            money_fmt,
            _number_or_blank(row["TROCO"]),
        )

    last_row = max(len(df), 1)
    worksheet.autofilter(0, 0, last_row, len(FINAL_COLUMNS) - 1)

    widths = {
        "CPF": 15,
        "BENEFÍCIO": 14,
        "FEZ EM": 10,
        "FALTAM": 10,
        "PARCELA": 13,
        "SALDO DEV": 14,
        "TAXA": 10,
        "NOME": 34,
        "DATA DE NASC": 14,
        "DDD C6": 9,
        "TEL C6": 13,
        "TELEFONE C6": 15,
        "TELEFONE BASE": 16,
        "COEF": 10,
        "BRUTO": 14,
        "TROCO": 14,
    }
    formats = {
        "CPF": digit_number_fmt,
        "BENEFÍCIO": digit_number_fmt,
        "FEZ EM": int_fmt,
        "FALTAM": int_fmt,
        "PARCELA": money_fmt,
        "SALDO DEV": money_fmt,
        "TAXA": number_fmt,
        "NOME": text_fmt,
        "DATA DE NASC": date_fmt,
        "DDD C6": digit_number_fmt,
        "TEL C6": digit_number_fmt,
        "TELEFONE C6": digit_number_fmt,
        "TELEFONE BASE": text_fmt,
        "COEF": coef_fmt,
        "BRUTO": money_fmt,
        "TROCO": money_fmt,
    }

    for column, width in widths.items():
        idx = col_index[column]
        worksheet.set_column(idx, idx, width, formats[column])

    workbook.set_properties({"title": "C6 filtrado", "subject": "Exportação C6 tratada"})
    workbook.close()


def dataframe_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    write_xlsx(df, buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_argus_discadora_dataframe(df: pd.DataFrame, campanha: str = "") -> pd.DataFrame:
    output = pd.DataFrame(index=df.index)
    output["CAMPANHA"] = campanha
    output["CPF"] = df["CPF"].map(_normalize_cpf_key)
    output["Taxa"] = df["TAXA"].map(_format_br_decimal)
    output["FEZ EM"] = df["FEZ EM"].map(_format_int)
    output["FALTAM"] = df["FALTAM"].map(_format_int)
    output[" PARCELA "] = df["PARCELA"].map(_format_br_parcela)
    output["TROCO"] = df["TROCO"].map(_format_br_accounting)
    output["Telefone"] = df["TELEFONE BASE"].map(lambda value: _safe_text(value).strip())
    output["Nome"] = df["NOME"].map(lambda value: _safe_text(value).strip())
    output["Data Nascimento"] = df["DATA DE NASC"].map(_format_br_date)
    output["TELEFONE C6"] = df["TELEFONE C6"].map(_format_identifier_exact)
    return output[ARGUS_DISCADORA_COLUMNS]


def argus_discadora_csv_bytes(df: pd.DataFrame, campanha: str = "") -> bytes:
    return _argus_dataframe_to_csv_bytes(build_argus_discadora_dataframe(df, campanha))


def write_argus_discadora_csv(
    df: pd.DataFrame, output_path: str | Path, campanha: str = ""
) -> None:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(argus_discadora_csv_bytes(df, campanha))


def build_argus_discadora_from_file(
    source: str | Path | bytes | BinaryIO, campanha: str = ""
) -> pd.DataFrame:
    if not campanha:
        campanha = build_argus_campaign_name(source)
    return build_argus_discadora_dataframe(read_processed_c6_file(source), campanha)


def write_argus_discadora_csv_from_file(
    source: str | Path | bytes | BinaryIO, output_path: str | Path, campanha: str = ""
) -> None:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(argus_discadora_csv_bytes_from_file(source, campanha))


def argus_discadora_csv_bytes_from_file(
    source: str | Path | bytes | BinaryIO, campanha: str = ""
) -> bytes:
    return _argus_dataframe_to_csv_bytes(build_argus_discadora_from_file(source, campanha))


def build_argus_campaign_name(
    source: str | Path | bytes | BinaryIO,
    base_dir: str | Path | Iterable[str | Path] = ARGUS_DISC_BASE_DIRS,
) -> str:
    source_name = Path(str(getattr(source, "name", source))).stem
    uf = extract_uf_from_filename(source_name)
    next_number = find_next_argus_campaign_number(uf, base_dir)
    return f"{uf} {next_number} REFIN INSS C6"


def build_argus_output_filename(
    source: str | Path | bytes | BinaryIO,
    base_dir: str | Path | Iterable[str | Path] = ARGUS_DISC_BASE_DIRS,
) -> str:
    source_stem = _clean_source_stem(Path(str(getattr(source, "name", source))).stem)
    return f"{build_argus_campaign_name(source, base_dir)} {source_stem}.csv"


def build_argus_output_path(
    source: str | Path | bytes | BinaryIO,
    base_dir: str | Path | Iterable[str | Path] = ARGUS_DISC_BASE_DIRS,
    today: date | None = None,
) -> Path:
    source_name = Path(str(getattr(source, "name", source))).stem
    uf = extract_uf_from_filename(source_name)
    resolved_base_dir = resolve_argus_disc_base_dir(base_dir)
    return (
        resolved_base_dir
        / _argus_month_folder_name(today)
        / uf
        / build_argus_output_filename(source, resolved_base_dir)
    )


def build_lemit_processed_stem(source: str | Path | bytes | BinaryIO, row_count: int) -> str:
    source_name = Path(str(getattr(source, "name", source))).stem
    uf = extract_uf_from_filename(source_name)
    return f"{row_count} refin c6 {uf} (1 ou 0 cartões) sem RL pra disc"


def build_lemit_processed_xlsx_path(
    source: str | Path | bytes | BinaryIO, row_count: int
) -> Path:
    source_path = Path(str(getattr(source, "name", source)))
    return source_path.with_name(f"{build_lemit_processed_stem(source, row_count)}.xlsx")


def build_argus_output_filename_from_stem(
    source: str | Path | bytes | BinaryIO,
    row_count: int,
    base_dir: str | Path | Iterable[str | Path] = ARGUS_DISC_BASE_DIRS,
) -> str:
    processed_stem = build_lemit_processed_stem(source, row_count)
    return f"{build_argus_campaign_name(source, base_dir)} {processed_stem}.csv"


def build_argus_output_path_from_stem(
    source: str | Path | bytes | BinaryIO,
    row_count: int,
    base_dir: str | Path | Iterable[str | Path] = ARGUS_DISC_BASE_DIRS,
    today: date | None = None,
) -> Path:
    source_name = Path(str(getattr(source, "name", source))).stem
    uf = extract_uf_from_filename(source_name)
    resolved_base_dir = resolve_argus_disc_base_dir(base_dir)
    return (
        resolved_base_dir
        / _argus_month_folder_name(today)
        / uf
        / build_argus_output_filename_from_stem(source, row_count, resolved_base_dir)
    )


def extract_uf_from_filename(filename: str) -> str:
    normalized = _normalize_text(filename)
    tokens = [token for token in re.split(r"[^A-Z0-9]+", normalized) if token]
    for idx, token in enumerate(tokens[:-1]):
        if token == "C6" and tokens[idx + 1] in BRAZIL_UFS:
            return tokens[idx + 1]

    ufs = [token for token in tokens if token in BRAZIL_UFS]
    if not ufs:
        raise ValueError("Nao encontrei a UF do estado no nome do arquivo.")
    return ufs[-1]


def find_next_argus_campaign_number(
    uf: str, base_dir: str | Path | Iterable[str | Path] = ARGUS_DISC_BASE_DIRS
) -> int:
    base_path = resolve_argus_disc_base_dir(base_dir)

    pattern = re.compile(rf"\b{re.escape(uf)}\s+(\d+)\b")
    numbers: list[int] = []

    for path in base_path.rglob("*"):
        names = [path.name, *path.parts]
        if not any(_has_uf_token(name, uf) for name in names):
            continue
        for name in names:
            match = pattern.search(_normalize_text(name))
            if match:
                numbers.append(int(match.group(1)))

    return max(numbers, default=0) + 1


def resolve_argus_disc_base_dir(
    base_dir: str | Path | Iterable[str | Path] = ARGUS_DISC_BASE_DIRS,
) -> Path:
    if isinstance(base_dir, (str, Path)):
        candidates = (Path(base_dir),)
    else:
        candidates = tuple(Path(path) for path in base_dir)

    for candidate in candidates:
        if candidate.exists():
            return candidate

    tested_paths = " | ".join(str(path) for path in candidates)
    raise FileNotFoundError(f"Pasta da Argus nao encontrada. Caminhos testados: {tested_paths}")


def _argus_month_folder_name(today: date | None = None) -> str:
    current = today or date.today()
    return f"{MONTH_NAMES_PT_BR[current.month]} {current.year}"


def _argus_dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(
        index=False,
        sep=";",
        encoding="utf-8-sig",
        quoting=csv.QUOTE_ALL,
    ).encode("utf-8-sig")


def read_processed_c6_file(source: str | Path | bytes | BinaryIO) -> pd.DataFrame:
    source_name = str(getattr(source, "name", ""))
    if isinstance(source, (str, Path)):
        source_name = str(source)

    if source_name.lower().endswith(".xlsx"):
        raw = _read_visible_xlsx(source)
    elif source_name.lower().endswith(".xls"):
        raw = pd.read_excel(source, dtype=str)
    else:
        data = _read_bytes(source)
        last_error: Exception | None = None
        raw = None
        for encoding in ("utf-8-sig", "cp1252", "latin1"):
            try:
                text = data.decode(encoding)
                raw = pd.read_csv(io.StringIO(text), sep=None, engine="python", dtype=str)
                break
            except Exception as exc:
                last_error = exc
        if raw is None:
            raise ValueError(f"Nao consegui ler a planilha para Argus. Ultimo erro: {last_error}")

    raw.columns = [str(column).strip().lstrip("\ufeff") for column in raw.columns]
    columns = {_normalize_header(column): column for column in raw.columns}
    aliases = {
        "CPF": "CPF",
        "TAXA": "TAXA",
        "FEZEM": "FEZ EM",
        "FALTAM": "FALTAM",
        "PARCELA": "PARCELA",
        "TROCO": "TROCO",
        "NOME": "NOME",
        "DATADENASC": "DATA DE NASC",
        "DATANASCIMENTO": "DATA DE NASC",
        "TELEFONEC6": "TELEFONE C6",
        "TELEFONEBASE": "TELEFONE BASE",
    }
    required = (
        "CPF",
        "TAXA",
        "FEZ EM",
        "FALTAM",
        "PARCELA",
        "TROCO",
        "NOME",
        "DATA DE NASC",
        "TELEFONE C6",
        "TELEFONE BASE",
    )

    output = pd.DataFrame()
    for normalized, target in aliases.items():
        if normalized in columns:
            output[target] = raw[columns[normalized]]

    missing = [column for column in required if column not in output.columns]
    if missing:
        raise ValueError(
            "Planilha para Argus sem coluna obrigatoria: " + ", ".join(missing)
        )

    output["CPF"] = output["CPF"].map(_normalize_cpf_key)
    output["TAXA"] = output["TAXA"].map(_parse_br_number)
    output["FEZ EM"] = output["FEZ EM"].map(_parse_br_number).round().astype("Int64")
    output["FALTAM"] = output["FALTAM"].map(_parse_br_number).round().astype("Int64")
    output["PARCELA"] = output["PARCELA"].map(_parse_br_number)
    output["TROCO"] = output["TROCO"].map(_parse_br_number)
    output["NOME"] = output["NOME"].fillna("").astype(str).str.strip()
    output["DATA DE NASC"] = output["DATA DE NASC"].map(_parse_br_date_value)
    output["TELEFONE C6"] = output["TELEFONE C6"].map(_format_identifier_exact)
    output["TELEFONE BASE"] = output["TELEFONE BASE"].map(_format_identifier_exact)
    return output[list(required)]


def _read_visible_xlsx(source: str | Path | bytes | BinaryIO) -> pd.DataFrame:
    from openpyxl import load_workbook

    workbook_source = source
    if not isinstance(source, (str, Path)):
        workbook_source = io.BytesIO(_read_bytes(source))

    workbook = load_workbook(workbook_source, data_only=True, read_only=False)
    worksheet = workbook.active

    visible_rows = []
    for excel_row_number, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
        if worksheet.row_dimensions[excel_row_number].hidden:
            continue
        visible_rows.append(list(row))

    if not visible_rows:
        return pd.DataFrame()

    header_index = None
    for index, row in enumerate(visible_rows):
        if any(value not in (None, "") for value in row):
            header_index = index
            break
    if header_index is None:
        return pd.DataFrame()

    headers = ["" if value is None else str(value).strip() for value in visible_rows[header_index]]
    data_rows = visible_rows[header_index + 1 :]
    return pd.DataFrame(data_rows, columns=headers, dtype=str).dropna(how="all")


def build_filtered_from_csv(
    source: str | Path | bytes | BinaryIO,
    options: FilterOptions,
    phone_base_dir: str | Path | Iterable[str | Path] | None = None,
    lemit_source: str | Path | bytes | BinaryIO | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    raw = read_c6_csv(source)
    transformed = transform_c6_dataframe(raw)
    filtered = apply_filters(transformed, options)
    if phone_base_dir is not None or lemit_source is not None:
        filtered = enrich_with_phone_sources(
            filtered, phone_base_dir, lemit_source, options.telefone_data_inicial
        )
    return transformed, filtered


def enrich_with_phone_base(df: pd.DataFrame, phone_base_dir: str | Path) -> pd.DataFrame:
    return enrich_with_phone_sources(df, phone_base_dir, None)


def enrich_with_phone_sources(
    df: pd.DataFrame,
    phone_base_dir: str | Path | Iterable[str | Path] | None = None,
    lemit_source: str | Path | bytes | BinaryIO | None = None,
    telefone_data_inicial: date | None = None,
) -> pd.DataFrame:
    output = df.copy()
    if output.empty:
        output["TELEFONE BASE"] = ""
        return output

    lookups: list[dict[str, tuple[str, ...]]] = []
    if phone_base_dir is not None:
        lookups.append(
            load_phone_lookup_for_cpfs(output["CPF"], phone_base_dir, telefone_data_inicial)
        )
    if lemit_source is not None:
        lookups.append(load_lemit_phone_lookup(lemit_source))

    phone_lookup = _merge_phone_lookups(*lookups)
    cpf_keys = output["CPF"].map(_normalize_cpf_key)
    output["TELEFONE BASE"] = cpf_keys.map(
        lambda cpf: " / ".join(phone_lookup.get(cpf, ()))
    ).fillna("")
    return output


def load_phone_lookup_for_cpfs(
    cpfs: Iterable[object],
    phone_base_dir: str | Path | Iterable[str | Path],
    telefone_data_inicial: date | None = None,
) -> dict[str, tuple[str, ...]]:
    base_dirs = resolve_phone_base_dirs(phone_base_dir)

    target_cpfs = {_normalize_cpf_key(cpf) for cpf in cpfs}
    target_cpfs.discard("")
    phone_lists: dict[str, list[str]] = {}
    phone_sets: dict[str, set[str]] = {}

    for base_dir in base_dirs:
        for csv_path in sorted(base_dir.glob("*.csv")):
            for chunk in _read_phone_base_chunks(csv_path, telefone_data_inicial is not None):
                chunk.columns = [str(column).strip() for column in chunk.columns]
                normalized_columns = {_normalize_header(column): column for column in chunk.columns}
                cpf_column = normalized_columns.get("CPF")
                phone_column = normalized_columns.get("TELEFONE")
                date_column = normalized_columns.get("DATADAULTIMAMODIFICACAO")
                if cpf_column is None or phone_column is None:
                    continue

                cpf_series = chunk[cpf_column].map(_normalize_cpf_key)
                phone_series = chunk[phone_column].map(_normalize_phone)
                matched = pd.DataFrame({"cpf": cpf_series, "phone": phone_series})
                if telefone_data_inicial is not None and date_column is not None:
                    date_series = chunk[date_column].map(_parse_br_date_value)
                    min_date = pd.Timestamp(telefone_data_inicial)
                    matched = matched[date_series.isna() | (date_series >= min_date)]
                matched = matched[
                    matched["cpf"].isin(target_cpfs) & (matched["phone"] != "")
                ]

                for cpf, phone in matched.itertuples(index=False):
                    seen = phone_sets.setdefault(cpf, set())
                    if phone in seen:
                        continue
                    seen.add(phone)
                    phone_lists.setdefault(cpf, []).append(phone)

    return {cpf: tuple(phones) for cpf, phones in phone_lists.items()}


def resolve_phone_base_dirs(phone_base_dir: str | Path | Iterable[str | Path]) -> tuple[Path, ...]:
    if isinstance(phone_base_dir, (str, Path)):
        candidates = (Path(phone_base_dir),)
    else:
        candidates = tuple(Path(path) for path in phone_base_dir)

    existing = tuple(candidate for candidate in candidates if candidate.exists())
    if existing:
        return existing

    tested_paths = " | ".join(str(path) for path in candidates)
    raise FileNotFoundError(
        f"Pasta da base de telefones não encontrada. Caminhos testados: {tested_paths}"
    )


def resolve_phone_base_dir(phone_base_dir: str | Path | Iterable[str | Path]) -> Path:
    return resolve_phone_base_dirs(phone_base_dir)[0]

    if isinstance(phone_base_dir, (str, Path)):
        candidates = (Path(phone_base_dir),)
    else:
        candidates = tuple(Path(path) for path in phone_base_dir)

    for candidate in candidates:
        if candidate.exists():
            return candidate

    tested_paths = " | ".join(str(path) for path in candidates)
    raise FileNotFoundError(
        f"Pasta da base de telefones não encontrada. Caminhos testados: {tested_paths}"
    )


def load_lemit_phone_lookup(source: str | Path | bytes | BinaryIO) -> dict[str, tuple[str, ...]]:
    df = read_lemit_file(source)
    normalized_columns = {_normalize_header(column): column for column in df.columns}
    cpf_column = _required_column(normalized_columns, "CPF")
    phone_column = _required_column(normalized_columns, "TELEFONE")
    ddd_column = normalized_columns.get("DDD")
    ranking_column = normalized_columns.get("RANKING")
    score_column = normalized_columns.get("SCORE")

    work_df = df.copy()
    work_df["_cpf"] = work_df[cpf_column].map(_normalize_cpf_key)
    if ddd_column is not None:
        ddd_series = work_df[ddd_column].map(_digits_only)
    else:
        ddd_series = pd.Series([""] * len(work_df), index=work_df.index)
    work_df["_phone"] = [
        _normalize_lemit_phone(ddd, phone)
        for ddd, phone in zip(ddd_series, work_df[phone_column], strict=False)
    ]

    if ranking_column is not None:
        work_df["_ranking"] = pd.to_numeric(work_df[ranking_column], errors="coerce")
    else:
        work_df["_ranking"] = pd.NA
    if score_column is not None:
        work_df["_score"] = pd.to_numeric(work_df[score_column], errors="coerce")
    else:
        work_df["_score"] = pd.NA

    work_df = work_df[(work_df["_cpf"] != "") & (work_df["_phone"] != "")]
    work_df = work_df.sort_values(
        by=["_cpf", "_ranking", "_score"],
        ascending=[True, True, False],
        na_position="last",
        kind="stable",
    )

    phone_lists: dict[str, list[str]] = {}
    phone_sets: dict[str, set[str]] = {}
    for cpf, phone in work_df[["_cpf", "_phone"]].itertuples(index=False):
        seen = phone_sets.setdefault(cpf, set())
        if phone in seen:
            continue
        seen.add(phone)
        phone_lists.setdefault(cpf, []).append(phone)

    return {cpf: tuple(phones) for cpf, phones in phone_lists.items()}


def filter_to_lemit_cpfs(
    df: pd.DataFrame, lemit_source: str | Path | bytes | BinaryIO
) -> pd.DataFrame:
    lemit_df = read_lemit_file(lemit_source)
    normalized_columns = {_normalize_header(column): column for column in lemit_df.columns}
    cpf_column = _required_column(normalized_columns, "CPF")
    lemit_cpfs = {
        cpf
        for cpf in lemit_df[cpf_column].map(_normalize_cpf_key)
        if cpf
    }
    cpf_keys = df["CPF"].map(_normalize_cpf_key)
    return df.loc[cpf_keys.isin(lemit_cpfs)].copy().reset_index(drop=True)


def read_lemit_file(source: str | Path | bytes | BinaryIO) -> pd.DataFrame:
    source_name = str(getattr(source, "name", ""))
    if isinstance(source, (str, Path)):
        source_name = str(source)

    if source_name.lower().endswith(".xlsx"):
        return _read_visible_xlsx(source)
    if source_name.lower().endswith(".xls"):
        return pd.read_excel(source, dtype=str)

    data = _read_bytes(source)
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "cp1252", "latin1"):
        try:
            text = data.decode(encoding)
            df = pd.read_csv(io.StringIO(text), sep=None, engine="python", dtype=str)
            df.columns = [str(column).strip().lstrip("\ufeff") for column in df.columns]
            return df
        except Exception as exc:
            last_error = exc

    raise ValueError(f"Não consegui ler o retorno Lemiti. Último erro: {last_error}")


def _merge_phone_lookups(
    *lookups: dict[str, tuple[str, ...]]
) -> dict[str, tuple[str, ...]]:
    merged_lists: dict[str, list[str]] = {}
    merged_sets: dict[str, set[str]] = {}
    for lookup in lookups:
        for cpf, phones in lookup.items():
            seen = merged_sets.setdefault(cpf, set())
            current = merged_lists.setdefault(cpf, [])
            for phone in phones:
                if phone in seen:
                    continue
                seen.add(phone)
                current.append(phone)

    return {cpf: tuple(phones) for cpf, phones in merged_lists.items()}


def get_missing_phone_cpfs(df: pd.DataFrame) -> pd.DataFrame:
    if "TELEFONE BASE" not in df.columns:
        missing = df
    else:
        phone_text = df["TELEFONE BASE"].fillna("").astype(str).str.strip()
        missing = df.loc[phone_text == ""]

    cpfs = (
        missing["CPF"]
        .map(_normalize_cpf_key)
        .loc[lambda series: series != ""]
        .drop_duplicates()
    )
    return pd.DataFrame({"cpf": cpfs})


def missing_phone_cpfs_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return get_missing_phone_cpfs(df).to_csv(index=False, encoding="utf-8-sig").encode(
        "utf-8-sig"
    )


def write_missing_phone_cpfs_csv(df: pd.DataFrame, output_path: str | Path) -> None:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    get_missing_phone_cpfs(df).to_csv(output_path, index=False, encoding="utf-8-sig")


def _read_phone_base_chunks(csv_path: Path, include_modified_date: bool = False):
    last_error: Exception | None = None
    wanted_columns = {"CPF", "TELEFONE"}
    if include_modified_date:
        wanted_columns.add("DATADAULTIMAMODIFICACAO")

    for encoding in ("utf-8-sig", "cp1252", "latin1"):
        try:
            chunks = pd.read_csv(
                csv_path,
                sep=";",
                dtype=str,
                encoding=encoding,
                usecols=lambda column: _normalize_header(column) in wanted_columns,
                chunksize=200_000,
                on_bad_lines="skip",
            )
            for chunk in chunks:
                yield chunk
            return
        except Exception as exc:
            last_error = exc

    raise ValueError(f"Não consegui ler a base de telefones {csv_path}: {last_error}")


def _required_column(columns: dict[str, str], normalized_name: str) -> str:
    if normalized_name not in columns:
        raise ValueError(f"Retorno Lemiti sem coluna obrigatória: {normalized_name}")
    return columns[normalized_name]


def _normalize_header(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value).strip())
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"\s+", "", text).upper()


def _normalize_text(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value).strip())
    text = "".join(char for char in text if not unicodedata.combining(char))
    return text.upper()


def _has_uf_token(value: object, uf: str) -> bool:
    tokens = [token for token in re.split(r"[^A-Z0-9]+", _normalize_text(value)) if token]
    return uf in tokens


def _clean_source_stem(stem: str) -> str:
    return re.sub(r"\s*\(\d+\)\s*$", "", stem).strip()


def _read_bytes(source: str | Path | bytes | BinaryIO) -> bytes:
    if isinstance(source, bytes):
        return source

    if isinstance(source, (str, Path)):
        return Path(source).read_bytes()

    if hasattr(source, "getvalue"):
        return source.getvalue()

    return source.read()


def _digits_only(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\D+", "", str(value).strip())


def _normalize_cpf_key(value: object) -> str:
    digits = _digits_only(value)
    if not digits:
        return ""
    return digits[-11:].zfill(11)


def _normalize_phone(value: object) -> str:
    digits = _digits_only(value)
    if len(digits) < 8:
        return ""
    return digits


def _normalize_lemit_phone(ddd: object, phone: object) -> str:
    ddd_digits = _digits_only(ddd)
    phone_digits = _digits_only(phone)
    if len(phone_digits) < 8:
        return ""
    if len(phone_digits) <= 9 and ddd_digits:
        return ddd_digits[-2:] + phone_digits
    return phone_digits


def _parse_br_number(value: object) -> float | None:
    if pd.isna(value):
        return None

    text = str(value).strip()
    if not text:
        return None

    text = text.replace("R$", "").replace("\xa0", "").replace(" ", "")
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    elif text.count(".") > 1:
        parts = text.split(".")
        text = "".join(parts[:-1]) + "." + parts[-1]

    try:
        return float(text)
    except ValueError:
        return None


def _parse_int_series(series: pd.Series) -> pd.Series:
    return series.map(_parse_br_number).round().astype("Int64")


def _safe_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value)


def _format_int(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(int(float(value)))


def _format_br_decimal(value: object) -> str:
    if pd.isna(value):
        return ""
    return f"{float(value):.2f}".replace(".", ",")


def _format_br_parcela(value: object) -> str:
    if pd.isna(value):
        return ""

    number = float(value)
    if number.is_integer() and abs(number) < 1000:
        return str(int(number))

    formatted = f"{number:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    if "," in formatted:
        formatted = formatted.rstrip("0").rstrip(",")
    if number.is_integer() and abs(number) >= 1000:
        return formatted + ",00"
    return formatted


def _format_br_accounting(value: object) -> str:
    if pd.isna(value):
        return ""
    number = float(value)
    formatted = f"{abs(number):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    if number < 0:
        return f"R$ ({formatted})"
    return f"R$ {formatted}"


def _format_br_date(value: object) -> str:
    if pd.isna(value):
        return ""
    parsed = _parse_br_date_value(value)
    if pd.isna(parsed):
        return ""
    return pd.to_datetime(parsed).strftime("%d/%m/%Y")


def _parse_br_date_value(value: object):
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value

    if isinstance(value, datetime):
        return pd.Timestamp(value)

    if isinstance(value, date):
        return pd.Timestamp(value.year, value.month, value.day)

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")
        except Exception:
            return pd.NaT

    text = str(value).strip()
    if not text:
        return pd.NaT

    for dayfirst in (True, False):
        parsed = pd.to_datetime(text, dayfirst=dayfirst, errors="coerce")
        if not pd.isna(parsed):
            return parsed

    return pd.NaT


def _format_identifier_exact(value: object) -> str:
    if pd.isna(value):
        return ""

    if isinstance(value, int):
        return str(value)

    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.15g}"

    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def _number_or_blank(value: object) -> float | str:
    if pd.isna(value):
        return ""
    return float(value)


def _write_number(worksheet, row: int, col: int, value: object, cell_format) -> None:
    if pd.isna(value):
        worksheet.write_blank(row, col, None, cell_format)
    else:
        worksheet.write_number(row, col, float(value), cell_format)


def _write_digits_number(worksheet, row: int, col: int, value: object, cell_format) -> None:
    digits = _digits_only(value)
    if not digits:
        worksheet.write_blank(row, col, None, cell_format)
    else:
        worksheet.write_number(row, col, int(digits), cell_format)


def _write_date(worksheet, row: int, col: int, value: object, cell_format) -> None:
    if pd.isna(value):
        worksheet.write_blank(row, col, None, cell_format)
        return

    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()
    elif isinstance(value, date) and not isinstance(value, datetime):
        value = datetime(value.year, value.month, value.day)

    worksheet.write_datetime(row, col, value, cell_format)


def _parse_cli_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Filtra exportação C6 e gera XLSX formatado.")
    parser.add_argument("--input", required=True, help="Caminho da Extração DTF CSV.")
    parser.add_argument("--output", required=True, help="Caminho do XLSX final.")
    parser.add_argument(
        "--phone-base-dir",
        default="",
        help=(
            "Pasta dos CSVs da base de telefones. Se ficar vazio, tenta G: e depois H:."
        ),
    )
    parser.add_argument("--lemit-input", default="", help="Caminho do retorno Lemiti.")
    parser.add_argument("--faltam", default="", help="Prazo(s) FALTAM. Ex.: 67 ou 67,72.")
    parser.add_argument(
        "--faltam-modo",
        default="maior",
        choices=["maior", "a_partir", "ate", "exato"],
        help="Como aplicar o filtro FALTAM.",
    )
    parser.add_argument("--taxa-min", type=float, default=0.0, help="Taxa mínima.")
    parser.add_argument("--troco-min", type=float, default=0.0, help="Troco mínimo, filtro maior que.")
    parser.add_argument("--data-inicial", default="01/01/1954", help="Data inicial.")
    parser.add_argument("--data-final", default="31/12/1995", help="Data final.")
    return parser.parse_args()


def main() -> None:
    args = _parse_cli_args()
    options = FilterOptions(
        faltam=parse_faltam_values(args.faltam),
        faltam_modo=args.faltam_modo,
        taxa_min=args.taxa_min,
        troco_min=args.troco_min,
        data_inicial=parse_date_value(args.data_inicial),
        data_final=parse_date_value(args.data_final),
    )
    lemit_input = args.lemit_input or None
    _, filtered = build_filtered_from_csv(
        args.input,
        options,
        args.phone_base_dir or DEFAULT_PHONE_BASE_DIRS,
        lemit_input,
    )
    write_xlsx(filtered, args.output)
    print(f"Linhas exportadas: {len(filtered)}")
    print(f"Arquivo: {Path(args.output).resolve()}")


if __name__ == "__main__":
    main()
