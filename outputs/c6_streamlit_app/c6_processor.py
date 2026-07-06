from __future__ import annotations

import argparse
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


@dataclass(frozen=True)
class FilterOptions:
    faltam: tuple[int, ...] = ()
    faltam_modo: str = "maior"
    taxa_min: float = 0.0
    troco_min: float = 0.0
    data_inicial: date | None = date(1954, 1, 1)
    data_final: date | None = date(1995, 12, 31)


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


def build_filtered_from_csv(
    source: str | Path | bytes | BinaryIO,
    options: FilterOptions,
    phone_base_dir: str | Path | None = None,
    lemit_source: str | Path | bytes | BinaryIO | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    raw = read_c6_csv(source)
    transformed = transform_c6_dataframe(raw)
    filtered = apply_filters(transformed, options)
    if phone_base_dir is not None or lemit_source is not None:
        filtered = enrich_with_phone_sources(filtered, phone_base_dir, lemit_source)
    return transformed, filtered


def enrich_with_phone_base(df: pd.DataFrame, phone_base_dir: str | Path) -> pd.DataFrame:
    return enrich_with_phone_sources(df, phone_base_dir, None)


def enrich_with_phone_sources(
    df: pd.DataFrame,
    phone_base_dir: str | Path | None = None,
    lemit_source: str | Path | bytes | BinaryIO | None = None,
) -> pd.DataFrame:
    output = df.copy()
    if output.empty:
        output["TELEFONE BASE"] = ""
        return output

    lookups: list[dict[str, tuple[str, ...]]] = []
    if phone_base_dir is not None:
        lookups.append(load_phone_lookup_for_cpfs(output["CPF"], phone_base_dir))
    if lemit_source is not None:
        lookups.append(load_lemit_phone_lookup(lemit_source))

    phone_lookup = _merge_phone_lookups(*lookups)
    cpf_keys = output["CPF"].map(_normalize_cpf_key)
    output["TELEFONE BASE"] = cpf_keys.map(
        lambda cpf: " / ".join(phone_lookup.get(cpf, ()))
    ).fillna("")
    return output


def load_phone_lookup_for_cpfs(
    cpfs: Iterable[object], phone_base_dir: str | Path
) -> dict[str, tuple[str, ...]]:
    base_dir = Path(phone_base_dir)
    if not base_dir.exists():
        raise FileNotFoundError(f"Pasta da base de telefones não encontrada: {base_dir}")

    target_cpfs = {_normalize_cpf_key(cpf) for cpf in cpfs}
    target_cpfs.discard("")
    phone_lists: dict[str, list[str]] = {}
    phone_sets: dict[str, set[str]] = {}

    for csv_path in sorted(base_dir.glob("*.csv")):
        for chunk in _read_phone_base_chunks(csv_path):
            chunk.columns = [str(column).strip().upper() for column in chunk.columns]
            if "CPF" not in chunk.columns or "TELEFONE" not in chunk.columns:
                continue

            cpf_series = chunk["CPF"].map(_normalize_cpf_key)
            phone_series = chunk["TELEFONE"].map(_normalize_phone)
            matched = pd.DataFrame({"cpf": cpf_series, "phone": phone_series})
            matched = matched[matched["cpf"].isin(target_cpfs) & (matched["phone"] != "")]

            for cpf, phone in matched.itertuples(index=False):
                seen = phone_sets.setdefault(cpf, set())
                if phone in seen:
                    continue
                seen.add(phone)
                phone_lists.setdefault(cpf, []).append(phone)

    return {cpf: tuple(phones) for cpf, phones in phone_lists.items()}


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


def read_lemit_file(source: str | Path | bytes | BinaryIO) -> pd.DataFrame:
    source_name = str(getattr(source, "name", ""))
    if isinstance(source, (str, Path)):
        source_name = str(source)

    if source_name.lower().endswith((".xlsx", ".xls")):
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


def _read_phone_base_chunks(csv_path: Path):
    last_error: Exception | None = None
    wanted_columns = {"CPF", "TELEFONE"}

    for encoding in ("utf-8-sig", "cp1252", "latin1"):
        try:
            chunks = pd.read_csv(
                csv_path,
                sep=";",
                dtype=str,
                encoding=encoding,
                usecols=lambda column: str(column).strip().upper() in wanted_columns,
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
        default=str(DEFAULT_PHONE_BASE_DIR),
        help="Pasta dos CSVs da base de telefones.",
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
        args.input, options, args.phone_base_dir, lemit_input
    )
    write_xlsx(filtered, args.output)
    print(f"Linhas exportadas: {len(filtered)}")
    print(f"Arquivo: {Path(args.output).resolve()}")


if __name__ == "__main__":
    main()
