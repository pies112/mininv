from __future__ import annotations

import ast
import csv
import json
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

import numpy as np
import pandas as pd

try:
    import pycountry
except ImportError as exc:
    raise ImportError("Установите pycountry: pip install pycountry") from exc


VERSION = "4"
PROJECT_ROOT = Path(
    "/Users/vsevolodkarass/Library/Mobile Documents/com~apple~CloudDocs/"
    "Desktop/Рабочий стол — MacBook Air — Всеволод/"
    "Инженерный проект/Версия рейтинга 1"
)

INPUT_FILENAME = "3_patents_2020_2025_final.xlsx"
OUTPUT_DIRNAME = "4_Финальная патентная база 2020-2025"
OUTPUT_FILENAME = "4_patents_2020_2025_final.xlsx"
OUTPUT_PARQUET_FILENAME = "4_patents_2020_2025_final.parquet"


@dataclass(frozen=True)
class SourceSpec:
    object_type: str
    path: Path
    delimiter: str


SOURCE_SPECS = [
    SourceSpec(
        "ТИМС",
        Path(
            "/Users/vsevolodkarass/Library/Mobile Documents/com~apple~CloudDocs/"
            "Desktop/Рабочий стол — MacBook Air — Всеволод/"
            "Инженерный проект/Версия рейтинга 1/Выгрузка данных/Роспатент/"
            "Открытый реестр топологий интегральных микросхем.csv"
        ),
        ",",
    ),
    SourceSpec(
        "Промышленный образец",
        Path(
            "/Users/vsevolodkarass/Library/Mobile Documents/com~apple~CloudDocs/"
            "Desktop/Рабочий стол — MacBook Air — Всеволод/"
            "Инженерный проект/Версия рейтинга 1/Выгрузка данных/Роспатент/"
            "Открытый реестр промышленных образцов Российской Федерации.csv"
        ),
        ";",
    ),
    SourceSpec(
        "Программа для ЭВМ",
        Path(
            "/Users/vsevolodkarass/Library/Mobile Documents/com~apple~CloudDocs/"
            "Desktop/Рабочий стол — MacBook Air — Всеволод/"
            "Инженерный проект/Версия рейтинга 1/Выгрузка данных/Роспатент/"
            "Открытый реестр программ для электронно-вычислительных машин.csv"
        ),
        ";",
    ),
    SourceSpec(
        "Полезная модель",
        Path(
            "/Users/vsevolodkarass/Library/Mobile Documents/com~apple~CloudDocs/"
            "Desktop/Рабочий стол — MacBook Air — Всеволод/"
            "Инженерный проект/Версия рейтинга 1/Выгрузка данных/Роспатент/"
            "Открытый реестр полезных моделей Российской Федерации.csv"
        ),
        ";",
    ),
    SourceSpec(
        "Изобретение",
        Path(
            "/Users/vsevolodkarass/Library/Mobile Documents/com~apple~CloudDocs/"
            "Desktop/Рабочий стол — MacBook Air — Всеволод/"
            "Инженерный проект/Версия рейтинга 1/Выгрузка данных/Роспатент/"
            "Открытый реестр изобретений Российской Федерации.csv"
        ),
        ";",
    ),
    SourceSpec(
        "База данных",
        Path(
            "/Users/vsevolodkarass/Library/Mobile Documents/com~apple~CloudDocs/"
            "Desktop/Рабочий стол — MacBook Air — Всеволод/"
            "Инженерный проект/Версия рейтинга 1/Выгрузка данных/Роспатент/"
            "Открытый реестр баз данных.csv"
        ),
        ";",
    ),
]

CLASSIC_TYPES = {"Изобретение", "Полезная модель", "Промышленный образец"}
DIGITAL_TYPES = {"Программа для ЭВМ", "База данных", "ТИМС"}
ROUTE_CODES = {"WO", "EP", "EA", "AP", "OA", "GC", "EM", "IB", "PCT"}

ISO2_CODES = {country.alpha_2 for country in pycountry.countries}
ISO2_CODES.update({"XK"})
CODE_ALIASES = {"UK": "GB", "EL": "GR"}

COUNTRY_NAME_ALIASES = {
    "РОССИЯ": "RU",
    "РОССИЙСКАЯ ФЕДЕРАЦИЯ": "RU",
    "RUSSIA": "RU",
    "RUSSIAN FEDERATION": "RU",
    "США": "US",
    "СОЕДИНЕННЫЕ ШТАТЫ": "US",
    "UNITED STATES": "US",
    "UNITED STATES OF AMERICA": "US",
    "КИТАЙ": "CN",
    "КНР": "CN",
    "CHINA": "CN",
    "ГЕРМАНИЯ": "DE",
    "GERMANY": "DE",
    "ФРАНЦИЯ": "FR",
    "FRANCE": "FR",
    "ВЕЛИКОБРИТАНИЯ": "GB",
    "СОЕДИНЕННОЕ КОРОЛЕВСТВО": "GB",
    "UNITED KINGDOM": "GB",
    "ИРЛАНДИЯ": "IE",
    "IRELAND": "IE",
    "КАЗАХСТАН": "KZ",
    "KAZAKHSTAN": "KZ",
    "БЕЛАРУСЬ": "BY",
    "БЕЛОРУССИЯ": "BY",
    "BELARUS": "BY",
    "АРМЕНИЯ": "AM",
    "ARMENIA": "AM",
    "КЫРГЫЗСТАН": "KG",
    "КИРГИЗИЯ": "KG",
    "KYRGYZSTAN": "KG",
    "УЗБЕКИСТАН": "UZ",
    "UZBEKISTAN": "UZ",
    "УКРАИНА": "UA",
    "UKRAINE": "UA",
    "НИДЕРЛАНДЫ": "NL",
    "NETHERLANDS": "NL",
    "ШВЕЙЦАРИЯ": "CH",
    "SWITZERLAND": "CH",
    "ЯПОНИЯ": "JP",
    "JAPAN": "JP",
    "РЕСПУБЛИКА КОРЕЯ": "KR",
    "ЮЖНАЯ КОРЕЯ": "KR",
    "SOUTH KOREA": "KR",
    "REPUBLIC OF KOREA": "KR",
    "ИНДИЯ": "IN",
    "INDIA": "IN",
    "ИТАЛИЯ": "IT",
    "ITALY": "IT",
    "ИСПАНИЯ": "ES",
    "SPAIN": "ES",
    "ИЗРАИЛЬ": "IL",
    "ISRAEL": "IL",
    "КАНАДА": "CA",
    "CANADA": "CA",
    "АВСТРАЛИЯ": "AU",
    "AUSTRALIA": "AU",
    "АВСТРИЯ": "AT",
    "AUSTRIA": "AT",
    "ШВЕЦИЯ": "SE",
    "SWEDEN": "SE",
    "ФИНЛЯНДИЯ": "FI",
    "FINLAND": "FI",
    "НОРВЕГИЯ": "NO",
    "NORWAY": "NO",
    "ДАНИЯ": "DK",
    "DENMARK": "DK",
    "ПОЛЬША": "PL",
    "POLAND": "PL",
    "ЧЕХИЯ": "CZ",
    "CZECHIA": "CZ",
    "CZECH REPUBLIC": "CZ",
    "СЕРБИЯ": "RS",
    "SERBIA": "RS",
    "ТУРЦИЯ": "TR",
    "TURKEY": "TR",
    "TÜRKIYE": "TR",
    "СИНГАПУР": "SG",
    "SINGAPORE": "SG",
    "ТАЙВАНЬ": "TW",
    "TAIWAN": "TW",
}

IPC_FULL_RE = re.compile(r"\b([A-HY]\d{2}[A-Z]\s*\d{1,4}(?:/\d{1,8})?)\b", re.I)
IPC_SUBCLASS_RE = re.compile(r"\b([A-HY]\d{2}[A-Z])\b", re.I)
LOCARNO_RE = re.compile(r"\b(\d{1,2}[-.]\d{1,2})\b")


SHEET_ALIASES = {
    "Изобретение": ("изобрет",),
    "Полезная модель": ("полез",),
    "Промышленный образец": ("промышлен",),
    "Программа для ЭВМ": ("программ", "эвм"),
    "База данных": ("баз", "данн"),
    "ТИМС": ("тимс", "тополог"),
}


def norm_text(value: Any) -> str:
    return unicodedata.normalize("NFC", str(value)).replace("\xa0", " ").strip()


def norm_name(value: Any) -> str:
    return re.sub(r"[^a-zа-яё0-9]+", "_", norm_text(value).casefold()).strip("_")


def resolve_macos_path(path: Path) -> Path:
    path = Path(path)
    if path.exists():
        return path
    parts = path.parts
    if not parts:
        return path
    current = Path(parts[0]) if path.is_absolute() else Path(".")
    start = 1 if path.is_absolute() else 0
    for part in parts[start:]:
        if not current.exists() or not current.is_dir():
            return path
        target = norm_text(part).casefold()
        match = next((item for item in current.iterdir() if norm_text(item.name).casefold() == target), None)
        if match is None:
            return path
        current = match
    return current


def find_final_workbook() -> Path:
    candidates = [
        PROJECT_ROOT / "Выгрузка данных" / "3_Финальная патентная база 2020-2025" / INPUT_FILENAME,
        PROJECT_ROOT / "Выгрузка данных" / INPUT_FILENAME,
        PROJECT_ROOT / INPUT_FILENAME,
    ]
    for candidate in candidates:
        resolved = resolve_macos_path(candidate)
        if resolved.exists():
            return resolved
    target = norm_text(INPUT_FILENAME).casefold()
    for root in [PROJECT_ROOT / "Выгрузка данных", PROJECT_ROOT, Path.cwd()]:
        root = resolve_macos_path(root)
        if not root.exists():
            continue
        for item in root.rglob("*.xlsx"):
            if norm_text(item.name).casefold() == target:
                return item
    raise FileNotFoundError(f"Не найден {INPUT_FILENAME} внутри проекта.")


def find_col(df: pd.DataFrame, candidates: Iterable[str], required: bool = False) -> str | None:
    mapping = {norm_name(column): column for column in df.columns}
    for candidate in candidates:
        key = norm_name(candidate)
        if key in mapping:
            return mapping[key]
    if required:
        raise KeyError(f"Не найден столбец из вариантов: {list(candidates)}")
    return None


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip().casefold() in {"", "nan", "none", "null", "[]", "{}"}


def parse_nested(value: Any) -> Any:
    if is_blank(value):
        return None
    if isinstance(value, (dict, list, tuple, set)):
        return value
    text = str(value).strip()
    for parser in (json.loads, ast.literal_eval):
        try:
            return parser(text)
        except Exception:
            continue
    return text


def flatten_text(value: Any) -> str:
    parsed = parse_nested(value)
    if parsed is None:
        return ""
    if isinstance(parsed, dict):
        return " ".join(flatten_text(item) for item in parsed.values())
    if isinstance(parsed, (list, tuple, set)):
        return " ".join(flatten_text(item) for item in parsed)
    return str(parsed)


def normalize_country_code(code: str) -> str | None:
    code = str(code).strip().upper()
    code = CODE_ALIASES.get(code, code)
    if code in ISO2_CODES:
        return code
    return None


def country_codes_from_value(value: Any, column_name: str = "") -> set[str]:
    result: set[str] = set()
    parsed = parse_nested(value)
    column_norm = norm_name(column_name)
    context_is_country = any(token in column_norm for token in ("country", "countries", "страна", "страны", "nationality"))

    def visit(item: Any, parent_key: str = "") -> None:
        if item is None:
            return
        if isinstance(item, dict):
            for key, val in item.items():
                key_norm = norm_name(key)
                if any(token in key_norm for token in ("country", "countries", "страна", "страны", "nationality", "гражданство")):
                    visit_country_value(val)
                else:
                    visit(val, key_norm)
            return
        if isinstance(item, (list, tuple, set)):
            for val in item:
                visit(val, parent_key)
            return
        text = str(item)
        for code in re.findall(r"\(([A-Z]{2})\)|\[([A-Z]{2})\]", text.upper()):
            for candidate in code:
                if candidate:
                    normalized = normalize_country_code(candidate)
                    if normalized:
                        result.add(normalized)
        for match in re.findall(r"(?i)(?:country|country_code|страна|код страны)\s*[:=]\s*[\"']?([A-Z]{2})", text):
            normalized = normalize_country_code(match)
            if normalized:
                result.add(normalized)
        upper = re.sub(r"\s+", " ", text.upper()).strip()
        for name, code in COUNTRY_NAME_ALIASES.items():
            if name in upper:
                result.add(code)
        if context_is_country or any(token in parent_key for token in ("country", "страна", "nationality")):
            for token in re.split(r"[^A-Z]+", upper):
                if len(token) == 2:
                    normalized = normalize_country_code(token)
                    if normalized:
                        result.add(normalized)

    def visit_country_value(item: Any) -> None:
        if isinstance(item, (dict, list, tuple, set)):
            visit(item, "country")
            return
        text = str(item)
        upper = re.sub(r"\s+", " ", text.upper()).strip()
        direct = normalize_country_code(upper)
        if direct:
            result.add(direct)
        for token in re.split(r"[^A-Z]+", upper):
            if len(token) == 2:
                normalized = normalize_country_code(token)
                if normalized:
                    result.add(normalized)
        for name, code in COUNTRY_NAME_ALIASES.items():
            if name in upper:
                result.add(code)

    visit(parsed)
    return result


def generic_country_codes(value: Any) -> set[str]:
    result = country_codes_from_value(value, "country")
    text = flatten_text(value).upper()
    for token in re.findall(r"(?<![A-Z])([A-Z]{2})(?![A-Z])", text):
        normalized = normalize_country_code(token)
        if normalized:
            result.add(normalized)
    return result


def route_codes_from_value(value: Any) -> set[str]:
    text = flatten_text(value).upper()
    return {code for code in ROUTE_CODES if re.search(rf"(?<![A-Z]){re.escape(code)}(?![A-Z])", text)}


def extract_ipc(value: Any) -> set[str]:
    text = flatten_text(value).upper()
    full = {re.sub(r"\s+", "", match) for match in IPC_FULL_RE.findall(text)}
    if full:
        return full
    return {match.upper() for match in IPC_SUBCLASS_RE.findall(text)}


def extract_locarno(value: Any) -> set[str]:
    return {match.replace(".", "-") for match in LOCARNO_RE.findall(flatten_text(value))}


def serialize_codes(values: Iterable[str]) -> str:
    return "; ".join(sorted({str(value).strip().upper() for value in values if str(value).strip()}))


def normalize_registration(value: Any) -> str:
    if is_blank(value):
        return ""
    text = str(value).strip().upper()
    text = re.sub(r"\.0$", "", text)
    text = re.sub(r"[^A-Z0-9]", "", text)
    if text.startswith("RU"):
        text = text[2:]
    text = re.sub(r"(?:A1|A2|A3|C1|C2|C9|U1|U8|U9|S1|S2|S9|D1)$", "", text)
    return text


def combine_cell(old: Any, new: Any) -> str:
    values: list[str] = []
    for value in (old, new):
        if is_blank(value):
            continue
        text = str(value).strip()
        if text not in values:
            values.append(text)
    return " | ".join(values)


def source_columns_relevant(columns: Iterable[str]) -> list[str]:
    selected: list[str] = []
    tokens = (
        "registration", "регистрац", "author", "inventor", "creator", "автор",
        "holder", "assignee", "applicant", "owner", "right", "правооблад", "патентооблад",
        "country", "страна", "nationality", "ipc", "мпк", "locarno", "classification", "классиф",
    )
    for column in columns:
        normalized = norm_name(column)
        if any(token in normalized for token in tokens):
            selected.append(column)
    return selected


def classify_participant_columns(columns: Iterable[str]) -> tuple[list[str], list[str], list[str]]:
    author_cols: list[str] = []
    holder_cols: list[str] = []
    classification_cols: list[str] = []
    for column in columns:
        normalized = norm_name(column)
        if any(token in normalized for token in ("author", "inventor", "creator", "автор")):
            author_cols.append(column)
        if any(token in normalized for token in ("holder", "assignee", "applicant", "owner", "right_holder", "right_holders", "правооблад", "патентооблад")):
            holder_cols.append(column)
        if any(token in normalized for token in ("ipc", "мпк", "locarno", "classification", "классиф")):
            classification_cols.append(column)
    return author_cols, holder_cols, classification_cols


def read_source_lookup(spec: SourceSpec, needed_regs: set[str], chunk_size: int = 100_000) -> tuple[dict[str, dict[str, str]], dict[str, Any]]:
    path = resolve_macos_path(spec.path)
    if not path.exists():
        raise FileNotFoundError(f"Не найден исходный файл Роспатента: {path}")

    header = pd.read_csv(path, sep=spec.delimiter, dtype=str, nrows=0, encoding="utf-8-sig").columns.tolist()
    reg_col = next((column for column in header if norm_name(column) == "registration_number"), None)
    if reg_col is None:
        reg_col = next((column for column in header if "registration" in norm_name(column) and "number" in norm_name(column)), None)
    if reg_col is None:
        raise KeyError(f"В {path.name} не найден registration number")

    usecols = list(dict.fromkeys([reg_col] + source_columns_relevant(header)))
    lookup: dict[str, dict[str, str]] = {}
    rows_read = 0
    matched_rows = 0

    for chunk in pd.read_csv(
        path,
        sep=spec.delimiter,
        dtype=str,
        usecols=usecols,
        chunksize=chunk_size,
        encoding="utf-8-sig",
        low_memory=False,
        on_bad_lines="warn",
    ):
        rows_read += len(chunk)
        keys = chunk[reg_col].map(normalize_registration)
        mask = keys.isin(needed_regs)
        if not mask.any():
            continue
        selected = chunk.loc[mask].copy()
        selected["__reg_key"] = keys.loc[mask]
        matched_rows += len(selected)
        for _, row in selected.iterrows():
            key = row["__reg_key"]
            if not key:
                continue
            record = lookup.setdefault(key, {})
            for column in usecols:
                if column == reg_col:
                    continue
                record[column] = combine_cell(record.get(column), row.get(column))

    audit = {
        "object_type": spec.object_type,
        "source_file": str(path),
        "delimiter": spec.delimiter,
        "rows_read": rows_read,
        "needed_registration_numbers": len(needed_regs),
        "matched_registration_numbers": len(lookup),
        "matched_source_rows": matched_rows,
        "registration_column": reg_col,
        "selected_source_columns": "; ".join(usecols),
    }
    return lookup, audit


def choose_data_sheets(excel: pd.ExcelFile) -> dict[str, str]:
    result: dict[str, str] = {}
    for object_type, fragments in SHEET_ALIASES.items():
        for sheet in excel.sheet_names:
            normalized = norm_text(sheet).casefold()
            if object_type == "Программа для ЭВМ":
                if "программ" in normalized:
                    result[object_type] = sheet
                    break
            elif object_type == "База данных":
                if "баз" in normalized and "дан" in normalized:
                    result[object_type] = sheet
                    break
            elif any(fragment in normalized for fragment in fragments):
                result[object_type] = sheet
                break
    missing = set(SHEET_ALIASES) - set(result)
    if missing:
        raise RuntimeError(f"Не найдены листы типов объектов: {sorted(missing)}. Листы книги: {excel.sheet_names}")
    return result


def get_existing_columns(df: pd.DataFrame, keywords: Iterable[str], exclude: Iterable[str] = ()) -> list[str]:
    result: list[str] = []
    for column in df.columns:
        normalized = norm_name(column)
        if any(keyword in normalized for keyword in keywords) and not any(token in normalized for token in exclude):
            result.append(column)
    return result


def first_numeric(row: pd.Series, columns: Iterable[str]) -> int | None:
    for column in columns:
        value = pd.to_numeric(pd.Series([row.get(column)]), errors="coerce").iloc[0]
        if pd.notna(value):
            return int(value)
    return None


def list_length(value: Any) -> int:
    parsed = parse_nested(value)
    if parsed is None:
        return 0
    if isinstance(parsed, (list, tuple, set, dict)):
        return len(parsed)
    text = str(parsed).strip()
    if not text:
        return 0
    return len([part for part in re.split(r"\s*[|;,]\s*", text) if part])


def excel_safe(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    for column in result.select_dtypes(include=["object"]).columns:
        result[column] = result[column].map(lambda value: "" if is_blank(value) else str(value)[:32700])
    return result


def parquet_safe(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    for column in result.select_dtypes(include=["object"]).columns:
        result[column] = result[column].map(lambda value: None if is_blank(value) else str(value))
    return result


def process_sheet(object_type: str, df: pd.DataFrame, source_lookup: dict[str, dict[str, str]]) -> pd.DataFrame:
    result = df.copy()
    if "object_type" not in result.columns:
        result["object_type"] = object_type
    else:
        result["object_type"] = result["object_type"].fillna(object_type)

    reg_col = find_col(result, ["registration_number", "registration number"], required=True)
    result["registration_number_key_v4"] = result[reg_col].map(normalize_registration)

    final_author_cols, final_holder_cols, final_class_cols = classify_participant_columns(result.columns)
    google_ipc_cols = get_existing_columns(result, ("gp_ipc", "google_ipc", "ipc_codes_google"))
    google_cpc_cols = get_existing_columns(result, ("gp_cpc", "google_cpc", "cpc_codes_google"))
    family_country_cols = get_existing_columns(result, ("family_country", "foreign_country"))
    route_cols = get_existing_columns(result, ("route_code", "international_route"))
    forward_count_cols = get_existing_columns(result, ("forward_citation_count", "forward_citations_count", "citations_forward_count"))
    forward_list_cols = get_existing_columns(result, ("forward_citations",), exclude=("count",))

    author_countries_final: list[str] = []
    foreign_author_countries: list[str] = []
    holder_countries_final: list[str] = []
    foreign_holder_countries: list[str] = []
    author_sources: list[str] = []
    holder_sources: list[str] = []
    family_countries_final: list[str] = []
    foreign_family_countries: list[str] = []
    route_codes_final: list[str] = []
    international_flags: list[int] = []
    international_basis: list[str] = []
    all_countries_final: list[str] = []
    ipc_rospatent_values: list[str] = []
    ipc_google_values: list[str] = []
    ipc_final_values: list[str] = []
    ipc_source_values: list[str] = []
    locarno_values: list[str] = []
    citation_counts: list[float] = []
    cited_flags: list[float] = []

    for _, row in result.iterrows():
        reg_key = row["registration_number_key_v4"]
        source_record = source_lookup.get(reg_key, {})
        source_author_cols, source_holder_cols, source_class_cols = classify_participant_columns(source_record.keys())

        author_codes_source: set[str] = set()
        author_codes_final_fields: set[str] = set()
        holder_codes_source: set[str] = set()
        holder_codes_final_fields: set[str] = set()

        for column in source_author_cols:
            author_codes_source.update(country_codes_from_value(source_record.get(column), column))
        for column in final_author_cols:
            author_codes_final_fields.update(country_codes_from_value(row.get(column), column))
        for column in source_holder_cols:
            holder_codes_source.update(country_codes_from_value(source_record.get(column), column))
        for column in final_holder_cols:
            holder_codes_final_fields.update(country_codes_from_value(row.get(column), column))

        author_codes = author_codes_source | author_codes_final_fields
        holder_codes = {"RU"} | holder_codes_source | holder_codes_final_fields
        foreign_authors = author_codes - {"RU"}
        foreign_holders = holder_codes - {"RU"}

        if author_codes_source and author_codes_final_fields:
            author_source = "rospatent+final_fields"
        elif author_codes_source:
            author_source = "rospatent"
        elif author_codes_final_fields:
            author_source = "final_fields"
        else:
            author_source = "not_available"

        if holder_codes_source or holder_codes_final_fields:
            holder_source_parts = []
            if holder_codes_source:
                holder_source_parts.append("rospatent")
            if holder_codes_final_fields:
                holder_source_parts.append("final_fields")
            holder_source_parts.append("ru_university_holder")
            holder_source = "+".join(holder_source_parts)
        else:
            holder_source = "ru_university_holder"

        family_codes = {"RU"}
        routes: set[str] = set()
        if object_type in {"Изобретение", "Полезная модель"}:
            for column in family_country_cols:
                family_codes.update(generic_country_codes(row.get(column)))
            for column in route_cols:
                routes.update(route_codes_from_value(row.get(column)))
            routes.update({code for code in family_codes if code in ROUTE_CODES})
            family_codes -= ROUTE_CODES
        foreign_family = family_codes - {"RU"}
        international_flag = int(bool(foreign_family or routes)) if object_type in CLASSIC_TYPES else 0
        if object_type == "Промышленный образец":
            basis = "industrial_design_ru_only_accepted"
        elif object_type in DIGITAL_TYPES:
            basis = "not_applicable_digital_object"
        elif foreign_family and routes:
            basis = "foreign_country_and_route"
        elif foreign_family:
            basis = "foreign_country"
        elif routes:
            basis = "international_route"
        else:
            basis = "ru_only_no_foreign_evidence"

        ipc_source: set[str] = set()
        ipc_google: set[str] = set()
        locarno_source: set[str] = set()
        for column in source_class_cols:
            ipc_source.update(extract_ipc(source_record.get(column)))
            locarno_source.update(extract_locarno(source_record.get(column)))
        for column in final_class_cols:
            normalized = norm_name(column)
            if normalized.startswith("rospatent_raw") or "ipc_codes_rospatent" in normalized:
                ipc_source.update(extract_ipc(row.get(column)))
            if "locarno" in normalized:
                locarno_source.update(extract_locarno(row.get(column)))
        for column in google_ipc_cols:
            ipc_google.update(extract_ipc(row.get(column)))
        for column in google_cpc_cols:
            ipc_google.update(extract_ipc(row.get(column)))
        ipc_final = ipc_source | ipc_google
        if ipc_source and ipc_google:
            ipc_source_label = "rospatent+google"
        elif ipc_source:
            ipc_source_label = "rospatent"
        elif ipc_google:
            ipc_source_label = "google"
        else:
            ipc_source_label = "missing"

        if object_type == "Промышленный образец":
            citation_count = 0
            cited_flag = 0
        elif object_type in {"Изобретение", "Полезная модель"}:
            citation_count = first_numeric(row, forward_count_cols)
            if citation_count is None:
                citation_count = max([list_length(row.get(column)) for column in forward_list_cols] + [0])
            cited_flag = int(citation_count > 0)
        else:
            citation_count = np.nan
            cited_flag = np.nan

        all_codes = {"RU"} | author_codes | holder_codes | family_codes
        author_countries_final.append(serialize_codes(author_codes))
        foreign_author_countries.append(serialize_codes(foreign_authors))
        holder_countries_final.append(serialize_codes(holder_codes))
        foreign_holder_countries.append(serialize_codes(foreign_holders))
        author_sources.append(author_source)
        holder_sources.append(holder_source)
        family_countries_final.append(serialize_codes(family_codes))
        foreign_family_countries.append(serialize_codes(foreign_family))
        route_codes_final.append(serialize_codes(routes))
        international_flags.append(international_flag)
        international_basis.append(basis)
        all_countries_final.append(serialize_codes(all_codes))
        ipc_rospatent_values.append(serialize_codes(ipc_source))
        ipc_google_values.append(serialize_codes(ipc_google))
        ipc_final_values.append(serialize_codes(ipc_final))
        ipc_source_values.append(ipc_source_label)
        locarno_values.append(serialize_codes(locarno_source))
        citation_counts.append(citation_count)
        cited_flags.append(cited_flag)

    result["registration_country_final"] = "RU"
    result["registration_country_source_final"] = "Rospatent Russian national register"
    result["author_country_codes_final"] = author_countries_final
    result["foreign_author_country_codes_final"] = foreign_author_countries
    result["author_country_source_final"] = author_sources
    result["author_country_data_available_flag"] = result["author_country_codes_final"].astype(str).str.strip().ne("").astype(int)
    result["foreign_author_present_flag_final"] = result["foreign_author_country_codes_final"].astype(str).str.strip().ne("").astype(int)
    result["holder_country_codes_final"] = holder_countries_final
    result["foreign_holder_country_codes_final"] = foreign_holder_countries
    result["holder_country_source_final"] = holder_sources
    result["foreign_holder_present_flag_final"] = result["foreign_holder_country_codes_final"].astype(str).str.strip().ne("").astype(int)
    result["family_country_codes_final"] = family_countries_final
    result["foreign_family_country_codes_final"] = foreign_family_countries
    result["international_route_codes_final"] = route_codes_final
    result["international_patent_flag_final"] = international_flags
    result["international_basis_final"] = international_basis
    result["country_codes_all_final"] = all_countries_final
    result["ipc_codes_rospatent"] = ipc_rospatent_values
    result["ipc_codes_google"] = ipc_google_values
    result["ipc_codes_final"] = ipc_final_values
    result["ipc_source_final"] = ipc_source_values
    result["locarno_codes_rospatent_final"] = locarno_values
    result["forward_citation_count_final"] = citation_counts
    result["cited_patent_flag_final"] = cited_flags
    result["digital_international_collaboration_flag_final"] = np.where(
        result["object_type"].isin(DIGITAL_TYPES),
        result["foreign_holder_present_flag_final"],
        np.nan,
    )
    return result


def create_qc(processed: dict[str, pd.DataFrame], removed: pd.DataFrame, bad_keys: set[str], source_audit: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    combined = pd.concat([df.assign(source_object_type=object_type) for object_type, df in processed.items()], ignore_index=True, sort=False)
    key_col = find_col(combined, ["patent_key"], required=True)
    university_col = find_col(combined, ["university_canonical", "унифицированное название вуза"])
    qc: list[dict[str, Any]] = []

    def add(check: str, value: Any, expected: str, passed: bool, note: str = "") -> None:
        qc.append({"check": check, "value": value, "expected": expected, "passed": int(bool(passed)), "note": note})

    add("Удалённых parse_validation_failed patent_key", len(bad_keys), "1", len(bad_keys) == 1)
    add("Удалённых строк", len(removed), ">=1", len(removed) >= 1)
    add("Строк в итоговой базе", len(combined), ">0", len(combined) > 0)
    add("Уникальных объектов", combined[key_col].nunique(dropna=True), ">0", combined[key_col].nunique(dropna=True) > 0)
    if university_col:
        duplicates = int(combined.duplicated([key_col, university_col]).sum())
        add("Дубли patent_key + university", duplicates, "0", duplicates == 0)
    add("Пустые страны регистрации", int(combined["registration_country_final"].fillna("").astype(str).str.strip().eq("").sum()), "0", combined["registration_country_final"].fillna("").astype(str).str.strip().ne("").all())
    add("Правообладатели без RU", int(~combined["holder_country_codes_final"].fillna("").astype(str).str.contains(r"(^|;\s*)RU($|;)", regex=True).sum()), "0", combined["holder_country_codes_final"].fillna("").astype(str).str.contains(r"(^|;\s*)RU($|;)", regex=True).all())
    design = combined["object_type"].eq("Промышленный образец")
    add("Промышленные образцы с international != 0", int(pd.to_numeric(combined.loc[design, "international_patent_flag_final"], errors="coerce").fillna(-1).ne(0).sum()), "0", pd.to_numeric(combined.loc[design, "international_patent_flag_final"], errors="coerce").fillna(-1).eq(0).all())
    add("Промышленные образцы с cited != 0", int(pd.to_numeric(combined.loc[design, "cited_patent_flag_final"], errors="coerce").fillna(-1).ne(0).sum()), "0", pd.to_numeric(combined.loc[design, "cited_patent_flag_final"], errors="coerce").fillna(-1).eq(0).all())

    summary_rows: list[dict[str, Any]] = []
    for object_type, df in processed.items():
        key = find_col(df, ["patent_key"], required=True)
        university = find_col(df, ["university_canonical", "унифицированное название вуза"])
        summary_rows.append({
            "object_type": object_type,
            "rows": len(df),
            "unique_objects": df[key].nunique(dropna=True),
            "universities": df[university].nunique(dropna=True) if university else np.nan,
            "author_country_coverage": round(df["author_country_data_available_flag"].mean(), 4),
            "foreign_author_rows": int(df["foreign_author_present_flag_final"].sum()),
            "foreign_holder_rows": int(df["foreign_holder_present_flag_final"].sum()),
            "international_patent_rows": int(pd.to_numeric(df["international_patent_flag_final"], errors="coerce").fillna(0).sum()),
            "ipc_coverage": round(df["ipc_codes_final"].fillna("").astype(str).str.strip().ne("").mean(), 4),
        })

    country_summary = (
        combined.groupby("object_type", dropna=False)
        .agg(
            rows=(key_col, "size"),
            registration_country_coverage=("registration_country_final", lambda s: round(s.fillna("").astype(str).str.strip().ne("").mean(), 4)),
            author_country_coverage=("author_country_data_available_flag", "mean"),
            foreign_author_rows=("foreign_author_present_flag_final", "sum"),
            foreign_holder_rows=("foreign_holder_present_flag_final", "sum"),
            digital_international_collaboration_rows=("digital_international_collaboration_flag_final", lambda s: int(pd.to_numeric(s, errors="coerce").fillna(0).sum())),
        )
        .reset_index()
    )
    return pd.DataFrame(qc), pd.DataFrame(summary_rows), country_summary


def write_outputs(processed: dict[str, pd.DataFrame], removed: pd.DataFrame, qc: pd.DataFrame, summary: pd.DataFrame, country_summary: pd.DataFrame, source_audit: pd.DataFrame, input_path: Path, output_dir: Path) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    output_xlsx = output_dir / OUTPUT_FILENAME
    output_parquet = output_dir / OUTPUT_PARQUET_FILENAME
    removed_csv = output_dir / "4_removed_invalid_object.csv"

    readme = pd.DataFrame([
        {"parameter": "Версия", "value": VERSION},
        {"parameter": "Исходная итоговая книга", "value": str(input_path)},
        {"parameter": "Исходные реестры", "value": "Шесть оригинальных CSV Роспатента по указанным локальным путям"},
        {"parameter": "Удаление", "value": "Все строки patent_key со статусом parse_validation_failed удалены"},
        {"parameter": "Страна регистрации", "value": "RU для всех типов объектов, поскольку записи взяты из российских государственных реестров"},
        {"parameter": "Страны авторов", "value": "Из исходных полей Роспатента и доступных полей итоговой книги; без необоснованного присвоения RU при отсутствии данных"},
        {"parameter": "Страны правообладателей", "value": "RU добавляется как страна университета-правообладателя; иностранные коды добавляются из исходных данных"},
        {"parameter": "Международный патент", "value": "1 при подтверждённой зарубежной стране семейства или международном маршруте; иначе RU и 0"},
        {"parameter": "Промышленные образцы", "value": "Сохранены; RU; international=0; cited=0"},
        {"parameter": "МПК", "value": "Объединение оригинальных полей Роспатента и безопасно извлечённых кодов Google"},
        {"parameter": "Создано UTC", "value": datetime.now(timezone.utc).isoformat()},
    ])

    combined = pd.concat([df.assign(source_object_type=object_type) for object_type, df in processed.items()], ignore_index=True, sort=False)
    parquet_safe(combined).to_parquet(output_parquet, index=False)
    removed.to_csv(removed_csv, index=False, encoding="utf-8-sig")

    sheet_names = {
        "Изобретение": "Изобретения",
        "Полезная модель": "Полезные модели",
        "Промышленный образец": "Промышленные образцы",
        "Программа для ЭВМ": "Программы ЭВМ",
        "База данных": "Базы данных",
        "ТИМС": "ТИМС",
    }

    with pd.ExcelWriter(
        output_xlsx,
        engine="xlsxwriter",
        engine_kwargs={"options": {"constant_memory": True, "strings_to_urls": False, "nan_inf_to_errors": True}},
    ) as writer:
        service_sheets = {
            "4_README": readme,
            "4_QC_финал": qc,
            "4_Сводка_типы": summary,
            "4_Страны_QC": country_summary,
            "4_Источники": source_audit,
            "4_Удаленные": removed,
        }
        for sheet, frame in service_sheets.items():
            excel_safe(frame).to_excel(writer, sheet_name=sheet[:31], index=False)
        for object_type, frame in processed.items():
            excel_safe(frame).to_excel(writer, sheet_name=sheet_names[object_type][:31], index=False)

        workbook = writer.book
        header_format = workbook.add_format({"bold": True, "text_wrap": True, "valign": "top", "bg_color": "#D9EAF7", "border": 1})
        for sheet_name, worksheet in writer.sheets.items():
            worksheet.freeze_panes(1, 0)
            worksheet.set_row(0, 32, header_format)
            worksheet.autofilter(0, 0, worksheet.dim_rowmax, worksheet.dim_colmax)
            worksheet.set_default_row(18)
            worksheet.set_column(0, max(worksheet.dim_colmax, 0), 16)

    return {"excel": output_xlsx, "parquet": output_parquet, "removed_csv": removed_csv}


def main() -> dict[str, Path]:
    input_path = find_final_workbook()
    output_dir = resolve_macos_path(PROJECT_ROOT / "Выгрузка данных") / OUTPUT_DIRNAME
    print(f"Исходная итоговая книга: {input_path}")
    print(f"Папка результатов: {output_dir}")

    excel = pd.ExcelFile(input_path, engine="openpyxl")
    sheet_map = choose_data_sheets(excel)
    print("Листы данных:")
    for object_type, sheet in sheet_map.items():
        print(f"  {object_type}: {sheet}")

    data: dict[str, pd.DataFrame] = {}
    needed_regs: dict[str, set[str]] = {}
    for object_type, sheet in sheet_map.items():
        frame = pd.read_excel(excel, sheet_name=sheet, dtype=object)
        reg_col = find_col(frame, ["registration_number", "registration number"], required=True)
        data[object_type] = frame
        needed_regs[object_type] = {normalize_registration(value) for value in frame[reg_col] if normalize_registration(value)}
        print(f"Прочитан {object_type}: {len(frame):,} строк; {len(needed_regs[object_type]):,} регистрационных номеров")

    bad_keys: set[str] = set()
    for frame in data.values():
        status_col = find_col(frame, ["gp_status", "google_status"])
        key_col = find_col(frame, ["patent_key"])
        if status_col and key_col:
            mask = frame[status_col].fillna("").astype(str).str.strip().str.casefold().eq("parse_validation_failed")
            bad_keys.update(frame.loc[mask, key_col].dropna().astype(str))

    for sheet in excel.sheet_names:
        if "error" not in norm_name(sheet) and "ошиб" not in norm_name(sheet):
            continue
        try:
            error_frame = pd.read_excel(excel, sheet_name=sheet, dtype=object)
        except Exception:
            continue
        status_col = find_col(error_frame, ["gp_status", "status"])
        key_col = find_col(error_frame, ["patent_key"])
        if status_col and key_col:
            mask = error_frame[status_col].fillna("").astype(str).str.strip().str.casefold().eq("parse_validation_failed")
            bad_keys.update(error_frame.loc[mask, key_col].dropna().astype(str))

    removed_parts: list[pd.DataFrame] = []
    if bad_keys:
        for object_type, frame in list(data.items()):
            key_col = find_col(frame, ["patent_key"], required=True)
            mask = frame[key_col].fillna("").astype(str).isin(bad_keys)
            if mask.any():
                removed = frame.loc[mask].copy()
                removed.insert(0, "object_type_removed", object_type)
                removed_parts.append(removed)
                data[object_type] = frame.loc[~mask].reset_index(drop=True)
    removed = pd.concat(removed_parts, ignore_index=True, sort=False) if removed_parts else pd.DataFrame()
    print(f"Удалено невалидных patent_key: {len(bad_keys)}; строк: {len(removed)}")

    source_lookup: dict[str, dict[str, dict[str, str]]] = {}
    audit_rows: list[dict[str, Any]] = []
    spec_map = {spec.object_type: spec for spec in SOURCE_SPECS}
    for object_type in SHEET_ALIASES:
        lookup, audit = read_source_lookup(spec_map[object_type], needed_regs[object_type])
        source_lookup[object_type] = lookup
        audit_rows.append(audit)
        print(f"Роспатент {object_type}: найдено {len(lookup):,} из {len(needed_regs[object_type]):,} номеров")

    processed: dict[str, pd.DataFrame] = {}
    for object_type, frame in data.items():
        print(f"Обработка {object_type}...")
        processed[object_type] = process_sheet(object_type, frame, source_lookup[object_type])

    source_audit = pd.DataFrame(audit_rows)
    qc, summary, country_summary = create_qc(processed, removed, bad_keys, source_audit)
    outputs = write_outputs(processed, removed, qc, summary, country_summary, source_audit, input_path, output_dir)

    print("\nОбработка завершена")
    print(qc.to_string(index=False))
    print("\nСводка:")
    print(summary.to_string(index=False))
    print("\nФайлы:")
    for name, path in outputs.items():
        print(f"  {name}: {path}")
    return outputs


if __name__ == "__main__":
    main()
