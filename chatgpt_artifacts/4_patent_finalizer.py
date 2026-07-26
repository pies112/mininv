from __future__ import annotations

import ast
import csv
import json
import re
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

import numpy as np
import pandas as pd

VERSION = "4"

PROJECT_ROOT = Path(
    "/Users/vsevolodkarass/Library/Mobile Documents/com~apple~CloudDocs/"
    "Desktop/Рабочий стол — MacBook Air — Всеволод/"
    "Инженерный проект/Версия рейтинга 1"
)

INPUT_FILENAME = "3_patents_2020_2025_final.xlsx"
OUTPUT_DIR = PROJECT_ROOT / "Выгрузка данных" / "4_Финальная патентная база 2020-2025"
OUTPUT_XLSX = OUTPUT_DIR / "4_patents_2020_2025_final_with_countries.xlsx"
OUTPUT_PARQUET = OUTPUT_DIR / "4_patents_2020_2025_final_with_countries.parquet"
OUTPUT_REMOVED = OUTPUT_DIR / "4_removed_parse_validation_failed.csv"
OUTPUT_SOURCE_MATCH = OUTPUT_DIR / "4_rospatent_source_match_qc.csv"

ROSPATENT_DIR = PROJECT_ROOT / "Выгрузка данных" / "Роспатент"
ROSPATENT_FILES = {
    "Изобретение": "Открытый реестр изобретений Российской Федерации.csv",
    "Полезная модель": "Открытый реестр полезных моделей Российской Федерации.csv",
    "Промышленный образец": "Открытый реестр промышленных образцов Российской Федерации.csv",
    "Программа для ЭВМ": "Открытый реестр программ для электронно-вычислительных машин.csv",
    "База данных": "Открытый реестр баз данных.csv",
    "ТИМС": "Открытый реестр топологий интегральных микросхем.csv",
}

CLASSIC_TYPES = {"Изобретение", "Полезная модель", "Промышленный образец"}
DIGITAL_TYPES = {"Программа для ЭВМ", "База данных", "ТИМС"}
ROUTE_CODES = {"WO", "EP", "EA", "AP", "OA", "GC", "EM", "IB", "PCT"}
ISO2_RE = re.compile(r"(?<![A-Z])[A-Z]{2}(?![A-Z])")
IPC_RE = re.compile(r"\b([A-HY]\d{2}[A-Z](?:\s*\d{1,4}(?:/\d{1,8})?)?)\b", re.I)
LOCARNO_RE = re.compile(r"\b(\d{1,2}[-.]\d{1,2})\b")

COUNTRY_NAME_MAP = {
    "россия": "RU", "российская федерация": "RU", "russia": "RU", "russian federation": "RU",
    "сша": "US", "соединенные штаты": "US", "соединённые штаты": "US", "united states": "US", "usa": "US",
    "китай": "CN", "china": "CN", "германия": "DE", "germany": "DE", "франция": "FR", "france": "FR",
    "великобритания": "GB", "соединенное королевство": "GB", "соединённое королевство": "GB", "united kingdom": "GB",
    "япония": "JP", "japan": "JP", "корея": "KR", "южная корея": "KR", "republic of korea": "KR",
    "индия": "IN", "india": "IN", "италия": "IT", "italy": "IT", "испания": "ES", "spain": "ES",
    "швейцария": "CH", "switzerland": "CH", "нидерланды": "NL", "netherlands": "NL",
    "беларусь": "BY", "белоруссия": "BY", "belarus": "BY", "казахстан": "KZ", "kazakhstan": "KZ",
    "украина": "UA", "ukraine": "UA", "ирландия": "IE", "ireland": "IE", "австрия": "AT", "austria": "AT",
    "канада": "CA", "canada": "CA", "австралия": "AU", "australia": "AU", "швеция": "SE", "sweden": "SE",
    "финляндия": "FI", "finland": "FI", "норвегия": "NO", "norway": "NO", "польша": "PL", "poland": "PL",
    "чехия": "CZ", "czech republic": "CZ", "czechia": "CZ", "израиль": "IL", "israel": "IL",
    "сингапур": "SG", "singapore": "SG", "тайвань": "TW", "taiwan": "TW", "бразилия": "BR", "brazil": "BR",
}


def norm_text(value: Any) -> str:
    return unicodedata.normalize("NFC", str(value)).replace("\xa0", " ").strip()


def norm_col(value: Any) -> str:
    return re.sub(r"[^a-zа-яё0-9]+", "_", norm_text(value).casefold()).strip("_")


def norm_reg(value: Any) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    text = norm_text(value).upper()
    text = re.sub(r"\.0$", "", text)
    text = re.sub(r"[^A-ZА-Я0-9]", "", text)
    return text


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    return norm_text(value).casefold() in {"", "nan", "none", "null", "[]", "{}"}


def parse_list_like(value: Any) -> list[Any]:
    if is_blank(value):
        return []
    if isinstance(value, list):
        return value
    if isinstance(value, (tuple, set)):
        return list(value)
    if isinstance(value, dict):
        return list(value.values())
    text = norm_text(value)
    for parser in (json.loads, ast.literal_eval):
        try:
            parsed = parser(text)
            if isinstance(parsed, list):
                return parsed
            if isinstance(parsed, (tuple, set)):
                return list(parsed)
            if isinstance(parsed, dict):
                return list(parsed.values())
            return [parsed]
        except Exception:
            pass
    return [part.strip() for part in re.split(r"[|;]\s*", text) if part.strip()]


def flatten(value: Any) -> str:
    if is_blank(value):
        return ""
    if isinstance(value, dict):
        return " ".join(flatten(v) for v in value.values())
    if isinstance(value, (list, tuple, set)):
        return " ".join(flatten(v) for v in value)
    return norm_text(value)


def serialize_codes(values: Iterable[str]) -> str:
    return "; ".join(sorted({str(v).strip().upper() for v in values if str(v).strip()}))


def extract_country_codes(value: Any) -> set[str]:
    text = flatten(value)
    upper = text.upper()
    codes: set[str] = set()
    patterns = [
        r"(?i)[\"']?country(?:_code)?[\"']?\s*[:=]\s*[\"']?([A-Z]{2})[\"']?",
        r"\(([A-Z]{2})\)", r"\[([A-Z]{2})\]", r"\bCOUNTRY\s+([A-Z]{2})\b",
    ]
    for pattern in patterns:
        codes.update(m.upper() for m in re.findall(pattern, text))
    for name, code in COUNTRY_NAME_MAP.items():
        if name in text.casefold():
            codes.add(code)
    return {c for c in codes if c not in ROUTE_CODES}


def extract_structured_codes(value: Any) -> set[str]:
    codes: set[str] = set()
    for item in parse_list_like(value):
        text = flatten(item).upper()
        codes.update(ISO2_RE.findall(text))
        if "PCT" in text:
            codes.add("PCT")
    return codes


def extract_ipc(value: Any) -> set[str]:
    return {re.sub(r"\s+", "", m.upper()) for m in IPC_RE.findall(flatten(value).upper())}


def extract_locarno(value: Any) -> set[str]:
    return {m.replace(".", "-") for m in LOCARNO_RE.findall(flatten(value))}


def find_column(columns: Iterable[str], candidates: Iterable[str], contains: bool = False) -> str | None:
    mapping = {norm_col(c): c for c in columns}
    for candidate in candidates:
        key = norm_col(candidate)
        if key in mapping:
            return mapping[key]
    if contains:
        for column in columns:
            n = norm_col(column)
            if any(norm_col(c) in n for c in candidates):
                return column
    return None


def resolve_file(root: Path, filename: str) -> Path:
    direct = root / filename
    if direct.exists():
        return direct
    target = norm_text(filename).casefold()
    if root.exists():
        for p in root.rglob("*"):
            if p.is_file() and norm_text(p.name).casefold() == target:
                return p
    raise FileNotFoundError(f"Не найден файл: {filename}\nИскали внутри: {root}")


def find_input_xlsx() -> Path:
    candidates = [
        PROJECT_ROOT / "Выгрузка данных" / "3_Финальная патентная база 2020-2025" / INPUT_FILENAME,
        PROJECT_ROOT / "Выгрузка данных" / INPUT_FILENAME,
        PROJECT_ROOT / INPUT_FILENAME,
        Path.cwd() / INPUT_FILENAME,
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    for root in [PROJECT_ROOT, Path.cwd()]:
        if root.exists():
            for p in root.rglob("*.xlsx"):
                if norm_text(p.name).casefold() == norm_text(INPUT_FILENAME).casefold():
                    return p
    raise FileNotFoundError(f"Не найден {INPUT_FILENAME}")


def detect_delimiter(path: Path) -> str:
    with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as fh:
        sample = fh.read(10000)
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except Exception:
        return ";" if sample.count(";") >= sample.count(",") else ","


def infer_object_type(sheet_name: str, df: pd.DataFrame) -> pd.Series:
    col = find_column(df.columns, ["object_type", "тип объекта"])
    if col:
        return df[col].fillna("").astype(str)
    name = sheet_name.casefold()
    mapping = {
        "изобрет": "Изобретение", "полез": "Полезная модель", "промышлен": "Промышленный образец",
        "программ": "Программа для ЭВМ", "баз": "База данных", "тимс": "ТИМС", "тополог": "ТИМС",
    }
    value = next((v for k, v in mapping.items() if k in name), "")
    return pd.Series(value, index=df.index, dtype="object")


def participant_columns(columns: Iterable[str], role: str) -> list[str]:
    result = []
    for col in columns:
        n = norm_col(col)
        if role == "author" and any(k in n for k in ["author", "inventor", "автор"]):
            result.append(col)
        if role == "holder" and any(k in n for k in ["holder", "assignee", "правооблад"]):
            result.append(col)
    return result


def classification_columns(columns: Iterable[str], kind: str) -> list[str]:
    result = []
    for col in columns:
        n = norm_col(col)
        if kind == "ipc" and any(k in n for k in ["ipc", "мпк", "patent_classification"]):
            result.append(col)
        if kind == "locarno" and "locarno" in n:
            result.append(col)
    return result


def scan_rospatent_sources(needed: dict[str, set[str]]) -> tuple[dict[tuple[str, str], dict[str, Any]], pd.DataFrame]:
    records: dict[tuple[str, str], dict[str, Any]] = {}
    qc = []
    for object_type, filename in ROSPATENT_FILES.items():
        path = resolve_file(ROSPATENT_DIR, filename)
        delimiter = detect_delimiter(path)
        header = pd.read_csv(path, sep=delimiter, encoding="utf-8-sig", nrows=0).columns.tolist()
        reg_col = find_column(header, ["registration number", "registration_number", "номер регистрации"], contains=True)
        if not reg_col:
            raise KeyError(f"В {path.name} не найден столбец registration number")
        author_cols = participant_columns(header, "author")
        holder_cols = participant_columns(header, "holder")
        ipc_cols = classification_columns(header, "ipc")
        locarno_cols = classification_columns(header, "locarno")
        wanted = needed.get(object_type, set())
        read_rows = matched_rows = 0
        for chunk in pd.read_csv(
            path, sep=delimiter, encoding="utf-8-sig", dtype=str, chunksize=100000,
            low_memory=False, on_bad_lines="skip"
        ):
            read_rows += len(chunk)
            keys = chunk[reg_col].map(norm_reg)
            matched = chunk.loc[keys.isin(wanted)].copy()
            if matched.empty:
                continue
            matched["__reg_key"] = keys.loc[matched.index]
            for _, row in matched.iterrows():
                key = (object_type, row["__reg_key"])
                author_codes: set[str] = set()
                holder_codes: set[str] = {"RU"}
                ipc_codes: set[str] = set()
                locarno_codes: set[str] = set()
                for c in author_cols:
                    author_codes |= extract_country_codes(row.get(c))
                for c in holder_cols:
                    holder_codes |= extract_country_codes(row.get(c))
                for c in ipc_cols:
                    ipc_codes |= extract_ipc(row.get(c))
                for c in locarno_cols:
                    locarno_codes |= extract_locarno(row.get(c))
                if not author_codes:
                    author_codes = {"RU"}
                records[key] = {
                    "rospatent_author_country_codes": serialize_codes(author_codes),
                    "rospatent_holder_country_codes": serialize_codes(holder_codes),
                    "rospatent_ipc_codes": serialize_codes(ipc_codes),
                    "rospatent_locarno_codes": serialize_codes(locarno_codes),
                }
                matched_rows += 1
        qc.append({
            "object_type": object_type, "source_file": path.name, "delimiter": delimiter,
            "rows_read": read_rows, "needed_registration_numbers": len(wanted),
            "matched_source_rows": matched_rows, "unique_matches": sum(1 for k in records if k[0] == object_type),
            "author_columns": "; ".join(author_cols), "holder_columns": "; ".join(holder_cols),
        })
    return records, pd.DataFrame(qc)


def list_count(value: Any) -> int:
    return len(parse_list_like(value))


def first_numeric(row: pd.Series, candidates: list[str]) -> int | None:
    for col in candidates:
        if col in row.index:
            value = pd.to_numeric(pd.Series([row[col]]), errors="coerce").iloc[0]
            if pd.notna(value):
                return int(value)
    return None


def excel_safe(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            try:
                out[col] = out[col].dt.tz_localize(None)
            except Exception:
                pass
        elif out[col].dtype == "object":
            out[col] = out[col].map(lambda v: "" if is_blank(v) else str(v)[:32700])
    return out


def main() -> dict[str, Path]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    input_xlsx = find_input_xlsx()
    print("Исходный итоговый файл:", input_xlsx)
    print("Папка результатов:", OUTPUT_DIR)

    sheets = pd.read_excel(input_xlsx, sheet_name=None, dtype=object)
    data_sheets: dict[str, pd.DataFrame] = {}
    other_sheets: dict[str, pd.DataFrame] = {}
    for name, df in sheets.items():
        if find_column(df.columns, ["patent_key"]):
            data_sheets[name] = df.copy()
        else:
            other_sheets[name] = df.copy()
    if not data_sheets:
        raise RuntimeError("В итоговой книге не найдено листов со столбцом patent_key")

    bad_keys: set[str] = set()
    for df in data_sheets.values():
        key_col = find_column(df.columns, ["patent_key"])
        status_col = find_column(df.columns, ["gp_status", "google_status"])
        if status_col:
            mask = df[status_col].fillna("").astype(str).str.casefold().eq("parse_validation_failed")
            bad_keys |= set(df.loc[mask, key_col].dropna().astype(str))

    removed_parts = []
    for name, df in list(data_sheets.items()):
        key_col = find_column(df.columns, ["patent_key"])
        mask = df[key_col].fillna("").astype(str).isin(bad_keys)
        if mask.any():
            part = df.loc[mask].copy()
            part.insert(0, "source_sheet", name)
            removed_parts.append(part)
        data_sheets[name] = df.loc[~mask].reset_index(drop=True)
    removed = pd.concat(removed_parts, ignore_index=True, sort=False) if removed_parts else pd.DataFrame()
    removed.to_csv(OUTPUT_REMOVED, index=False, encoding="utf-8-sig")
    print(f"Удалено невалидных patent_key: {len(bad_keys)}; строк: {len(removed)}")

    needed: dict[str, set[str]] = {t: set() for t in ROSPATENT_FILES}
    for name, df in data_sheets.items():
        df["object_type"] = infer_object_type(name, df)
        reg_col = find_column(df.columns, ["registration_number", "registration number", "номер регистрации"], contains=True)
        if not reg_col:
            raise KeyError(f"На листе {name} не найден registration_number")
        df["__registration_key"] = df[reg_col].map(norm_reg)
        for object_type, group in df.groupby("object_type"):
            if object_type in needed:
                needed[object_type] |= set(group["__registration_key"].dropna()) - {""}
        data_sheets[name] = df

    print("Сканирование оригинальных реестров Роспатента...")
    source_records, source_qc = scan_rospatent_sources(needed)
    source_qc.to_csv(OUTPUT_SOURCE_MATCH, index=False, encoding="utf-8-sig")
    print(source_qc.to_string(index=False))

    processed: dict[str, pd.DataFrame] = {}
    for name, df in data_sheets.items():
        df = df.copy()
        source_author = []
        source_holder = []
        source_ipc = []
        source_locarno = []
        for _, row in df.iterrows():
            rec = source_records.get((str(row["object_type"]), str(row["__registration_key"])), {})
            source_author.append(rec.get("rospatent_author_country_codes", ""))
            source_holder.append(rec.get("rospatent_holder_country_codes", "RU"))
            source_ipc.append(rec.get("rospatent_ipc_codes", ""))
            source_locarno.append(rec.get("rospatent_locarno_codes", ""))
        df["rospatent_author_country_codes"] = source_author
        df["rospatent_holder_country_codes"] = source_holder
        df["rospatent_ipc_codes"] = source_ipc
        df["rospatent_locarno_codes"] = source_locarno

        author_cols = participant_columns(df.columns, "author")
        holder_cols = participant_columns(df.columns, "holder")
        family_cols = [c for c in df.columns if any(k in norm_col(c) for k in ["family_country", "foreign_country"])]
        route_cols = [c for c in df.columns if "route" in norm_col(c)]
        ipc_cols = classification_columns(df.columns, "ipc")
        locarno_cols = classification_columns(df.columns, "locarno")

        registration_countries = []
        author_countries = []
        foreign_author_countries = []
        holder_countries = []
        foreign_holder_countries = []
        all_participant_countries = []
        family_countries = []
        foreign_family_countries = []
        route_codes_final = []
        international_flags = []
        international_basis = []
        citation_counts = []
        cited_flags = []
        ipc_final = []
        ipc_source = []
        locarno_final = []

        forward_count_cols = [c for c in df.columns if any(k in norm_col(c) for k in ["forward_citation_count", "citations_forward_count"])]
        forward_list_cols = [c for c in df.columns if "forward_citation" in norm_col(c) and "count" not in norm_col(c)]

        for _, row in df.iterrows():
            object_type = str(row["object_type"])
            registration_countries.append("RU")

            authors = extract_country_codes(row.get("rospatent_author_country_codes"))
            holders = {"RU"} | extract_country_codes(row.get("rospatent_holder_country_codes"))
            for c in author_cols:
                authors |= extract_country_codes(row.get(c))
            for c in holder_cols:
                holders |= extract_country_codes(row.get(c))
            if not authors:
                authors = {"RU"}
            author_countries.append(serialize_codes(authors))
            foreign_author_countries.append(serialize_codes(authors - {"RU"}))
            holder_countries.append(serialize_codes(holders))
            foreign_holder_countries.append(serialize_codes(holders - {"RU"}))
            all_participant_countries.append(serialize_codes(authors | holders))

            foreign: set[str] = set()
            routes: set[str] = set()
            for c in family_cols:
                values = extract_structured_codes(row.get(c))
                routes |= values & ROUTE_CODES
                foreign |= values - ROUTE_CODES - {"RU"}
            for c in route_cols:
                values = extract_structured_codes(row.get(c))
                routes |= values & ROUTE_CODES
            if object_type == "Промышленный образец" or object_type in DIGITAL_TYPES:
                foreign = set()
                routes = set()
            family = {"RU"} | foreign
            flag = int(bool(foreign or routes)) if object_type in {"Изобретение", "Полезная модель"} else 0
            family_countries.append(serialize_codes(family))
            foreign_family_countries.append(serialize_codes(foreign))
            route_codes_final.append(serialize_codes(routes))
            international_flags.append(flag)
            if flag and foreign and routes:
                international_basis.append("foreign_country_and_route")
            elif flag and foreign:
                international_basis.append("foreign_country")
            elif flag and routes:
                international_basis.append("international_route")
            elif object_type == "Промышленный образец":
                international_basis.append("industrial_design_ru_only_accepted")
            elif object_type in DIGITAL_TYPES:
                international_basis.append("not_applicable_digital_object")
            else:
                international_basis.append("ru_only_no_foreign_evidence")

            if object_type == "Промышленный образец":
                count = 0
            elif object_type in {"Изобретение", "Полезная модель"}:
                count = first_numeric(row, forward_count_cols)
                if count is None:
                    count = max([list_count(row.get(c)) for c in forward_list_cols] + [0])
            else:
                count = np.nan
            citation_counts.append(count)
            cited_flags.append(int(count > 0) if pd.notna(count) else np.nan)

            ipc_codes = extract_ipc(row.get("rospatent_ipc_codes"))
            for c in ipc_cols:
                ipc_codes |= extract_ipc(row.get(c))
            ipc_final.append(serialize_codes(ipc_codes))
            if extract_ipc(row.get("rospatent_ipc_codes")) and len(ipc_codes) > len(extract_ipc(row.get("rospatent_ipc_codes"))):
                ipc_source.append("rospatent+google_or_existing")
            elif extract_ipc(row.get("rospatent_ipc_codes")):
                ipc_source.append("rospatent")
            elif ipc_codes:
                ipc_source.append("google_or_existing")
            else:
                ipc_source.append("not_applicable" if object_type not in {"Изобретение", "Полезная модель"} else "missing")

            loc_codes = extract_locarno(row.get("rospatent_locarno_codes"))
            for c in locarno_cols:
                loc_codes |= extract_locarno(row.get(c))
            locarno_final.append(serialize_codes(loc_codes))

        df["registration_country_final"] = registration_countries
        df["author_country_codes_final"] = author_countries
        df["foreign_author_country_codes_final"] = foreign_author_countries
        df["holder_country_codes_final"] = holder_countries
        df["foreign_holder_country_codes_final"] = foreign_holder_countries
        df["participant_country_codes_all_final"] = all_participant_countries
        df["family_country_codes_final"] = family_countries
        df["foreign_family_country_codes_final"] = foreign_family_countries
        df["international_route_codes_final"] = route_codes_final
        df["international_patent_flag_final"] = international_flags
        df["international_basis_final"] = international_basis
        df["forward_citation_count_final"] = citation_counts
        df["cited_patent_flag_final"] = cited_flags
        df["ipc_codes_final"] = ipc_final
        df["ipc_source_final"] = ipc_source
        df["locarno_codes_final"] = locarno_final
        df["country_data_source_final"] = "rospatent_registry+google_patents_where_available"
        df.drop(columns=["__registration_key"], inplace=True, errors="ignore")
        processed[name] = df

    combined_parts = []
    for name, df in processed.items():
        part = df.copy()
        part.insert(0, "source_sheet", name)
        combined_parts.append(part)
    combined = pd.concat(combined_parts, ignore_index=True, sort=False)
    key_col = find_column(combined.columns, ["patent_key"])
    uni_col = find_column(combined.columns, ["university_canonical", "унифицированное название вуза"])

    qc_rows = []
    def add_qc(check: str, value: Any, expected: str, passed: bool, note: str = ""):
        qc_rows.append({"check": check, "value": value, "expected": expected, "passed": int(passed), "note": note})
    add_qc("Удалено parse_validation_failed patent_key", len(bad_keys), "1", len(bad_keys) == 1)
    add_qc("Строк в итоговой базе", len(combined), ">0", len(combined) > 0)
    add_qc("Уникальных объектов", combined[key_col].nunique(), ">0", combined[key_col].nunique() > 0)
    if uni_col:
        dups = int(combined.duplicated([key_col, uni_col]).sum())
        add_qc("Дубли patent_key + university", dups, "0", dups == 0)
    add_qc("Пустая страна регистрации", int(combined["registration_country_final"].eq("").sum()), "0", not combined["registration_country_final"].eq("").any())
    add_qc("Пустые страны авторов", int(combined["author_country_codes_final"].eq("").sum()), "0", not combined["author_country_codes_final"].eq("").any(), "При отсутствии явного кода используется консервативное RU")
    add_qc("Пустые страны правообладателей", int(combined["holder_country_codes_final"].eq("").sum()), "0", not combined["holder_country_codes_final"].eq("").any())
    add_qc("Пустые страны патентного семейства", int(combined["family_country_codes_final"].eq("").sum()), "0", not combined["family_country_codes_final"].eq("").any())
    add_qc("Промышленные образцы с international != 0", int((combined.loc[combined["object_type"].eq("Промышленный образец"), "international_patent_flag_final"] != 0).sum()), "0", bool((combined.loc[combined["object_type"].eq("Промышленный образец"), "international_patent_flag_final"] == 0).all()))
    qc = pd.DataFrame(qc_rows)

    summary = combined.groupby("object_type", dropna=False).agg(
        rows=(key_col, "size"), unique_objects=(key_col, "nunique"),
        international_objects=("international_patent_flag_final", "sum"),
        cited_objects=("cited_patent_flag_final", lambda s: int(pd.to_numeric(s, errors="coerce").fillna(0).sum())),
        foreign_author_rows=("foreign_author_country_codes_final", lambda s: int(s.fillna("").astype(str).str.strip().ne("").sum())),
        foreign_holder_rows=("foreign_holder_country_codes_final", lambda s: int(s.fillna("").astype(str).str.strip().ne("").sum())),
    ).reset_index()

    readme = pd.DataFrame([
        {"parameter": "Версия", "value": VERSION},
        {"parameter": "Исходный файл", "value": str(input_xlsx)},
        {"parameter": "Удаление ошибки", "value": "Полностью удалён patent_key со статусом parse_validation_failed"},
        {"parameter": "Страна регистрации", "value": "RU для всех объектов национальных реестров Роспатента"},
        {"parameter": "Международность", "value": "1 только при подтверждённой зарубежной стране семейства или международном маршруте; иначе RU и 0"},
        {"parameter": "Промышленные образцы", "value": "Сохранены; RU; international=0; cited=0"},
        {"parameter": "Страны участников", "value": "Для всех шести типов добавлены страны авторов и правообладателей по оригинальным реестрам и доступным полям итоговой базы"},
        {"parameter": "Создано UTC", "value": datetime.now(timezone.utc).isoformat()},
    ])

    parquet_df = combined.copy()
    for col in parquet_df.select_dtypes(include=["object"]).columns:
        parquet_df[col] = parquet_df[col].map(lambda v: None if is_blank(v) else str(v))
    parquet_df.to_parquet(OUTPUT_PARQUET, index=False)

    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False, "constant_memory": True}}) as writer:
        excel_safe(readme).to_excel(writer, sheet_name="4_README", index=False)
        excel_safe(qc).to_excel(writer, sheet_name="4_QC_финал", index=False)
        excel_safe(summary).to_excel(writer, sheet_name="4_Сводка_типы", index=False)
        excel_safe(source_qc).to_excel(writer, sheet_name="4_QC_Роспатент", index=False)
        excel_safe(removed).to_excel(writer, sheet_name="4_Удаленные", index=False)
        for name, df in processed.items():
            excel_safe(df).to_excel(writer, sheet_name=name[:31], index=False)
        existing = {"4_README", "4_QC_финал", "4_Сводка_типы", "4_QC_Роспатент", "4_Удаленные"} | {n[:31] for n in processed}
        for name, df in other_sheets.items():
            safe = name[:31]
            if safe in existing:
                safe = ("old_" + safe)[:31]
            excel_safe(df).to_excel(writer, sheet_name=safe, index=False)
        for worksheet in writer.sheets.values():
            worksheet.freeze_panes(1, 0)

    print("\nГотово.")
    print("Excel:", OUTPUT_XLSX)
    print("Parquet:", OUTPUT_PARQUET)
    print("Удалённые строки:", OUTPUT_REMOVED)
    print("\nQC:")
    print(qc.to_string(index=False))
    print("\nСводка:")
    print(summary.to_string(index=False))
    return {"excel": OUTPUT_XLSX, "parquet": OUTPUT_PARQUET, "removed": OUTPUT_REMOVED, "source_qc": OUTPUT_SOURCE_MATCH}


if __name__ == "__main__":
    outputs = main()
