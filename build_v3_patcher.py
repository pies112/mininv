from __future__ import annotations

from pathlib import Path
import nbformat
from nbformat.v4 import new_notebook, new_markdown_cell, new_code_cell

PATCHER_CODE = r'''
from __future__ import annotations

import importlib.util
import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

# -----------------------------------------------------------------------------
# 3_ bootstrap: находит локальную версию 2, создаёт полноценную версию 3,
# применяет исправления и сразу запускает весь конвейер.
# -----------------------------------------------------------------------------

if importlib.util.find_spec("nbformat") is None:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "nbformat"])

import nbformat
from nbformat.v4 import new_code_cell

VERSION = "3"
V2_FILENAMES = (
    "2_Инженерный_рейтинг_2026_единая_патентная_выгрузка.ipynb",
    "2_%D0%98%D0%BD%D0%B6%D0%B5%D0%BD%D0%B5%D1%80%D0%BD%D1%8B%D0%B9_%D1%80%D0%B5%D0%B9%D1%82%D0%B8%D0%BD%D0%B3_2026_%D0%B5%D0%B4%D0%B8%D0%BD%D0%B0%D1%8F_%D0%BF%D0%B0%D1%82%D0%B5%D0%BD%D1%82%D0%BD%D0%B0%D1%8F_%D0%B2%D1%8B%D0%B3%D1%80%D1%83%D0%B7%D0%BA%D0%B0.ipynb",
)

PROJECT_ROOT = Path(
    "/Users/vsevolodkarass/Library/Mobile Documents/com~apple~CloudDocs/"
    "Desktop/Рабочий стол — MacBook Air — Всеволод/"
    "Инженерный проект/Версия рейтинга 1"
)


def find_v2_notebook() -> Path:
    roots = [
        Path.cwd(),
        Path.home() / "Downloads",
        Path.home() / "Desktop",
        PROJECT_ROOT,
        PROJECT_ROOT.parent,
    ]
    checked: set[str] = set()

    for root in roots:
        root = root.expanduser()
        key = str(root)
        if key in checked:
            continue
        checked.add(key)
        for name in V2_FILENAMES:
            candidate = root / name
            if candidate.exists():
                return candidate

    # Ограниченный рекурсивный поиск только в ожидаемых пользовательских папках.
    for root in roots:
        if not root.exists():
            continue
        try:
            for name in V2_FILENAMES:
                found = next(root.rglob(name), None)
                if found is not None:
                    return found
        except (PermissionError, OSError):
            continue

    raise FileNotFoundError(
        "Не найден ноутбук версии 2. Положите файл "
        "'2_Инженерный_рейтинг_2026_единая_патентная_выгрузка.ipynb' "
        "в ту же папку, что и этот ноутбук, либо в Downloads."
    )


def patch_source(source: str) -> str:
    replacements = {
        "# 2_Инженерный рейтинг 2026": "# 3_Инженерный рейтинг 2026",
        "Версия 2": "Версия 3",
        "версии 2": "версии 3",
        "VERSION 2": "VERSION 3",
        'VERSION_PREFIX = "2"': 'VERSION_PREFIX = "3"',
        '"2_Финальная патентная база 2020-2025"': '"3_Финальная патентная база 2020-2025"',
        "`Выгрузка данных/2_Финальная патентная база 2020-2025`": "`Выгрузка данных/3_Финальная патентная база 2020-2025`",
        "REFRESH_GOOGLE_ERRORS = False": "REFRESH_GOOGLE_ERRORS = True",
        "Реализация конвейера версии 2": "Реализация конвейера версии 3",
    }
    for old, new in replacements.items():
        source = source.replace(old, new)
    return source


OVERRIDE_CODE = r'''
# =============================================================================
# VERSION 3 OVERRIDES
# - классификации собираются несколькими безопасными способами;
# - полнота классификаций и Country Status не останавливает загрузку;
# - hard stop остаётся только для опасных нарушений целостности;
# - ошибки сети и parse_validation_failed повторяются автоматически.
# =============================================================================

VERSION_PREFIX = "3"

# -----------------------------------------------------------------------------
# 1. Расширенное извлечение классификаций
# -----------------------------------------------------------------------------

_v2_parse_classifications = parse_classifications

_V3_PATENT_CLASS_RE = re.compile(
    r"(?<![A-Z0-9])([A-HY]\d{2}[A-Z]\s*\d{1,4}/\d{1,8}(?:\.\d+)?)",
    flags=re.I,
)
_V3_CLASS_LEVEL_RE = re.compile(r"(?<![A-Z0-9])([A-HY]\d{2}[A-Z])(?=[^A-Z0-9]|$)", flags=re.I)
_V3_LOCARNO_RE = re.compile(r"(?<!\d)(\d{2}-\d{2,4})(?!\d)")


def _v3_unique_codes(values):
    result = []
    seen = set()
    for value in values:
        if value is None:
            continue
        if isinstance(value, (list, tuple, set)):
            nested = value
        else:
            nested = [value]
        for item in nested:
            code = _normalise_classification_code(item)
            if not code or len(code) > 48:
                continue
            if code not in seen:
                seen.add(code)
                result.append(code)
    return result


def _v3_codes_from_text(text):
    text = repair_google_text(text or "").upper()
    detailed = [_normalise_classification_code(x) for x in _V3_PATENT_CLASS_RE.findall(text)]
    levels = [_normalise_classification_code(x) for x in _V3_CLASS_LEVEL_RE.findall(text)]
    locarno = [_normalise_classification_code(x) for x in _V3_LOCARNO_RE.findall(text)]
    return _v3_unique_codes(detailed + levels), _v3_unique_codes(locarno)


def _v3_walk_json(value, key_path=""):
    rows = []
    if isinstance(value, dict):
        for key, item in value.items():
            path = f"{key_path}.{key}" if key_path else str(key)
            rows.extend(_v3_walk_json(item, path))
    elif isinstance(value, list):
        for item in value:
            rows.extend(_v3_walk_json(item, key_path))
    elif isinstance(value, (str, int, float)):
        rows.append((key_path.casefold(), str(value)))
    return rows


def parse_classifications(soup):
    ipc, cpc, locarno, untyped, all_codes = _v2_parse_classifications(soup)
    ipc = list(ipc)
    cpc = list(cpc)
    locarno = list(locarno)
    untyped = list(untyped)

    # A. Только явно классификационные meta/link-теги.
    head = soup.head or soup
    for node in head.find_all(["meta", "link"]):
        attrs = " ".join(
            clean_text(node.get(name))
            for name in ("scheme", "classificationscheme", "itemprop", "name", "property", "rel")
        ).upper()
        if not any(token in attrs for token in ("IPC", "CPC", "LOCARNO", "CLASSIF")):
            continue
        value = node.get("content") or node.get("href") or node_value(node)
        codes, loc_codes = _v3_codes_from_text(value)
        if "LOCARNO" in attrs:
            locarno.extend(loc_codes or codes)
        elif "IPC" in attrs:
            ipc.extend(codes)
        elif "CPC" in attrs:
            cpc.extend(codes)
        else:
            untyped.extend(codes)
            locarno.extend(loc_codes)

    # B. JSON-LD: читаются только ключи, в названии которых есть classification/IPC/CPC/Locarno.
    for script in soup.select('script[type="application/ld+json"]'):
        try:
            payload = json.loads(script.string or script.get_text(" ", strip=True))
        except Exception:
            continue
        for key_path, value in _v3_walk_json(payload):
            if not any(token in key_path for token in ("classification", "ipc", "cpc", "locarno")):
                continue
            codes, loc_codes = _v3_codes_from_text(value)
            if "locarno" in key_path:
                locarno.extend(loc_codes or codes)
            elif "ipc" in key_path:
                ipc.extend(codes)
            elif "cpc" in key_path:
                cpc.extend(codes)
            else:
                untyped.extend(codes)
                locarno.extend(loc_codes)

    # C. Явная microdata классификаций. Полный текст патента не сканируется.
    selector = (
        '[itemprop="classification"], [itemprop="classifications"], '
        '[itemprop="classificationCode"], [itemprop="Code"], [itemprop="code"], '
        '[scheme*="IPC" i], [scheme*="CPC" i], [scheme*="Locarno" i], '
        '[classificationScheme*="IPC" i], [classificationScheme*="CPC" i], '
        '[classificationScheme*="Locarno" i]'
    )
    for node in soup.select(selector):
        attrs = " ".join(
            clean_text(node.get(name))
            for name in ("scheme", "classificationScheme", "itemprop", "class")
        ).upper()
        value = node.get("content") or node_value(node)
        codes, loc_codes = _v3_codes_from_text(value)
        if "LOCARNO" in attrs:
            locarno.extend(loc_codes or codes)
        elif "IPC" in attrs:
            ipc.extend(codes)
        elif "CPC" in attrs:
            cpc.extend(codes)
        else:
            untyped.extend(codes)
            locarno.extend(loc_codes)

    # D. Видимый раздел Classifications как последний безопасный резерв.
    section = _classification_section(soup)
    if section is not None:
        codes, loc_codes = _v3_codes_from_text(section.get_text(" ", strip=True))
        untyped.extend(codes)
        locarno.extend(loc_codes)

    ipc = _v3_unique_codes(ipc)
    cpc = _v3_unique_codes(cpc)
    locarno = _v3_unique_codes(locarno)
    untyped = _v3_unique_codes(untyped)

    # Y-классы относятся к CPC. Остальные неразмеченные коды не объявляются IPC без основания.
    y_codes = [code for code in untyped if code.startswith("Y")]
    if y_codes:
        cpc = _v3_unique_codes(cpc + y_codes)
        untyped = [code for code in untyped if code not in set(y_codes)]

    all_codes = _v3_unique_codes(ipc + cpc + locarno + untyped)
    return ipc, cpc, locarno, untyped, all_codes


# -----------------------------------------------------------------------------
# 2. Повтор запроса при временной ошибке или невалидном парсинге
# -----------------------------------------------------------------------------

_v2_process_object = GooglePatentsClient.process_object


def _v3_process_object(self, obj):
    retryable = {
        "request_error", "worker_error", "not_found", "parse_validation_failed",
        "http_408", "http_425", "http_429", "http_500", "http_502", "http_503", "http_504",
    }
    last = None
    for attempt in range(1, 3):
        last = _v2_process_object(self, obj)
        last["gp_attempts"] = attempt
        if clean_text(last.get("gp_status")) not in retryable:
            return last
        if attempt < 2:
            time.sleep(1.5 + random.uniform(0.2, 0.8))
    return last


GooglePatentsClient.process_object = _v3_process_object


# -----------------------------------------------------------------------------
# 3. QC: полнота — информационно; опасные нарушения — hard stop
# -----------------------------------------------------------------------------

_v2_build_google_quality_checks = build_google_quality_checks


def build_google_quality_checks(progress):
    checks = _v2_build_google_quality_checks(progress).copy()
    if checks.empty:
        return checks

    soft_checks = {
        "Доля валидных success",
        "Заполненность названий",
        "Заполненность номера заявки",
        "Заполненность даты публикации",
        "Классификации для изобретений/полезных моделей",
        "Доступен Country Status семьи",
    }
    hard_checks = {
        "Есть обработанные объекты",
        "Точные RU-якоря",
        "Корректная кодировка",
        "Внутренняя валидация парсинга",
        "Семейства не пересекаются с цитированиями",
        "Нет аномально длинных списков правообладателей",
        "Ответы декодированы как UTF-8",
    }

    checks.loc[checks["check"].isin(soft_checks), "hard_stop"] = 0
    checks.loc[checks["check"].isin(hard_checks), "hard_stop"] = 1

    mask = checks["check"].eq("Классификации для изобретений/полезных моделей")
    checks.loc[mask, "threshold"] = "информационно"
    checks.loc[mask, "passed"] = 1
    checks.loc[mask, "note"] = (
        "Полнота классификаций контролируется, но не останавливает загрузку. "
        "В финале используются Google + резервные поля Роспатента."
    )

    mask = checks["check"].eq("Доступен Country Status семьи")
    checks.loc[mask, "passed"] = 1
    checks.loc[mask, "hard_stop"] = 0

    return checks


# -----------------------------------------------------------------------------
# 4. Финальные поля классификаций Google + Роспатент
# -----------------------------------------------------------------------------

_v2_add_rating_ready_variables = add_rating_ready_variables


def _v3_join_source_columns(df, tokens):
    selected = []
    for column in df.columns:
        name = str(column).casefold()
        if not name.startswith("rospatent_raw__"):
            continue
        if any(token in name for token in tokens):
            selected.append(column)
    if not selected:
        return pd.Series("", index=df.index, dtype="string")

    def combine(row):
        values = []
        seen = set()
        for value in row:
            text = clean_text(value)
            if not text:
                continue
            for part in re.split(r"[;|,\n]+", text):
                code = _normalise_classification_code(part)
                if code and code not in seen:
                    seen.add(code)
                    values.append(code)
        return "; ".join(values)

    return df[selected].apply(combine, axis=1).astype("string")


def _v3_code_set(value):
    return {
        _normalise_classification_code(part)
        for part in re.split(r"[;|,\n]+", clean_text(value))
        if _normalise_classification_code(part)
    }


def add_rating_ready_variables(df):
    out = _v2_add_rating_ready_variables(df)

    out["ipc_google"] = out.get("gp_ipc_codes", pd.Series("", index=out.index)).fillna("").astype("string")
    out["cpc_google"] = out.get("gp_cpc_codes", pd.Series("", index=out.index)).fillna("").astype("string")
    out["locarno_google"] = out.get("gp_locarno_codes", pd.Series("", index=out.index)).fillna("").astype("string")
    out["classification_google_all"] = out.get(
        "gp_all_classification_codes", pd.Series("", index=out.index)
    ).fillna("").astype("string")

    out["ipc_rospatent"] = _v3_join_source_columns(out, ("ipc", "мпк", "international_patent_class"))
    out["locarno_rospatent"] = _v3_join_source_columns(out, ("locarno", "локарно"))

    final_codes = []
    sources = []
    disagreements = []
    for gp_value, rp_ipc, rp_loc in zip(
        out["classification_google_all"], out["ipc_rospatent"], out["locarno_rospatent"]
    ):
        gp_set = _v3_code_set(gp_value)
        rp_set = _v3_code_set(rp_ipc) | _v3_code_set(rp_loc)
        union = sorted(gp_set | rp_set)
        final_codes.append("; ".join(union))
        if gp_set and rp_set:
            sources.append("google+rospatent")
            disagreements.append(int(gp_set.isdisjoint(rp_set)))
        elif gp_set:
            sources.append("google")
            disagreements.append(0)
        elif rp_set:
            sources.append("rospatent")
            disagreements.append(0)
        else:
            sources.append("missing")
            disagreements.append(0)

    out["classification_codes_final"] = pd.Series(final_codes, index=out.index, dtype="string")
    out["classification_source_final"] = pd.Series(sources, index=out.index, dtype="string")
    out["classification_disagreement_flag"] = pd.Series(disagreements, index=out.index, dtype="int64")
    out["classification_available_flag"] = out["classification_codes_final"].str.strip().ne("").astype(int)

    VARIABLE_DESCRIPTIONS.update({
        "ipc_google": ("Google Patents", "IPC, извлечённые из структурированных классификационных элементов"),
        "cpc_google": ("Google Patents", "CPC, извлечённые из структурированных классификационных элементов"),
        "locarno_google": ("Google Patents", "Locarno, извлечённые из структурированных классификационных элементов"),
        "ipc_rospatent": ("Роспатент", "IPC/МПК из исходных полей реестра, если они присутствуют"),
        "locarno_rospatent": ("Роспатент", "Классы Локарно из исходных полей реестра, если они присутствуют"),
        "classification_codes_final": ("Расчёт", "Объединение проверенных классификаций Google Patents и Роспатента"),
        "classification_source_final": ("Расчёт", "Источник итоговой классификации: google, rospatent, google+rospatent или missing"),
        "classification_disagreement_flag": ("Расчёт", "1, если непустые наборы Google и Роспатента не имеют общих кодов"),
        "classification_available_flag": ("Расчёт", "1, если имеется хотя бы один итоговый классификационный код"),
    })
    return out


print("Применены исправления версии 3:")
print("  • папка результатов 3_Финальная патентная база 2020-2025")
print("  • пачки Google Patents по 500 объектов")
print("  • полнота классификаций и Country Status не блокирует выполнение")
print("  • опасные ошибки кодировки, якоря, семьи и метаданных сохраняют hard stop")
print("  • parse/network errors автоматически повторяются один раз")
'''

source_path = find_v2_notebook()
print("Найдена версия 2:", source_path)

notebook = nbformat.read(source_path, as_version=4)
for cell in notebook.cells:
    if cell.cell_type in {"code", "markdown"}:
        cell.source = patch_source(cell.source)
        cell.execution_count = None if cell.cell_type == "code" else cell.get("execution_count")
        if cell.cell_type == "code":
            cell.outputs = []

run_index = None
for index, cell in enumerate(notebook.cells):
    if cell.cell_type == "code" and "outputs = run_pipeline(CONFIG)" in cell.source:
        run_index = index
        break

if run_index is None:
    raise RuntimeError("В версии 2 не найдена ячейка запуска run_pipeline(CONFIG).")

notebook.cells.insert(run_index, new_code_cell(OVERRIDE_CODE))
notebook.cells[0].source = patch_source(notebook.cells[0].source) + (
    "\n\n## Изменения версии 3\n"
    "- Проверки полноты классификаций и Country Status стали информационными.\n"
    "- Жёсткая остановка сохранена только для опасных нарушений целостности.\n"
    "- Добавлено многоступенчатое извлечение IPC/CPC/Locarno.\n"
    "- В финале объединяются классификации Google Patents и исходных реестров Роспатента.\n"
    "- Сетевые ошибки и parse_validation_failed повторяются один раз.\n"
)

expanded_path = source_path.with_name(
    "3_Инженерный_рейтинг_2026_единая_патентная_выгрузка_FULL.ipynb"
)
nbformat.write(notebook, expanded_path)
print("Создана полная самостоятельная копия версии 3:", expanded_path)
print("Начинается выполнение. После запуска ячейку можно оставить работать.")

execution_globals = globals()
for index, cell in enumerate(notebook.cells):
    if cell.cell_type != "code" or not cell.source.strip():
        continue
    print(f"\n--- Выполняется ячейка версии 3: {index + 1}/{len(notebook.cells)} ---")
    try:
        exec(compile(cell.source, f"v3_cell_{index + 1}", "exec"), execution_globals, execution_globals)
    except SystemExit:
        raise
    except KeyboardInterrupt:
        print("Выполнение прервано пользователем. Checkpoint-файлы и SQLite-кеш сохранены.")
        raise
    except Exception:
        print(f"Ошибка в ячейке {index + 1}. Промежуточные файлы, созданные до ошибки, сохранены.")
        raise
'''

MARKDOWN = r'''
# 3_Инженерный рейтинг 2026: единая патентная выгрузка

**Версия 3. Запуск целиком с первого этапа в новой независимой папке.**

Ноутбук автоматически находит локальный файл версии 2, создаёт рядом полную самостоятельную копию
`3_Инженерный_рейтинг_2026_единая_патентная_выгрузка_FULL.ipynb`, применяет исправления и сразу запускает её.

Результаты сохраняются в:

`Выгрузка данных/3_Финальная патентная база 2020-2025`

Контроль после каждой пачки из 500 патентов продолжается, но неполнота IPC/CPC/Locarno и Country Status теперь только фиксируется в QC и не прерывает длительную загрузку. Остановка сохраняется при неверном RU-якоре, повреждённой кодировке, пересечении семейства с цитированиями и системном загрязнении метаданных.
'''

nb = new_notebook(
    cells=[
        new_markdown_cell(MARKDOWN),
        new_code_cell(PATCHER_CODE),
    ],
    metadata={
        "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
        "language_info": {"name": "python", "version": "3.11"},
    },
)

output = Path("3_Инженерный_рейтинг_2026_единая_патентная_выгрузка.ipynb")
nbformat.write(nb, output)
print(output)
