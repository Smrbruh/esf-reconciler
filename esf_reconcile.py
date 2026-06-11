"""
esf_reconcile.py — Инструмент автоматической сверки ЭСФ между 1С и ИС «ЭСФ»
АО КазАгроФинанс
"""

import os
import sys
import re
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import yaml
from colorama import Fore, Style, init as colorama_init
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, numbers,
    Border, Side
)
from openpyxl.utils import get_column_letter

# ─── Инициализация colorama (поддержка Windows) ─────────────────────────────
colorama_init(autoreset=True)

# ─── Константы цветов для Excel ─────────────────────────────────────────────
CLR_HEADER_BG  = "1F3864"   # Тёмно-синий фон заголовков
CLR_HEADER_FG  = "FFFFFF"   # Белый текст заголовков
CLR_ONLY_1C    = "FFCCCC"   # Светло-красный: только в 1С
CLR_ONLY_ESF   = "FFE5CC"   # Светло-оранжевый: только в ИС ЭСФ
CLR_DUPLICATE  = "FFFFCC"   # Светло-жёлтый: дубликаты
CLR_MISMATCH   = "CCE5FF"   # Светло-синий: расхождения
CLR_MATCH      = "CCFFCC"   # Светло-зелёный: совпадения

# ─── Числовые форматы ───────────────────────────────────────────────────────
FMT_MONEY  = '#,##0.00'
FMT_DATE   = 'DD.MM.YYYY'
FMT_INT    = '#,##0'

# ─── Форматы дат, которые умеем парсить ─────────────────────────────────────
DATE_FORMATS = [
    "%d.%m.%Y",
    "%Y-%m-%d",
    "%d/%m/%Y",
    "%d-%m-%Y",
    "%Y.%m.%d",
]

logger = logging.getLogger(__name__)


# ════════════════════════════════════════════════════════════════════════════
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ВЫВОДА В КОНСОЛЬ
# ════════════════════════════════════════════════════════════════════════════

def log_ok(msg: str) -> None:
    print(f"{Fore.GREEN}✅  {msg}{Style.RESET_ALL}")

def log_info(msg: str) -> None:
    print(f"{Fore.CYAN}ℹ️   {msg}{Style.RESET_ALL}")

def log_warn(msg: str) -> None:
    print(f"{Fore.YELLOW}⚠️   {msg}{Style.RESET_ALL}")

def log_err(msg: str) -> None:
    print(f"{Fore.RED}❌  {msg}{Style.RESET_ALL}")

def log_step(msg: str) -> None:
    print(f"{Fore.MAGENTA}🔍  {msg}{Style.RESET_ALL}")

def log_save(msg: str) -> None:
    print(f"{Fore.BLUE}💾  {msg}{Style.RESET_ALL}")


# ════════════════════════════════════════════════════════════════════════════
# ДЕФОЛТНЫЙ КОНФИГ
# ════════════════════════════════════════════════════════════════════════════

DEFAULT_CONFIG = {
    "file_1c":  "input/1c_export.xlsx",
    "file_esf": "input/esf_export.xlsx",
    "key_column": {
        "in_1c":  "Номер ЭСФ",
        "in_esf": "Рег. номер",
    },
    "fields": {
        "Дата": {
            "col_1c":  "Дата документа",
            "col_esf": "Дата выписки",
        },
        "Сумма оборота": {
            "col_1c":  "Сумма без НДС",
            "col_esf": "Оборот",
        },
        "Сумма НДС": {
            "col_1c":  "НДС",
            "col_esf": "Сумма НДС",
        },
        "Итого с НДС": {
            "col_1c":  "Сумма с НДС",
            "col_esf": "Итого",
        },
        "ИНН контрагента": {
            "col_1c":  "БИН/ИИН",
            "col_esf": "БИН получателя",
        },
        "Контрагент": {
            "col_1c":  "Наименование контрагента",
            "col_esf": "Получатель",
        },
    },
    "amount_tolerance": 1.0,
    "csv_encoding":  "utf-8-sig",
    "csv_delimiter": ";",
    "output_dir": "output",
}

# Поля, которые считаются «суммовыми» (для применения погрешности)
AMOUNT_FIELD_KEYWORDS = ["сумма", "итого", "ндс", "оборот"]

# Поля, которые считаются «датными»
DATE_FIELD_KEYWORDS = ["дата"]


# ════════════════════════════════════════════════════════════════════════════
# 1. ЗАГРУЗКА КОНФИГА
# ════════════════════════════════════════════════════════════════════════════

def load_config(config_path: str = "config.yaml") -> dict:
    """Загружает config.yaml. Если файл не найден — создаёт дефолтный."""
    if not os.path.exists(config_path):
        log_warn(f"Файл конфигурации '{config_path}' не найден. Создаю файл с настройками по умолчанию.")
        with open(config_path, "w", encoding="utf-8") as fh:
            yaml.dump(
                DEFAULT_CONFIG,
                fh,
                allow_unicode=True,
                default_flow_style=False,
                sort_keys=False,
            )
        log_info(f"Файл '{config_path}' создан. Отредактируйте его и запустите скрипт снова.")
        sys.exit(0)

    with open(config_path, "r", encoding="utf-8") as fh:
        cfg = yaml.safe_load(fh)

    # Мержим с дефолтами, чтобы не падать на отсутствующих ключах
    merged = DEFAULT_CONFIG.copy()
    merged.update(cfg or {})
    return merged


# ════════════════════════════════════════════════════════════════════════════
# 2. ЗАГРУЗКА ФАЙЛОВ
# ════════════════════════════════════════════════════════════════════════════

def load_file(path: str, cfg: dict) -> pd.DataFrame:
    """
    Загружает Excel (.xlsx/.xls) или CSV файл в DataFrame.
    Убирает полностью пустые строки и столбцы.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(
            f"Файл не найден: {path}\n"
            "Проверьте путь в config.yaml (параметры file_1c и file_esf)."
        )

    suffix = path.suffix.lower()
    try:
        if suffix in (".xlsx", ".xls"):
            df = pd.read_excel(path, dtype=str)
        elif suffix == ".csv":
            df = pd.read_csv(
                path,
                dtype=str,
                encoding=cfg.get("csv_encoding", "utf-8-sig"),
                sep=cfg.get("csv_delimiter", ";"),
            )
        else:
            raise ValueError(
                f"Неподдерживаемый формат файла: '{suffix}'. "
                "Поддерживаются: .xlsx, .xls, .csv"
            )
    except Exception as exc:
        raise RuntimeError(f"Ошибка при чтении файла '{path}': {exc}") from exc

    # Убираем полностью пустые строки и столбцы
    df = df.dropna(how="all").dropna(axis=1, how="all")

    # Убираем строки, в которых все значения — одно и то же (повторные заголовки)
    df = _drop_repeated_headers(df)

    # Сбрасываем индекс
    df = df.reset_index(drop=True)
    return df


def _drop_repeated_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    Удаляет строки, которые совпадают с заголовком (т.е. строки-дубли шапки).
    Такое бывает при выгрузке из 1С с разбивкой по страницам.
    """
    if df.empty:
        return df
    header_vals = set(str(v).strip() for v in df.columns)
    mask = df.apply(
        lambda row: set(str(v).strip() for v in row.values) == header_vals,
        axis=1
    )
    return df[~mask]


# ════════════════════════════════════════════════════════════════════════════
# 3. НОРМАЛИЗАЦИЯ
# ════════════════════════════════════════════════════════════════════════════

def normalize_dates(series: pd.Series) -> pd.Series:
    """
    Приводит серию с датами к единому формату pd.Timestamp.
    Пробует несколько форматов, при неуспехе возвращает NaT.
    """
    def parse_one(val):
        if pd.isna(val) or str(val).strip() == "":
            return pd.NaT
        val_str = str(val).strip()
        # Если уже timestamp
        if isinstance(val, (pd.Timestamp, datetime)):
            return pd.Timestamp(val)
        for fmt in DATE_FORMATS:
            try:
                return datetime.strptime(val_str, fmt)
            except ValueError:
                continue
        # Последняя попытка — pandas
        try:
            return pd.to_datetime(val_str, dayfirst=True)
        except Exception:
            return pd.NaT

    return series.apply(parse_one)


def normalize_amounts(series: pd.Series) -> pd.Series:
    """
    Приводит серию со строковыми суммами к float.
    Убирает пробелы, заменяет запятую на точку.
    """
    def parse_one(val):
        if pd.isna(val) or str(val).strip() == "":
            return None
        val_str = str(val).strip()
        # Убираем пробелы (разделители тысяч в русской локали), меняем запятую
        val_str = val_str.replace(" ", "").replace("\xa0", "").replace(",", ".")
        # Убираем лишние символы валюты
        val_str = re.sub(r"[₸₽$€]", "", val_str).strip()
        try:
            return float(val_str)
        except ValueError:
            return None

    return series.apply(parse_one)


def normalize_string(val) -> str:
    """Trim whitespace + lowercase для регистронезависимого сравнения строк."""
    if pd.isna(val):
        return ""
    return str(val).strip().lower()


def normalize_key(val) -> str:
    """Нормализует ключевой номер ЭСФ: убирает пробелы, приводит к строке."""
    if pd.isna(val):
        return ""
    return str(val).strip()


# ════════════════════════════════════════════════════════════════════════════
# 4. ОСНОВНАЯ СВЕРКА
# ════════════════════════════════════════════════════════════════════════════

def compare_records(df_1c: pd.DataFrame, df_esf: pd.DataFrame, cfg: dict) -> dict:
    """
    Выполняет сверку двух датафреймов.

    Возвращает словарь:
        {
            'only_1c':    DataFrame,
            'only_esf':   DataFrame,
            'dup_1c':     DataFrame,
            'dup_esf':    DataFrame,
            'mismatches': DataFrame,
            'matches':    DataFrame,
        }
    """
    key_1c  = cfg["key_column"]["in_1c"]
    key_esf = cfg["key_column"]["in_esf"]
    fields  = cfg["fields"]
    tolerance = float(cfg.get("amount_tolerance", 1.0))

    # Проверяем наличие ключевого столбца
    _check_column(df_1c, key_1c, "1С")
    _check_column(df_esf, key_esf, "ИС ЭСФ")

    # Нормализуем ключи
    df_1c  = df_1c.copy()
    df_esf = df_esf.copy()
    df_1c["_key"]  = df_1c[key_1c].apply(normalize_key)
    df_esf["_key"] = df_esf[key_esf].apply(normalize_key)

    # Убираем строки с пустым ключом
    empty_1c  = (df_1c["_key"] == "").sum()
    empty_esf = (df_esf["_key"] == "").sum()
    if empty_1c:
        log_warn(f"В файле 1С найдено {empty_1c} строк с пустым номером ЭСФ — пропускаем.")
    if empty_esf:
        log_warn(f"В файле ИС ЭСФ найдено {empty_esf} строк с пустым номером ЭСФ — пропускаем.")
    df_1c  = df_1c[df_1c["_key"] != ""]
    df_esf = df_esf[df_esf["_key"] != ""]

    # ── Дубликаты ────────────────────────────────────────────────────────────
    dup_keys_1c  = df_1c[df_1c.duplicated(subset="_key", keep=False)]["_key"].unique()
    dup_keys_esf = df_esf[df_esf.duplicated(subset="_key", keep=False)]["_key"].unique()

    dup_1c_rows  = df_1c[df_1c["_key"].isin(dup_keys_1c)].copy()
    dup_esf_rows = df_esf[df_esf["_key"].isin(dup_keys_esf)].copy()

    # ── Для сравнения берём первую копию каждого дубля ───────────────────────
    df_1c_dedup  = df_1c.drop_duplicates(subset="_key", keep="first")
    df_esf_dedup = df_esf.drop_duplicates(subset="_key", keep="first")

    keys_1c  = set(df_1c_dedup["_key"])
    keys_esf = set(df_esf_dedup["_key"])

    only_1c_keys  = keys_1c  - keys_esf
    only_esf_keys = keys_esf - keys_1c
    common_keys   = keys_1c  & keys_esf

    # ── Только в одной системе ───────────────────────────────────────────────
    only_1c  = df_1c_dedup[df_1c_dedup["_key"].isin(only_1c_keys)].copy()
    only_esf = df_esf_dedup[df_esf_dedup["_key"].isin(only_esf_keys)].copy()

    # ── Общие записи: сравниваем реквизиты ──────────────────────────────────
    common_1c  = df_1c_dedup[df_1c_dedup["_key"].isin(common_keys)].set_index("_key")
    common_esf = df_esf_dedup[df_esf_dedup["_key"].isin(common_keys)].set_index("_key")

    mismatch_rows = []
    match_keys    = []

    for key in sorted(common_keys):
        row_1c  = common_1c.loc[key]
        row_esf = common_esf.loc[key]
        has_diff = False

        for field_name, col_map in fields.items():
            col_1c_name  = col_map["col_1c"]
            col_esf_name = col_map["col_esf"]

            # Получаем значения (None если колонки нет)
            val_1c  = _get_val(row_1c,  col_1c_name)
            val_esf = _get_val(row_esf, col_esf_name)

            # Определяем тип поля
            field_lower = field_name.lower()
            is_date   = any(kw in field_lower for kw in DATE_FIELD_KEYWORDS)
            is_amount = any(kw in field_lower for kw in AMOUNT_FIELD_KEYWORDS) and not is_date

            diff = False
            if is_date:
                diff = _dates_differ(val_1c, val_esf)
            elif is_amount:
                diff = _amounts_differ(val_1c, val_esf, tolerance)
            else:
                diff = _strings_differ(val_1c, val_esf)

            if diff:
                has_diff = True
                # Дата из 1С для отчёта
                date_1c_raw  = _get_val(row_1c,  _first_date_col(fields, "col_1c"))
                date_esf_raw = _get_val(row_esf, _first_date_col(fields, "col_esf"))
                mismatch_rows.append({
                    "Номер ЭСФ":          key,
                    "Дата (1С)":          date_1c_raw,
                    "Дата (ИС ЭСФ)":     date_esf_raw,
                    "Реквизит":           field_name,
                    "Значение в 1С":      val_1c,
                    "Значение в ИС ЭСФ":  val_esf,
                })

        if not has_diff:
            match_keys.append(key)

    matches = common_1c.loc[match_keys].reset_index() if match_keys else pd.DataFrame()

    mismatches = pd.DataFrame(mismatch_rows) if mismatch_rows else pd.DataFrame(
        columns=["Номер ЭСФ", "Дата (1С)", "Дата (ИС ЭСФ)", "Реквизит",
                 "Значение в 1С", "Значение в ИС ЭСФ"]
    )

    # ── Формируем таблицу дубликатов ─────────────────────────────────────────
    dup_1c_out  = _build_dup_table(dup_1c_rows,  "1С",      key_1c,  fields)
    dup_esf_out = _build_dup_table(dup_esf_rows, "ИС ЭСФ",  key_esf, fields)
    dup_all     = pd.concat([dup_1c_out, dup_esf_out], ignore_index=True)

    return {
        "only_1c":    only_1c,
        "only_esf":   only_esf,
        "dup_all":    dup_all,
        "mismatches": mismatches,
        "matches":    matches,
        "df_1c":      df_1c,
        "df_esf":     df_esf,
        "cfg":        cfg,
    }


# ─── Вспомогательные функции сравнения ──────────────────────────────────────

def _get_val(row, col_name: str):
    """Возвращает значение ячейки или None, если колонка отсутствует."""
    if col_name in row.index:
        v = row[col_name]
        return None if pd.isna(v) else v
    return None


def _dates_differ(val_1c, val_esf) -> bool:
    """True если даты отличаются (или одна из них пустая, а другая нет)."""
    ts_1c  = normalize_dates(pd.Series([val_1c]))[0]
    ts_esf = normalize_dates(pd.Series([val_esf]))[0]
    # Обе пустые — не расхождение
    if pd.isna(ts_1c) and pd.isna(ts_esf):
        return False
    # Одна пустая — расхождение
    if pd.isna(ts_1c) or pd.isna(ts_esf):
        return True
    return ts_1c.date() != ts_esf.date()


def _amounts_differ(val_1c, val_esf, tolerance: float) -> bool:
    """True если суммы отличаются больше чем на tolerance."""
    a = normalize_amounts(pd.Series([val_1c]))[0]
    b = normalize_amounts(pd.Series([val_esf]))[0]
    if a is None and b is None:
        return False
    if a is None or b is None:
        return True
    return abs(a - b) > tolerance


def _strings_differ(val_1c, val_esf) -> bool:
    """True если строки отличаются (регистронезависимо, без пробелов по краям)."""
    return normalize_string(val_1c) != normalize_string(val_esf)


def _check_column(df: pd.DataFrame, col: str, source: str) -> None:
    """Проверяет наличие колонки. Если нет — понятная ошибка."""
    if col not in df.columns:
        available = ", ".join(f'"{c}"' for c in df.columns[:10])
        raise ValueError(
            f"В файле {source} не найдена ключевая колонка «{col}».\n"
            f"Доступные колонки: {available}...\n"
            "Исправьте параметр key_column в config.yaml."
        )


def _first_date_col(fields: dict, side: str) -> str:
    """Возвращает название колонки с датой (первой найденной)."""
    for fname, cols in fields.items():
        if "дата" in fname.lower():
            return cols.get(side, "")
    return ""


def _build_dup_table(df_dups: pd.DataFrame, source: str, key_col: str, fields: dict) -> pd.DataFrame:
    """Формирует таблицу дубликатов с количеством повторений."""
    if df_dups.empty:
        return pd.DataFrame(columns=["Источник", "Номер ЭСФ", "Кол-во повторений", "Дата", "Сумма"])

    date_col = ""
    sum_col  = ""
    for fname, cols in fields.items():
        side = "col_1c" if source == "1С" else "col_esf"
        if "дата" in fname.lower() and not date_col:
            date_col = cols.get(side, "")
        if "итого" in fname.lower() and not sum_col:
            sum_col = cols.get(side, "")

    rows = []
    for key, group in df_dups.groupby("_key"):
        first = group.iloc[0]
        rows.append({
            "Источник":          source,
            "Номер ЭСФ":         key,
            "Кол-во повторений": len(group),
            "Дата":              first.get(date_col, "") if date_col and date_col in first.index else "",
            "Сумма":             first.get(sum_col, "")  if sum_col  and sum_col  in first.index else "",
        })
    return pd.DataFrame(rows)


# ════════════════════════════════════════════════════════════════════════════
# 5. ГЕНЕРАЦИЯ EXCEL-ОТЧЁТА
# ════════════════════════════════════════════════════════════════════════════

def generate_report(results: dict, output_dir: str) -> str:
    """
    Генерирует Excel-отчёт с 6 листами.
    Возвращает путь к созданному файлу.
    """
    os.makedirs(output_dir, exist_ok=True)
    now_str    = datetime.now().strftime("%Y-%m-%d_%H-%M")
    out_path   = os.path.join(output_dir, f"protocol_{now_str}.xlsx")

    cfg        = results["cfg"]
    fields     = cfg["fields"]

    # Определяем исходные колонки для основных листов
    key_1c  = cfg["key_column"]["in_1c"]
    key_esf = cfg["key_column"]["in_esf"]

    only_1c    = results["only_1c"]
    only_esf   = results["only_esf"]
    dup_all    = results["dup_all"]
    mismatches = results["mismatches"]
    matches    = results["matches"]
    df_1c      = results["df_1c"]
    df_esf     = results["df_esf"]

    wb = Workbook()
    wb.remove(wb.active)  # Удаляем стандартный лист

    # ── Лист 1: Сводка ───────────────────────────────────────────────────────
    _sheet_summary(wb, df_1c, df_esf, only_1c, only_esf, dup_all, mismatches, matches)

    # ── Лист 2: Только в 1С ──────────────────────────────────────────────────
    cols_1c = _esf_display_cols(only_1c, fields, "col_1c", key_1c)
    _sheet_esf_list(wb, "Только в 1С",   only_1c,  cols_1c, CLR_ONLY_1C)

    # ── Лист 3: Только в ИС ЭСФ ──────────────────────────────────────────────
    cols_esf = _esf_display_cols(only_esf, fields, "col_esf", key_esf)
    _sheet_esf_list(wb, "Только в ИС ЭСФ", only_esf, cols_esf, CLR_ONLY_ESF)

    # ── Лист 4: Дубликаты ────────────────────────────────────────────────────
    _sheet_duplicates(wb, dup_all)

    # ── Лист 5: Расхождения ───────────────────────────────────────────────────
    _sheet_mismatches(wb, mismatches)

    # ── Лист 6: Совпадения ───────────────────────────────────────────────────
    cols_match = _esf_display_cols(matches, fields, "col_1c", key_1c)
    _sheet_esf_list(wb, "Полностью совпадают", matches, cols_match, CLR_MATCH)

    wb.save(out_path)
    return out_path


# ─── Лист 1: Сводка ─────────────────────────────────────────────────────────

def _sheet_summary(wb, df_1c, df_esf, only_1c, only_esf, dup_all, mismatches, matches):
    ws = wb.create_sheet("Сводка")

    n_1c    = len(df_1c["_key"].unique())
    n_esf   = len(df_esf["_key"].unique())
    n_only1c   = len(only_1c)
    n_only_esf = len(only_esf)

    dup_1c_cnt  = dup_all[dup_all["Источник"] == "1С"]["Кол-во повторений"].sum()  \
                  if not dup_all.empty else 0
    dup_esf_cnt = dup_all[dup_all["Источник"] == "ИС ЭСФ"]["Кол-во повторений"].sum() \
                  if not dup_all.empty else 0
    dup_1c_keys  = len(dup_all[dup_all["Источник"] == "1С"])   if not dup_all.empty else 0
    dup_esf_keys = len(dup_all[dup_all["Источник"] == "ИС ЭСФ"]) if not dup_all.empty else 0

    n_mis   = mismatches["Номер ЭСФ"].nunique() if not mismatches.empty else 0
    n_match = len(matches)

    rows_data = [
        ("Показатель",                     "Количество"),
        ("Всего в 1С (уникальных номеров)", n_1c),
        ("Всего в ИС ЭСФ (уникальных номеров)", n_esf),
        ("Только в 1С",                    n_only1c),
        ("Только в ИС ЭСФ",               n_only_esf),
        (f"Дубликаты в 1С (номеров / записей)", f"{dup_1c_keys} / {int(dup_1c_cnt)}"),
        (f"Дубликаты в ИС ЭСФ (номеров / записей)", f"{dup_esf_keys} / {int(dup_esf_cnt)}"),
        ("Расхождения по реквизитам",       n_mis),
        ("Полностью совпадают",             n_match),
        ("", ""),
        ("Дата и время формирования протокола",
         datetime.now().strftime("%d.%m.%Y %H:%M:%S")),
    ]

    # Заголовок
    ws.merge_cells("A1:B1")
    title_cell = ws["A1"]
    title_cell.value = "ПРОТОКОЛ СВЕРКИ ЭСФ"
    title_cell.font  = Font(name="Calibri", bold=True, size=14, color=CLR_HEADER_FG)
    title_cell.fill  = PatternFill("solid", fgColor=CLR_HEADER_BG)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for i, (label, value) in enumerate(rows_data, start=2):
        ws.cell(row=i, column=1, value=label)
        ws.cell(row=i, column=2, value=value)

        label_cell = ws.cell(row=i, column=1)
        value_cell = ws.cell(row=i, column=2)

        if label == "Показатель":
            for cell in [label_cell, value_cell]:
                cell.font = Font(name="Calibri", bold=True, color=CLR_HEADER_FG)
                cell.fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
        else:
            label_cell.font = Font(name="Calibri", bold=True)
            value_cell.font = Font(name="Calibri")

        for cell in [label_cell, value_cell]:
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = _thin_border()

    ws.column_dimensions["A"].width = 48
    ws.column_dimensions["B"].width = 22
    ws.freeze_panes = "A3"


# ─── Листы 2, 3, 6 (Списки ЭСФ) ─────────────────────────────────────────────

def _esf_display_cols(df: pd.DataFrame, fields: dict, side: str, key_col: str) -> list:
    """Определяет список (orig_col, display_name) для вывода в листе."""
    cols = [(key_col, "Номер ЭСФ")]
    mapping = {
        "Дата":          "Дата",
        "Сумма оборота": "Сумма оборота",
        "Сумма НДС":     "Сумма НДС",
        "Итого с НДС":   "Итого",
        "ИНН контрагента": "ИНН",
        "Контрагент":    "Контрагент",
    }
    for field_name, display in mapping.items():
        if field_name in fields:
            orig = fields[field_name][side]
            if orig and (df.empty or orig in df.columns):
                cols.append((orig, display))
    return cols


def _sheet_esf_list(wb, sheet_name: str, df: pd.DataFrame, cols: list, row_color: str):
    """Создаёт лист со списком ЭСФ (листы 2, 3, 6)."""
    ws = wb.create_sheet(sheet_name)
    headers = [display for _, display in cols]
    _write_header_row(ws, headers, row=1)

    if df.empty:
        ws.cell(row=2, column=1, value="Нет данных")
        _autofit(ws, headers)
        return

    fill = PatternFill("solid", fgColor=row_color)
    date_cols    = {"Дата"}
    amount_cols  = {"Сумма оборота", "Сумма НДС", "Итого"}

    for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
        for c_idx, (orig_col, display) in enumerate(cols, start=1):
            raw_val = row.get(orig_col, "") if orig_col in row.index else ""
            cell    = ws.cell(row=r_idx, column=c_idx)

            if display in date_cols:
                parsed = normalize_dates(pd.Series([raw_val]))[0]
                if not pd.isna(parsed):
                    cell.value          = parsed.to_pydatetime()
                    cell.number_format  = FMT_DATE
                else:
                    cell.value = str(raw_val) if raw_val is not None else ""
            elif display in amount_cols:
                amount = normalize_amounts(pd.Series([raw_val]))[0]
                if amount is not None:
                    cell.value          = amount
                    cell.number_format  = FMT_MONEY
                else:
                    cell.value = str(raw_val) if raw_val is not None else ""
            else:
                cell.value = str(raw_val).strip() if raw_val is not None else ""

            cell.fill      = fill
            cell.font      = Font(name="Calibri", size=11)
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border    = _thin_border()

    ws.freeze_panes = "A2"
    _autofit(ws, headers)


# ─── Лист 4: Дубликаты ───────────────────────────────────────────────────────

def _sheet_duplicates(wb, dup_all: pd.DataFrame):
    ws = wb.create_sheet("Дубликаты")
    headers = ["Источник", "Номер ЭСФ", "Кол-во повторений", "Дата", "Сумма"]
    _write_header_row(ws, headers, row=1)

    if dup_all.empty:
        ws.cell(row=2, column=1, value="Дубликатов не обнаружено")
        _autofit(ws, headers)
        return

    fill = PatternFill("solid", fgColor=CLR_DUPLICATE)
    for r_idx, (_, row) in enumerate(dup_all.iterrows(), start=2):
        data = [
            row.get("Источник", ""),
            row.get("Номер ЭСФ", ""),
            row.get("Кол-во повторений", ""),
            row.get("Дата", ""),
            row.get("Сумма", ""),
        ]
        for c_idx, val in enumerate(data, start=1):
            cell       = ws.cell(row=r_idx, column=c_idx, value=str(val) if val is not None else "")
            cell.fill  = fill
            cell.font  = Font(name="Calibri", size=11)
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border    = _thin_border()

        # Форматируем «Кол-во повторений» как число
        cnt_cell = ws.cell(row=r_idx, column=3)
        try:
            cnt_cell.value         = int(row.get("Кол-во повторений", 0))
            cnt_cell.number_format = FMT_INT
        except (ValueError, TypeError):
            pass

    ws.freeze_panes = "A2"
    _autofit(ws, headers)


# ─── Лист 5: Расхождения ─────────────────────────────────────────────────────

def _sheet_mismatches(wb, mismatches: pd.DataFrame):
    ws = wb.create_sheet("Расхождения по реквизитам")
    headers = [
        "Номер ЭСФ",
        "Дата (1С)",
        "Дата (ИС ЭСФ)",
        "Реквизит",
        "Значение в 1С",
        "Значение в ИС ЭСФ",
    ]
    _write_header_row(ws, headers, row=1)

    if mismatches.empty:
        ws.cell(row=2, column=1, value="Расхождений не обнаружено")
        _autofit(ws, headers)
        return

    fill          = PatternFill("solid", fgColor=CLR_MISMATCH)
    bold_mismatch = Font(name="Calibri", size=11, bold=True)
    normal_font   = Font(name="Calibri", size=11)

    for r_idx, (_, row) in enumerate(mismatches.iterrows(), start=2):
        row_vals = [
            row.get("Номер ЭСФ", ""),
            row.get("Дата (1С)", ""),
            row.get("Дата (ИС ЭСФ)", ""),
            row.get("Реквизит", ""),
            row.get("Значение в 1С", ""),
            row.get("Значение в ИС ЭСФ", ""),
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell       = ws.cell(row=r_idx, column=c_idx,
                                 value=str(val) if val is not None else "")
            cell.fill  = fill
            # Жирный для расходящихся значений (колонки 5 и 6)
            cell.font  = bold_mismatch if c_idx in (5, 6) else normal_font
            cell.alignment = Alignment(horizontal="left", vertical="center",
                                       wrap_text=(c_idx in (5, 6)))
            cell.border    = _thin_border()

    ws.freeze_panes = "A2"
    _autofit(ws, headers)
    # Чуть пошире колонки с значениями
    ws.column_dimensions["E"].width = max(ws.column_dimensions["E"].width, 28)
    ws.column_dimensions["F"].width = max(ws.column_dimensions["F"].width, 28)


# ─── Утилиты форматирования ──────────────────────────────────────────────────

def _write_header_row(ws, headers: list, row: int = 1):
    """Записывает строку заголовков с тёмно-синим фоном и белым жирным текстом."""
    hdr_fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
    hdr_font = Font(name="Calibri", bold=True, size=11, color=CLR_HEADER_FG)
    for col_idx, title in enumerate(headers, start=1):
        cell           = ws.cell(row=row, column=col_idx, value=title)
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = _thin_border()
    ws.row_dimensions[row].height = 20


def _autofit(ws, headers: list):
    """Авторазмер колонок на основе заголовков и примерной ширины контента."""
    for i, header in enumerate(headers, start=1):
        col_letter = get_column_letter(i)
        max_len = len(str(header)) + 4
        # Смотрим первые 200 строк для оценки ширины
        for row in ws.iter_rows(min_row=2, max_row=min(ws.max_row, 201),
                                 min_col=i, max_col=i):
            for cell in row:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)) + 2)
        ws.column_dimensions[col_letter].width = min(max_len, 50)


def _thin_border() -> Border:
    thin = Side(style="thin", color="CCCCCC")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


# ════════════════════════════════════════════════════════════════════════════
# 6. ТОЧКА ВХОДА
# ════════════════════════════════════════════════════════════════════════════

def main():
    print()
    print(f"{Fore.CYAN}{'═'*60}")
    print(f"  ИНСТРУМЕНТ СВЕРКИ ЭСФ  |  АО КазАгроФинанс")
    print(f"{'═'*60}{Style.RESET_ALL}")
    print()

    # Загружаем конфиг
    cfg = load_config("config.yaml")

    # Загружаем файлы
    try:
        log_step("Загружаю файл 1С...")
        df_1c = load_file(cfg["file_1c"], cfg)
        log_ok(f"Файл 1С загружен: {len(df_1c):,} записей  [{cfg['file_1c']}]")
    except (FileNotFoundError, RuntimeError, ValueError) as exc:
        log_err(str(exc))
        sys.exit(1)

    try:
        log_step("Загружаю файл ИС «ЭСФ»...")
        df_esf = load_file(cfg["file_esf"], cfg)
        log_ok(f"Файл ИС ЭСФ загружен: {len(df_esf):,} записей  [{cfg['file_esf']}]")
    except (FileNotFoundError, RuntimeError, ValueError) as exc:
        log_err(str(exc))
        sys.exit(1)

    # Выполняем сверку
    log_step("Выполняется сверка...")
    try:
        results = compare_records(df_1c, df_esf, cfg)
    except ValueError as exc:
        log_err(str(exc))
        sys.exit(1)

    # Выводим краткую статистику
    only_1c_n    = len(results["only_1c"])
    only_esf_n   = len(results["only_esf"])
    dup_n        = len(results["dup_all"])
    mis_n        = results["mismatches"]["Номер ЭСФ"].nunique() \
                   if not results["mismatches"].empty else 0
    match_n      = len(results["matches"])

    print()
    log_info(f"Только в 1С:              {only_1c_n:>6,}")
    log_info(f"Только в ИС ЭСФ:          {only_esf_n:>6,}")
    log_info(f"Записи с дубликатами:     {dup_n:>6,}")
    if mis_n:
        log_warn(f"Найдено расхождений:      {mis_n:>6,}")
    else:
        log_ok(f"Найдено расхождений:      {mis_n:>6,}")
    log_ok(f"Полностью совпадают:      {match_n:>6,}")
    print()

    # Генерируем отчёт
    try:
        out_path = generate_report(results, cfg.get("output_dir", "output"))
        log_save(f"Отчёт сохранён: {out_path}")
    except Exception as exc:
        log_err(f"Ошибка при формировании отчёта: {exc}")
        import traceback; traceback.print_exc()
        sys.exit(1)

    print()
    print(f"{Fore.GREEN}{'═'*60}")
    print(f"  Сверка завершена успешно.")
    print(f"{'═'*60}{Style.RESET_ALL}")
    print()


if __name__ == "__main__":
    main()
