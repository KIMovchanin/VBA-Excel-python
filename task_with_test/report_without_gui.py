from __future__ import annotations

import argparse
import logging
from typing import Dict

import pandas as pd
import yaml


# ---------- ЛОГИ ----------

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)


# ---------- КОНФИГ ----------

def load_config(path: str = "config.yml") -> Dict:
    """Загружает YAML‑конфиг."""
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def _require_columns(df: pd.DataFrame, needed: list[str], where: str) -> None:
    """Проверяет наличие обязательных колонок; бросает KeyError при отсутствии."""
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise KeyError(f"В листе '{where}' отсутствуют колонки: {missing}. Есть: {list(df.columns)}")


# ---------- ЧИСТАЯ ЛОГИКА (тестируется напрямую) ----------

def build_report(df_raw: pd.DataFrame, df_dict: pd.DataFrame, config: Dict) -> pd.DataFrame:
    C = config["columns"]
    out_rev_col = config["output"]["revenue_column"]

    # Проверим обязательные поля
    _require_columns(df_raw, [C["date"], C["code"], C["qty"], C["price_override"]], where="raw")
    _require_columns(df_dict, [C["code"], C["price"], C["category"]], where="dict")

    # Типизация и приведение
    df_raw = df_raw.copy()
    df_dict = df_dict.copy()

    df_raw[C["date"]] = pd.to_datetime(df_raw[C["date"]], dayfirst=True, errors="coerce")
    df_raw[C["qty"]] = pd.to_numeric(df_raw[C["qty"]], errors="coerce")
    df_raw[C["price_override"]] = pd.to_numeric(df_raw[C["price_override"]], errors="coerce")
    df_dict[C["price"]] = pd.to_numeric(df_dict[C["price"]], errors="coerce")

    if C["category"] in df_raw.columns:
        df_raw = df_raw.drop(columns=[C["category"]])

    df_merge = pd.merge(
        df_raw,
        df_dict[[C["code"], C["price"], C["category"]]],
        on=C["code"],
        how="left",
        validate="many_to_one",
    )

    log.info(f"После merge: {df_merge.shape}")
    n_price_na = df_merge[C["price"]].isna().sum()
    log.info(f"Пустых цен после merge: {n_price_na}")

    # Бизнес‑логика цены и выручки
    df_merge["effective_price"] = df_merge[C["price_override"]].fillna(df_merge[C["price"]])
    df_merge["each_sum"] = df_merge["effective_price"] * df_merge[C["qty"]]

    out = (
        df_merge
        .groupby([C["category"], C["date"]], dropna=False)["each_sum"]
        .sum()
        .reset_index(name=out_rev_col)
    )

    return out


# ---------- IO-ОБЁРТКА (чтение/запись Excel) ----------

def create_report(file_name: str, out_path: str, config: Dict) -> pd.DataFrame | None:
    """Читает Excel, строит отчёт через build_report и сохраняет в Excel."""
    if not file_name.lower().endswith(".xlsx"):
        log.error("Неподходящее расширение файла. Нужно '.xlsx'.")
        return None

    try:
        excel_file = pd.ExcelFile(file_name)
    except FileNotFoundError:
        log.error("Файл не найден: %s", file_name)
        return None
    except Exception as e:
        log.error("Не удалось открыть Excel: %s", e)
        return None

    raw_sheet = config["sheets"]["raw"]
    dict_sheet = config["sheets"]["dict"]

    try:
        df_raw = excel_file.parse(sheet_name=raw_sheet)
        df_dict = excel_file.parse(sheet_name=dict_sheet)
    except ValueError as e:
        log.error("Ошибка чтения листов: %s", e)
        return None

    log.info(f"Прочитан лист {raw_sheet}: {df_raw.shape[0]} строк, {df_raw.shape[1]} столбцов")

    result = build_report(df_raw, df_dict, config)

    log.info(f"Готовим отчёт: {result.shape[0]} строк, {result.shape[1]} столбцов")

    out_sheet = config["output"]["sheet_name"]
    try:
        result.to_excel(out_path, sheet_name=out_sheet, index=False)
    except Exception as e:
        log.error("Ошибка записи '%s': %s", out_path, e)
        return None

    log.info("Отчёт сохранён: %s (лист: %s)", out_path, out_sheet)
    return result


# ---------- CLI ----------

def main() -> None:
    parser = argparse.ArgumentParser(description="Генерация отчёта из Excel по конфигу")
    parser.add_argument("--input", required=True, help="Входной .xlsx файл с листами raw/dict (или как в конфиге)")
    parser.add_argument("--out", default="report.xlsx", help="Выходной .xlsx")
    parser.add_argument("--config", default="config.yml", help="Путь к YAML‑конфигу")
    args = parser.parse_args()

    cfg = load_config(args.config)
    create_report(args.input, args.out, cfg)


if __name__ == "__main__":
    main()
