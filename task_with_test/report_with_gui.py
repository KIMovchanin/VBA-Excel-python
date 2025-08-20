from __future__ import annotations

import argparse
import logging
from typing import Dict

import pandas as pd
import yaml

# --- GUI ---
import tkinter as tk
from tkinter import filedialog, messagebox

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)


# ---------- КОНФИГ ----------
def load_config(path: str = "config.yml") -> Dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def _require_columns(df: pd.DataFrame, needed: list[str], where: str) -> None:
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise KeyError(
            f"В листе '{where}' отсутствуют колонки: {missing}. Есть: {list(df.columns)}"
        )


# ---------- ЧИСТАЯ ЛОГИКА ----------
def build_report(df_raw: pd.DataFrame, df_dict: pd.DataFrame, config: Dict) -> pd.DataFrame:
    C = config["columns"]
    out_rev_col = config["output"]["revenue_column"]

    _require_columns(df_raw, [C["date"], C["code"], C["qty"], C["price_override"]], where="raw")
    _require_columns(df_dict, [C["code"], C["price"], C["category"]], where="dict")

    df_raw = df_raw.copy()
    df_dict = df_dict.copy()

    df_raw[C["date"]] = pd.to_datetime(df_raw[C["date"]], dayfirst=True, errors="coerce")
    df_raw[C["qty"]] = pd.to_numeric(df_raw[C["qty"]], errors="coerce")
    df_raw[C["price_override"]] = pd.to_numeric(df_raw[C["price_override"]], errors="coerce")
    df_dict[C["price"]] = pd.to_numeric(df_dict[C["price"]], errors="coerce")

    # если category уже есть в raw, удаляем, чтобы не получить category_x/category_y
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
    log.info(f"Пустых цен после merge: {df_merge[C['price']].isna().sum()}")

    df_merge["effective_price"] = df_merge[C["price_override"]].fillna(df_merge[C["price"]])
    df_merge["each_sum"] = df_merge["effective_price"] * df_merge[C["qty"]]

    out = (
        df_merge.groupby([C["category"], C["date"]], dropna=False)["each_sum"]
        .sum()
        .reset_index(name=out_rev_col)
    )
    return out


# ---------- IO-ОБЁРТКА ----------
def create_report(file_name: str, out_path: str, config: Dict) -> pd.DataFrame | None:
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


# ---------- GUI ----------
def run_gui(config: Dict) -> None:
    root = tk.Tk()
    root.title("Excel Report Migrator")
    root.geometry("420x220")

    # поля
    tk.Label(root, text="Входной файл:").pack(anchor="w", padx=10, pady=(10, 0))
    entry_input = tk.Entry(root, width=50)
    entry_input.pack(padx=10)

    def select_file():
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            entry_input.delete(0, tk.END)
            entry_input.insert(0, path)
            # автоподстановка имени для отчёта
            if not entry_output.get():
                if path.lower().endswith(".xlsx"):
                    entry_output.insert(0, path[:-5] + "_report.xlsx")

    tk.Button(root, text="Выбрать…", command=select_file).pack(padx=10, pady=4)

    tk.Label(root, text="Выходной файл:").pack(anchor="w", padx=10, pady=(10, 0))
    entry_output = tk.Entry(root, width=50)
    entry_output.pack(padx=10)

    def save_file():
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="report.xlsx",
        )
        if path:
            entry_output.delete(0, tk.END)
            entry_output.insert(0, path)

    tk.Button(root, text="Сохранить как…", command=save_file).pack(padx=10, pady=4)

    def run_report():
        inp = entry_input.get().strip()
        outp = entry_output.get().strip()
        if not inp or not outp:
            messagebox.showerror("Ошибка", "Укажите входной и выходной файл")
            return
        try:
            res = create_report(inp, outp, config)
            if res is None:
                messagebox.showerror("Ошибка", "Не удалось сформировать отчёт. См. логи.")
            else:
                messagebox.showinfo("Готово", f"Отчёт сохранён:\n{outp}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    tk.Button(root, text="Сформировать отчёт", command=run_report).pack(pady=12)
    root.mainloop()


# ---------- CLI ----------
def main() -> None:
    parser = argparse.ArgumentParser(description="Генерация отчёта из Excel по конфигу")
    parser.add_argument("--input", help="Входной .xlsx файл (для CLI)")
    parser.add_argument("--out", default="report.xlsx", help="Выходной .xlsx (для CLI)")
    parser.add_argument("--config", default="config.yml", help="Путь к YAML‑конфигу")
    parser.add_argument("--gui", action="store_true", help="Запустить графический интерфейс")
    args = parser.parse_args()

    cfg = load_config(args.config)

    if args.gui:
        run_gui(cfg)
    else:
        if not args.input:
            parser.error("--input обязателен в режиме CLI (или используйте --gui)")
        create_report(args.input, args.out, cfg)


if __name__ == "__main__":
    main()
