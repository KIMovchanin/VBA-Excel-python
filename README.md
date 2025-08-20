# Excel Report Migrator

A small educational project for migrating logic from VBA to Python:

* Reads an Excel file, merges `raw` + `dict`, calculates revenue, and builds a report.
* Configurable via a YAML config.
* Has both GUI (tkinter) and non-GUI (CLI) versions.
* Covered with minimal pytest tests.

## 📂 Structure

excel\_learning/
.venv/ # (optional) virtual environment
.gitignore
README.md
task\_with\_test/
config.yml # project config (sheet/column names and output fields)
report\_with\_gui.py # GUI version (tkinter)
report\_without\_gui.py # CLI version
sales\_example.xlsx # example input data
tests/
init.py
test\_report.py # pytest tests for the "pure" build\_report logic

## 🧩 Dependencies

* Python 3.10+
* `pandas`, `openpyxl`, `pyyaml`, `pytest`
* `tkinter` — included in the standard Python for Windows/macOS (on Linux may require `sudo apt install python3-tk`)

Recommended way to install dependencies:

```bash
pip install pandas openpyxl pyyaml pytest
```

## ▶️ Run (CLI, without GUI)

From the project root or from the task\_with\_test folder:

```bash
# from root:
python task_with_test/report_without_gui.py --input task_with_test/sales_example.xlsx --out task_with_test/report.xlsx --config task_with_test/config.yml

# or from task_with_test folder:
cd task_with_test
python report_without_gui.py --input sales_example.xlsx --out report.xlsx --config config.yml
```

Parameters:

\--input — path to the input .xlsx (must contain sheets raw and dict, names can be changed in config)

\--out — path to the output .xlsx (default: report.xlsx)

\--config — path to config.yml (default: config.yml next to the script)

## 🖱️ Run (GUI)

GUI is a small tkinter window:

```bash
# from root:
python task_with_test/report_with_gui.py --config task_with_test/config.yml

# or from task_with_test folder:
cd task_with_test
python report_with_gui.py --config config.yml
```

What happens:

Click “Choose…” — select the input .xlsx.

Click “Save as…” — choose where to save report.xlsx.

Click “Generate report” — on success, you’ll see the notification “Report saved…”.

If --config is not provided, GUI will try to find config.yml next to the script.

## 🐍 Tests

Run from the project root:

```bash
python -m pytest task_with_test/tests -q
```

Expected: all tests green (3 passed).

Tests check:

* correct revenue calculation and price\_override handling,
* correct type conversion (dates/numbers),
* merge protection with validate="many\_to\_one" (duplicates in dictionary should raise an error).

## 📝 What’s inside the report

Sum by days and categories:

effective\_price = price\_override (if exists) else price

each\_sum = effective\_price \* qty

revenue = Σ each\_sum by group (category, date)

---

# Excel Report Migrator

Небольшой учебный проект для миграции логики из VBA в Python:
- Читает Excel-файл, сводит `raw` + `dict`, считает выручку, формирует отчёт.
- Настраивается через YAML-конфиг.
- Есть версия с GUI (tkinter) и без GUI (CLI).
- Покрыт минимальными pytest-тестами.

## 📂 Структура
excel_learning/
.venv/ # (опционально) виртуальное окружение
.gitignore
README.md
task_with_test/
config.yml # конфиг проекта (имена листов/колонок и выходные поля)
report_with_gui.py # версия с GUI (tkinter)
report_without_gui.py # версия без GUI (CLI)
sales_example.xlsx # пример входных данных
tests/
init.py
test_report.py # pytest-тесты для "чистой" логики build_report


## 🧩 Зависимости

- Python 3.10+  
- `pandas`, `openpyxl`, `pyyaml`, `pytest`  
- `tkinter` — идёт в составе стандартного Python для Windows/macOS (на Linux может понадобиться `sudo apt install python3-tk`)

Рекомендуемый способ поставить зависимости:
```bash
pip install pandas openpyxl pyyaml pytest
```

## ▶️ Запуск (CLI, без GUI)

Из корня проекта или из папки task_with_test:
```bash
# из корня:
python task_with_test/report_without_gui.py --input task_with_test/sales_example.xlsx --out task_with_test/report.xlsx --config task_with_test/config.yml

# или из папки task_with_test:
cd task_with_test
python report_without_gui.py --input sales_example.xlsx --out report.xlsx --config config.yml
```
Параметры:

--input — путь к входному .xlsx (должны быть листы raw и dict, имена можно менять в конфиге)

--out — путь к выходному .xlsx (по умолчанию report.xlsx)

--config — путь к config.yml (по умолчанию config.yml рядом со скриптом)

## 🖱️ Запуск (GUI)
GUI — это маленькое окно на tkinter:
```bash
# из корня:
python task_with_test/report_with_gui.py --config task_with_test/config.yml

# или из папки task_with_test:
cd task_with_test
python report_with_gui.py --config config.yml
```

Что происходит:

Нажми «Выбрать…» — укажи входной .xlsx.

Нажми «Сохранить как…» — укажи, куда сохранить report.xlsx.

Жми «Сформировать отчёт» — по успеху увидишь уведомление «Отчёт сохранён…».

Если не передавать --config, GUI попытается найти config.yml рядом со скриптом.

## 🐍Тесты

Запуск из корня:
```bash
python -m pytest task_with_test/tests -q
```

Ожидаемо: все тесты зелёные (3 passed).

Тесты проверяют:

корректный расчёт выручки и работу price_override,

корректное приведение типов (даты/числа),

защиту merge с validate="many_to_one" (дубликаты в справочнике должны падать).

## 📝 Что внутри отчёта

Сумма по дням и категориям:

effective_price = price_override (если есть) иначе price

each_sum = effective_price * qty

revenue = Σ each_sum по группе (category, date)