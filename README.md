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