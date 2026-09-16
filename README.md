# Excel Dependency Graph

Скрипт строит граф зависимостей для финансовой Excel-модели: для каждой ячейки на листе **Outputs** возвращает все ячейки, значения которых на неё влияют (транзитивно — включая промежуточные формульные ячейки Calc, P&L, CFS, PRINT и др.).

## Требования

- Python 3.8+
- Сторонних зависимостей нет: только стандартная библиотека (`zipfile`, `xml.etree.ElementTree`, `re`).

## Запуск

```bash
python3 dependency_graph.py [INPUT.xlsx] [-o OUTPUT.json] [--output-sheet SHEET]
```

| Параметр | По умолчанию | Описание |
|---|---|---|
| `INPUT.xlsx` | `data/model.xlsx` | Путь к файлу модели или директории |
| `-o, --out` | `dependencies.json` | Файл результата (JSON), только для одиночного файла |
| `--output-sheet` | `Outputs` | Лист, для ячеек которого строится граф |
| `--input-dir` | — | Директория с `*.xlsx` для пакетной обработки |
| `--output-dir` | `results/` | Папка для результатов (пакетный режим) |
| `--merge` | — | Собрать все результаты в один `merged_dependencies.json` |
| `--no-save` | — | Только расчёт: построить граф, вывести сводку, ничего не записывать |
| `--timeout N` | `0` | Пропускать файл, если построение графа заняло больше N секунд (0 = без лимита) |
| `--transpose` | — | Развернуть таблицу пропущенных функций (файлы в строках) |
| `--json` | — | Писать JSON (по умолчанию, если не указан ни `--json`, ни `--csv`) |
| `--csv` | — | Писать CSV пар `Inputs → Outputs` |
| `--csv-out` | — | Путь CSV в одиночном режиме (по умолчанию `<out>.csv`) |
| `--input-sheet` | `Inputs` | Лист входных ячеек для CSV |
| `--input-name-col` | авто | Колонка имён входов (по умолчанию ищется заголовок «Наименование») |
| `--output-name-col` | авто | Колонка имён выходов (по умолчанию ищется заголовок «Наименование») |
| `--years FROM:TO` | — | Ограничить годы входов (`FROM:TO`, `FROM:`, `:TO` или `YEAR`) |
| `--from YEAR` / `--to YEAR` | — | Границы годов входов |

## Запуск (одиночный файл)

```bash
python3 dependency_graph.py data/model.xlsx -o deps.json
```

## Пакетная обработка

```bash
# Каждый .xlsx → отдельный JSON в results/
python3 dependency_graph.py --input-dir data/ --output-dir results/

# Один общий файл со всеми результатами
python3 dependency_graph.py --input-dir data/ --merge

# Папку можно передать как позиционный аргумент
python3 dependency_graph.py data/
```

## Только расчёт (без записи)

Режим `--no-save` строит граф и замыкание зависимостей, печатает сводку
(число формул, выходов, пар «формула → ячейка», отброшенных динамических ссылок)
и **не пишет ничего на диск**. Подходит для быстрых проверок модели без генерации
гигабайтных JSON.

```bash
# Одиночный файл
python3 dependency_graph.py --no-save data/model.xlsx

# Вся папка
python3 dependency_graph.py --no-save --input-dir data/
```

Пример вывода (`--no-save --input-dir data/`):
```
[1/2] Processing: 2poj431skazt.xlsx
  → 2poj431skazt.xlsx: 1076 output formulas, 1076 with deps, 59560806 dep pairs, 3892 dynamic refs ignored
  ✓ computed (8.91s)
[2/2] Processing: model.xlsx
  → model.xlsx: 1479 output formulas, 1479 with deps, 680695 dep pairs, 23 dynamic refs ignored
  ✓ computed (0.80s)

Done. Successful: 2, Failed: 0, Total: 2 (9.70s)
```

Пример вывода пакетного режима с записью:
```
[1/2] Processing: 2poj431skazt.xlsx
  → 1076 cells traced, 3892 dynamic refs ignored (15.04s)
[2/2] Processing: model.xlsx
  → 1479 cells traced, 23 dynamic refs ignored (0.90s)

Done. Successful: 2, Failed: 0, Total: 2 (15.94s)
```

## Формат результата

```json
{
  "meta": {
    "source": "data/model.xlsx",
    "output_sheet": "Outputs",
    "output_cells_with_formulas": 1490,
    "output_cells_with_dependencies": 1479,
    "ignored_dynamic_refs": [
      {"cell": "Outputs!H18", "function": "INDIRECT", "dropped_arg": "Inputs!AA34"}
    ]
  },
  "dependencies": {
    "Outputs!H2": ["Calc!L78", "P&L!L8"]
  }
}
```

## CSV-выход (пары Inputs → Outputs)

Флаг `--csv` включает CSV, совместимый по ключу с `xl_stat/sensitivity2.py`:
одна строка на пару «вход → выход», только транзитивные предки с листа
`--input-sheet`, у которых есть имя в колонке имён и год в строке 1.
Колонки (`utf-8-sig`, открывается в Excel):

```
input_cell,input_name,input_year,output_cell,output_name,output_year
AF4,"Инфляция - Рост индекса потребительских цен в США...",2023,M3,Операционные расходы (без амортизации),2023
```

```bash
# одиночный файл → model.csv
python3 dependency_graph.py data/model.xlsx --csv

# батч → results/per_file/<stem>.csv
python3 dependency_graph.py --input-dir data/ --output-dir results/ --csv

# батч + слияние → results/inputs_outputs.csv
python3 dependency_graph.py --input-dir data/ --output-dir results/ --csv --merge

# и JSON, и CSV
python3 dependency_graph.py data/model.xlsx --json --csv
```

`--csv` без `--json` пишет только CSV; без обоих флагов по умолчанию пишется JSON.
Колонки имён определяются динамически по заголовку «Наименование» в строке 1
(переопределяются `--input-name-col` / `--output-name-col`). `--years/--from/--to`
ограничивают годы входов, как в `sensitivity2.py`.

## Ограничения

- Ссылки на диапазоны (например `SUM(J2:AY2)`) раскрываются в отдельные ячейки.
- Ветвления `CHOOSE`/`IF` не разрешаются: включаются все аргументы-кандидаты.
- Динамические ссылки `INDIRECT`/`OFFSET` игнорируются; пропущенные аргументы фиксируются в `meta.ignored_dynamic_refs`.
- Внешние ссылки на другие книги не поддерживаются.
