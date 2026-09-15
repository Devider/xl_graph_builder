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

## Ограничения

- Ссылки на диапазоны (например `SUM(J2:AY2)`) раскрываются в отдельные ячейки.
- Ветвления `CHOOSE`/`IF` не разрешаются: включаются все аргументы-кандидаты.
- Динамические ссылки `INDIRECT`/`OFFSET` игнорируются; пропущенные аргументы фиксируются в `meta.ignored_dynamic_refs`.
- Внешние ссылки на другие книги не поддерживаются.
