# Excel Dependency Graph

Скрипт строит граф зависимостей для финансовой Excel-модели: для каждой ячейки на листе **Outputs** возвращает все ячейки, значения которых на неё влияют (транзитивно — включая промежуточные формульные ячейки Calc, P&L, CFS, PRINT и др.).

## Требования

- Python 3.8+
- `openpyxl` (`pip install openpyxl`)

## Запуск

```bash
python3 dependency_graph.py [INPUT.xlsx] [-o OUTPUT.json] [--output-sheet SHEET]
```

| Параметр | По умолчанию | Описание |
|---|---|---|
| `INPUT.xlsx` | `data/model.xlsx` | Путь к файлу модели |
| `-o, --out` | `dependencies.json` | Файл результата (JSON) |
| `--output-sheet` | `Outputs` | Лист, для ячеек которого строится граф |

## Пример

```bash
python3 dependency_graph.py data/model.xlsx -o deps.json
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