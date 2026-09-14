#!/usr/bin/env python3
"""Trace cell dependencies in an Excel financial model.

For every formula cell on the "Outputs" sheet, collects all cells
(transitively) whose values influence it, and writes the result as JSON.

Complex dynamic references (INDIRECT / OFFSET) are deliberately ignored:
the affected references are registered in the output metadata instead of
being resolved.
"""

import argparse
import json
import re
import sys
from collections import defaultdict
from datetime import datetime, timezone

from openpyxl import load_workbook

INDIRECT_RE = re.compile(r"(?i)\b(?:INDIRECT|OFFSET)\s*\(")
_STRING_RE = re.compile(r'"[^"]*"')

# Matches 'Sheet Name'!A1, Sheet!A1:B2, A1, $A$1, etc.
CELL_RE = re.compile(
    r"(?<![A-Za-z0-9_.])"
    r"((?:(?:'[^']*')|(?:[A-Za-z_][A-Za-z0-9_.]*))!)?"
    r"(\$?[A-Za-z]{1,3}\$?[0-9]{1,7})"
    r"(?::(\$?[A-Za-z]{1,3}\$?[0-9]{1,7}))?"
    r"(?![A-Za-z0-9_.])"
)
SHEET_NAME_RE = re.compile(r"^(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.]*))$")


def col_to_index(col_str):
    """Convert an Excel column label (without $) to a 1-based index."""
    idx = 0
    for ch in col_str:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx


class FormulaParser:
    def __init__(self, default_sheet):
        self.default_sheet = default_sheet

    def parse(self, formula):
        """Return (deps, dropped) for a formula string.

        deps: list of (sheet, col_idx, row_idx) 1-based references.
        dropped: list of dynamic-ref arguments that were removed.
        """
        text = formula[1:] if formula.startswith("=") else formula
        deps = []
        dropped = []
        for func in ("INDIRECT", "OFFSET"):
            text = self._strip_dynamic(text, func, dropped)
        text = _STRING_RE.sub("", text)
        for sheet_part, start, end in CELL_RE.findall(text):
            sheet = self._resolve_sheet(sheet_part)
            start_col, start_row = self._split_cell(start)
            if end:
                end_col, end_row = self._split_cell(end)
                for col in range(min(start_col, end_col), max(start_col, end_col) + 1):
                    for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
                        deps.append((sheet, col, row))
            else:
                deps.append((sheet, start_col, start_row))
        return deps, dropped

    @staticmethod
    def _in_string(text, pos):
        """True if position pos lies inside a "..."-literal."""
        return text.count('"', 0, pos) % 2 == 1

    def _strip_dynamic(self, text, func, dropped):
        """Remove all calls of one dynamic function, recording their arguments.

        Handles nested occurrences by stripping one call at a time and never
        touches occurrences that sit inside string literals.
        """
        pat = re.compile(r"(?i)\b" + re.escape(func) + r"\s*\(")
        pos = 0
        while True:
            m = pat.search(text, pos)
            if not m:
                return text
            if self._in_string(text, m.start()):
                pos = m.end()
                continue
            start = m.start()
            open_idx = m.end() - 1
            depth = 0
            end = None
            for i in range(open_idx, len(text)):
                if text[i] == "(":
                    depth += 1
                elif text[i] == ")":
                    depth -= 1
                    if depth == 0:
                        end = i + 1
                        break
            if end is None:
                return text
            arg = text[open_idx + 1:end - 1].strip('" \t')
            if arg:
                dropped.append((func, arg))
            text = text[:start] + text[end:]
            pos = start
        return text

    def _resolve_sheet(self, sheet_part):
        if not sheet_part:
            return self.default_sheet
        name = sheet_part[:-1]
        m = SHEET_NAME_RE.match(name)
        if not m:
            return name
        return m.group(1) or m.group(2)

    @staticmethod
    def _split_cell(ref):
        ref = ref.replace("$", "")
        m = re.match(r"^([A-Za-z]+)([0-9]+)$", ref)
        return col_to_index(m.group(1)), int(m.group(2))


def cell_id(sheet, col, row):
    """Build a readable 'SheetName!A1' identifier."""
    return f"{sheet}!{to_col(col)}{row}"


def to_col(idx):
    chars = []
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        chars.append(chr(ord("A") + rem))
    return "".join(reversed(chars))


class DependencyGraph:
    def __init__(self, workbook_path):
        self.path = workbook_path
        self.wb = load_workbook(workbook_path, data_only=False)
        self.sheet_titles = {ws.title for ws in self.wb.worksheets}
        self.direct = {}  # (sheet, col, row) -> set of (sheet, col, row)
        self.formula_text = {}  # (sheet, col, row) -> original formula
        self.dropped_refs = []  # (cell, func, arg)
        self._index()

    def _index(self):
        for ws in self.wb.worksheets:
            parser = FormulaParser(default_sheet=ws.title)
            for row in ws.iter_rows():
                for cell in row:
                    if not isinstance(cell.value, str) or not cell.value.startswith("="):
                        continue
                    key = (ws.title, cell.column, cell.row)
                    deps, dropped = parser.parse(cell.value)
                    for func, arg in dropped:
                        self.dropped_refs.append(
                            {"cell": cell_id(*key), "function": func, "dropped_arg": arg}
                        )
                    self.formula_text[key] = cell.value
                    self.direct[key] = {(sheet, col, r) for sheet, col, r in deps}

    @staticmethod
    def _in_bounds(cell):
        sheet, col, row = cell
        return 1 <= col <= 16384 and 1 <= row <= 1048576

    def trace(self, start_sheet, col, row):
        """Return all cells transitively influencing (start_sheet, col, row)."""
        start = (start_sheet, col, row)
        if start not in self.direct:
            return set()
        seen = set()
        stack = list(self.direct[start])
        while stack:
            cell = stack.pop()
            if cell in seen:
                continue
            if not self._in_bounds(cell):
                seen.add(cell)
                continue
            seen.add(cell)
            if cell in self.direct:
                stack.extend(self.direct[cell])
        seen.discard(start)
        return seen

    def trace_outputs(self, output_sheet):
        """Trace every formula cell on the output sheet."""
        results = {}
        ws = self.wb[output_sheet]
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or not cell.value.startswith("="):
                    continue
                deps = self.trace(output_sheet, cell.column, cell.row)
                if deps:
                    results[cell_id(output_sheet, cell.column, cell.row)] = sorted(
                        cell_id(s, c, r) for s, c, r in deps
                    )
        return results

    def count_output_cells(self, output_sheet):
        ws = self.wb[output_sheet]
        return sum(
            1
            for row in ws.iter_rows()
            for cell in row
            if isinstance(cell.value, str) and cell.value.startswith("=")
        )


def main(argv=None):
    parser = argparse.ArgumentParser(description="Trace dependencies for Outputs cells in an Excel model.")
    parser.add_argument("input", nargs="?", default="data/model.xlsx", help="Path to the Excel workbook")
    parser.add_argument("-o", "--out", default="dependencies.json", help="Output JSON file")
    parser.add_argument("--output-sheet", default="Outputs", help="Sheet to trace (default: Outputs)")
    args = parser.parse_args(argv)

    graph = DependencyGraph(args.input)
    results = graph.trace_outputs(args.output_sheet)

    payload = {
        "meta": {
            "source": args.input,
            "output_sheet": args.output_sheet,
            "generated": datetime.now(timezone.utc).isoformat(),
            "output_cells_with_formulas": graph.count_output_cells(args.output_sheet),
            "output_cells_with_dependencies": len(results),
            "ignored_dynamic_refs": graph.dropped_refs,
        },
        "dependencies": dict(sorted(results.items())),
    }

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(f"Cells traced: {len(results)}")
    print(f"Dynamic refs ignored: {len(graph.dropped_refs)}")
    print(f"Written to: {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())