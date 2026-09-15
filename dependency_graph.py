#!/usr/bin/env python3
"""Trace cell dependencies in an Excel financial model.

For every formula cell on the "Outputs" sheet, collects all cells
(transitively) whose values influence it, and writes the result as JSON.

Complex dynamic references (INDIRECT / OFFSET) are deliberately ignored:
the affected references are registered in the output metadata instead of
being resolved.

Memory-friendly: the workbook is read in streaming (read-only) mode and
the result is streamed to disk cell-by-cell instead of being accumulated
in memory.
"""

import argparse
import gc
import json
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import load_workbook

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
    def __init__(self, workbook_path, output_sheet="Outputs"):
        self.output_sheet = output_sheet
        # read_only streams the workbook from disk: drastically less memory.
        self.wb = load_workbook(workbook_path, data_only=False, read_only=True)
        self.direct = {}  # (sheet, col, row) -> set of (sheet, col, row)
        self.dropped_refs = []  # {cell, function, dropped_arg}
        self.output_cells = []  # formula cells on the output sheet
        self._index()

    def _index(self):
        for ws in self.wb.worksheets:
            parser = FormulaParser(default_sheet=ws.title)
            is_output = ws.title == self.output_sheet
            for row in ws.iter_rows():
                for cell in row:
                    if not isinstance(cell.value, str) or not cell.value.startswith("="):
                        continue
                    key = (ws.title, cell.column, cell.row)
                    if is_output:
                        self.output_cells.append(key)
                    deps, dropped = parser.parse(cell.value)
                    for func, arg in dropped:
                        self.dropped_refs.append(
                            {"cell": cell_id(*key), "function": func, "dropped_arg": arg}
                        )
                    if deps:
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


def iter_dependency_entries(graph):
    """Yield (cell_id, sorted dep cell ids) for every output cell with deps.

    Output cells are visited in a deterministic (column, row) order and each
    cell's result is produced and discarded on the fly.
    """
    for sheet, col, row in sorted(graph.output_cells):
        deps = graph.trace(sheet, col, row)
        if deps:
            yield cell_id(sheet, col, row), sorted(cell_id(*c) for c in deps)


def count_dependency_entries(graph):
    """Transitive count of output cells that have at least one dependency."""
    return sum(1 for _ in iter_dependency_entries(graph))


def workbook_meta(source_path, output_sheet, generated, formula_count, entry_count, dropped_refs):
    return {
        "source": source_path,
        "output_sheet": output_sheet,
        "generated": generated,
        "output_cells_with_formulas": formula_count,
        "output_cells_with_dependencies": entry_count,
        "ignored_dynamic_refs": dropped_refs,
    }


def _write_entries(f, entries, indent):
    first = True
    for cell_id_str, deps in entries:
        if not first:
            f.write(",\n")
        f.write(indent)
        json.dump(cell_id_str, f, ensure_ascii=False)
        f.write(": ")
        json.dump(deps, f, ensure_ascii=False)
        first = False


def write_output_file(path, meta, entries):
    """Write a single-workbook result: {meta, dependencies} streamed to disk."""
    with open(path, "w", encoding="utf-8") as f:
        f.write('{\n  "meta": ')
        json.dump(meta, f, ensure_ascii=False, indent=2)
        f.write(',\n  "dependencies": {')
        _write_entries(f, entries, indent="    ")
        f.write('\n  }\n}\n')


def process_single(file_path, out_path, output_sheet):
    graph = DependencyGraph(file_path, output_sheet)
    try:
        entry_count = count_dependency_entries(graph)
        meta = workbook_meta(
            file_path, output_sheet, datetime.now(timezone.utc).isoformat(),
            len(graph.output_cells), entry_count, graph.dropped_refs,
        )
        write_output_file(out_path, meta, iter_dependency_entries(graph))
    finally:
        graph.wb.close()
        del graph
    return meta


def process_directory(input_dir, output_dir, output_sheet, merge=False):
    """Process all *.xlsx files in *input_dir*.

    Returns 0 on success (even if some files failed), -1 on fatal error.
    With --merge the merged file is streamed workbook-by-workbook, so memory
    stays bounded regardless of how many/ how large the workbooks are.
    """
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    xlsx_files = sorted(input_path.glob("*.xlsx"))
    if not xlsx_files:
        print(f"No *.xlsx files found in '{input_dir}'.")
        return 0

    batch_t0 = time.time()
    total = len(xlsx_files)
    successful = 0
    failed = 0

    merged_f = None
    first_workbook = True
    if merge:
        merged_f = open(output_path / "merged_dependencies.json", "w", encoding="utf-8")
        merged_f.write('{\n  "workbooks": {')

    try:
        for idx, xlsx_file in enumerate(xlsx_files, 1):
            t0 = time.time()
            print(f"[{idx}/{total}] Processing: {xlsx_file.name}")
            try:
                graph = DependencyGraph(str(xlsx_file), output_sheet)
                try:
                    entry_count = count_dependency_entries(graph)
                    meta = workbook_meta(
                        str(xlsx_file), output_sheet, datetime.now(timezone.utc).isoformat(),
                        len(graph.output_cells), entry_count, graph.dropped_refs,
                    )
                    entries = iter_dependency_entries(graph)
                    if merge:
                        if not first_workbook:
                            merged_f.write(",\n")
                        first_workbook = False
                        merged_f.write("    ")
                        json.dump(xlsx_file.name, merged_f, ensure_ascii=False)
                        merged_f.write(': {\n      "meta": ')
                        json.dump(meta, merged_f, ensure_ascii=False, separators=(",", ":"))
                        merged_f.write(',\n      "dependencies": {')
                        _write_entries(merged_f, entries, indent="        ")
                        merged_f.write("\n      }\n    }")
                    else:
                        out_file = output_path / (xlsx_file.stem + "_dependencies.json")
                        write_output_file(out_file, meta, entries)
                finally:
                    graph.wb.close()
                    del graph
            except Exception as exc:
                elapsed = time.time() - t0
                print(f"  ERROR: {xlsx_file.name} — {exc} ({elapsed:.2f}s)")
                failed += 1
                continue
            finally:
                gc.collect()

            elapsed = time.time() - t0
            successful += 1
            cells = meta["output_cells_with_dependencies"]
            dropped = len(meta["ignored_dynamic_refs"])
            print(f"  → {cells} cells traced, {dropped} dynamic refs ignored ({elapsed:.2f}s)")
    finally:
        if merged_f is not None:
            merged_meta = {
                "source_dir": str(input_path),
                "output_sheet": output_sheet,
                "generated": datetime.now(timezone.utc).isoformat(),
                "total_workbooks": total,
                "successful": successful,
                "failed": failed,
            }
            merged_f.write('\n  },\n  "meta": ')
            json.dump(merged_meta, merged_f, ensure_ascii=False, indent=2)
            merged_f.write("\n}\n")
            merged_f.close()

    total_elapsed = time.time() - batch_t0
    print(f"\nDone. Successful: {successful}, Failed: {failed}, Total: {total} ({total_elapsed:.2f}s)")
    return 0


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Trace dependencies for Outputs cells in an Excel model."
    )
    parser.add_argument(
        "input", nargs="?", default="data/model.xlsx",
        help="Path to a single Excel workbook (use --input-dir for batch mode)"
    )
    parser.add_argument("-o", "--out", default="dependencies.json", help="Output JSON file (single-file mode)")
    parser.add_argument("--output-sheet", default="Outputs", help="Sheet to trace (default: Outputs)")

    # Batch-mode arguments
    parser.add_argument(
        "--input-dir", default=None,
        help="Directory containing *.xlsx files to process (batch mode)"
    )
    parser.add_argument(
        "--output-dir", default="results",
        help="Directory for output JSON files (default: results/)"
    )
    parser.add_argument(
        "--merge", action="store_true", default=False,
        help="Merge all results into a single merged_dependencies.json"
    )

    args = parser.parse_args(argv)

    # Batch mode
    input_dir = args.input_dir or (args.input if args.input and Path(args.input).is_dir() else None)
    if input_dir is not None and Path(input_dir).is_dir():
        return process_directory(input_dir, args.output_dir, args.output_sheet, args.merge)

    # Single-file mode
    meta = process_single(args.input, args.out, args.output_sheet)
    print(f"Cells traced: {meta['output_cells_with_dependencies']}")
    print(f"Dynamic refs ignored: {len(meta['ignored_dynamic_refs'])}")
    print(f"Written to: {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())