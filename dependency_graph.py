#!/usr/bin/env python3
"""Trace cell dependencies in an Excel financial model.

For every formula cell on the "Outputs" sheet, collects all cells
(transitively) whose values influence it, and writes the result as JSON.

Highlights:
- No third-party dependencies: the .xlsx is read straight from its XML
  parts (zipfile + ElementTree), which is both fast and memory-friendly.
- Shared formulas (Excel writes large contiguous ranges as a single
  master formula) are resolved and their relative references are
  translated cell-by-cell, like Excel does.
- Complex dynamic references (INDIRECT / OFFSET) are deliberately
  ignored: the affected references are registered in the output metadata
  instead of being resolved.
- Transitive closure is computed via a single bitset-propagation pass
  over an SCC-condensed dependency DAG (Tarjan + Kahn + int-bitsets),
  giving O(V+E) performance regardless of output count.
"""

import argparse
import gc
import json
import re
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

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
COORD_RE = re.compile(r"^([A-Za-z]+)([0-9]+)$")
REF_RE = re.compile(r"^(\$?)([A-Za-z]+)(\$?)([0-9]+)$")


def col_to_index(col_str):
    """Convert an Excel column label (without $) to a 1-based index."""
    idx = 0
    for ch in col_str:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx


def to_col(idx):
    chars = []
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        chars.append(chr(ord("A") + rem))
    return "".join(reversed(chars))


def split_coord(coord):
    """'AB12' -> (col_index, row)  both 1-based."""
    m = COORD_RE.match(coord)
    return col_to_index(m.group(1)), int(m.group(2))


def shift_ref(ref, dc, dr):
    """Shift a cell reference by (dc, dr); $-anchored parts stay fixed."""
    m = REF_RE.match(ref)
    dollar_col, col, dollar_row, row = m.groups()
    c = col_to_index(col)
    r = int(row)
    if not dollar_col:
        c += dc
    if not dollar_row:
        r += dr
    c = max(1, c)
    r = max(1, r)
    return dollar_col + to_col(c) + dollar_row + str(r)


def translate_formula(text, master_coord, cell_coord):
    """Translate a shared-formula master text to a shared cell's position.

    Relative references are shifted by the offset between *master_coord*
    and *cell_coord*; references anchored with $ keep their position.
    """
    m_col, m_row = split_coord(master_coord)
    c_col, c_row = split_coord(cell_coord)
    dc = c_col - m_col
    dr = c_row - m_row
    if dc == 0 and dr == 0:
        return text

    def repl(m):
        sheet_part, start_ref, end_ref = m.group(1), m.group(2), m.group(3)
        new_start = shift_ref(start_ref, dc, dr)
        if end_ref:
            return (sheet_part or "") + new_start + ":" + shift_ref(end_ref, dc, dr)
        return (sheet_part or "") + new_start

    return CELL_RE.sub(repl, text)


def sheet_targets(z):
    """Return [(sheet_name, worksheet_xml_part), ...] in workbook order."""
    wb_root = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.get("Id"): rel.get("Target")
        for rel in rels
        if (rel.get("Type") or "").endswith("/worksheet")
    }
    sheets = []
    for s in wb_root.find(_NS + "sheets"):
        rid = s.get(_RNS + "id")
        target = targets[rid]
        if not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")
        sheets.append((s.get("name"), target))
    return sheets


def _iter_sheet_cells(z, target):
    """Yield (coord, ftype, si, text) for every formula cell in a sheet."""
    cur = {"coord": None, "has_f": False, "t": None, "si": None, "text": None}
    for ev, el in ET.iterparse(z.open(target), events=("start", "end")):
        tag = el.tag
        if tag == _NS + "c":
            if ev == "start":
                cur["coord"] = el.get("r")
            else:
                if cur["has_f"]:
                    yield (cur["coord"], cur["t"], cur["si"], cur["text"])
                cur["has_f"] = False
        elif tag == _NS + "f":
            if ev == "end":
                cur["has_f"] = True
                cur["t"] = el.get("t")
                cur["si"] = el.get("si")
                cur["text"] = el.text
        if ev == "end":
            el.clear()


def iter_sheet_formulas(z, target):
    """Yield (coord, formula) for every formula cell, shared formulas resolved.

    Shared formulas are reconstructed from their master formula with
    relative references translated to the shared cell's position.
    """
    masters = {}
    for coord, ftype, si, text in _iter_sheet_cells(z, target):
        if ftype == "shared" and si is not None and text:
            masters[si] = (coord, text)
    for coord, ftype, si, text in _iter_sheet_cells(z, target):
        if text:
            yield coord, text
        elif ftype == "shared" and si is not None and si in masters:
            mcoord, mtext = masters[si]
            yield coord, translate_formula(mtext, mcoord, coord)
        else:
            yield coord, None


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
        m = COORD_RE.match(ref)
        return col_to_index(m.group(1)), int(m.group(2))


def cell_id(sheet, col, row):
    """Build a readable 'SheetName!A1' identifier."""
    return f"{sheet}!{to_col(col)}{row}"


class DependencyGraph:
    def __init__(self, workbook_path, output_sheet="Outputs"):
        self.output_sheet = output_sheet
        self.direct = {}  # (sheet, col, row) -> set of (sheet, col, row)
        self.dropped_refs = []  # {cell, function, dropped_arg}
        self.output_cells = []  # formula cells on the output sheet
        self.output_with_deps = 0  # output formula cells that reference something
        self._names = {}  # cell -> "SheetName!A1" (built once, reused everywhere)
        self._transitive_deps = {}  # cell_id_str -> [dep_id_str, ...]
        with zipfile.ZipFile(workbook_path) as z:
            self._index(z)
        self._build_transitive_deps()

    def _index(self, z):
        for sheet_name, target in sheet_targets(z):
            parser = FormulaParser(default_sheet=sheet_name)
            is_output = sheet_name == self.output_sheet
            for coord, ftext in iter_sheet_formulas(z, target):
                if not ftext:
                    continue
                col, row = split_coord(coord)
                key = (sheet_name, col, row)
                if is_output:
                    self.output_cells.append(key)
                deps, dropped = parser.parse(ftext)
                for func, arg in dropped:
                    self.dropped_refs.append(
                        {"cell": cell_id(*key), "function": func, "dropped_arg": arg}
                    )
                if deps:
                    self.direct[key] = {(sheet, c, r) for sheet, c, r in deps}
                    if is_output:
                        self.output_with_deps += 1

    def _name(self, cell):
        name = self._names.get(cell)
        if name is None:
            name = cell_id(*cell)
            self._names[cell] = name
        return name

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

    def _build_transitive_deps(self):
        """Precompute transitive deps for all output cells using bitset propagation.

        Assigns each output cell a bit index, then propagates bitsets through
        the dependency DAG in one pass.  SCCs (circular references) are
        condensed first so the DAG traversal is correct even with cycles.
        """
        from collections import defaultdict, deque

        n_out = len(self.output_cells)
        if n_out == 0:
            return

        sorted_outputs = sorted(self.output_cells)
        out_idx = {c: i for i, c in enumerate(sorted_outputs)}

        # --- Tarjan SCC on formula cells --------------------------------
        formula_cells = set(self.direct.keys())
        saved_limit = sys.getrecursionlimit()
        sys.setrecursionlimit(max(saved_limit, len(formula_cells) + 1000))
        try:
            idx_map = {}
            low = {}
            on_stack = set()
            stk = []
            counter = [0]
            sccs = []

            def _sc(v):
                idx_map[v] = low[v] = counter[0]
                counter[0] += 1
                stk.append(v)
                on_stack.add(v)
                for w in self.direct.get(v, ()):
                    if w not in formula_cells:
                        continue
                    if w not in idx_map:
                        _sc(w)
                        low[v] = min(low[v], low[w])
                    elif w in on_stack:
                        low[v] = min(low[v], idx_map[w])
                if low[v] == idx_map[v]:
                    scc = []
                    while True:
                        w = stk.pop()
                        on_stack.discard(w)
                        scc.append(w)
                        if w == v:
                            break
                    sccs.append(scc)

            for cell in formula_cells:
                if cell not in idx_map:
                    _sc(cell)
        finally:
            sys.setrecursionlimit(saved_limit)

        cell_to_super = {}
        super_cells = []
        for scc in sccs:
            sn = len(super_cells)
            super_cells.append(scc)
            for cell in scc:
                cell_to_super[cell] = sn
        n_supers = len(super_cells)

        # --- Deps of each super-node (formula + leaf cells) -------------
        super_dep_cells = [set() for _ in range(n_supers)]
        for sn, cells in enumerate(super_cells):
            for cell in cells:
                super_dep_cells[sn].update(self.direct[cell])

        # --- DAG edges between super-nodes + topological sort -----------
        dag_adj = [[] for _ in range(n_supers)]
        in_deg = [0] * n_supers
        for sn_a in range(n_supers):
            seen = set()
            for dep in super_dep_cells[sn_a]:
                if dep in cell_to_super:
                    sn_b = cell_to_super[dep]
                    if sn_b != sn_a and sn_b not in seen:
                        seen.add(sn_b)
                        dag_adj[sn_a].append(sn_b)
                        in_deg[sn_b] += 1

        queue = deque(i for i in range(n_supers) if in_deg[i] == 0)
        topo = []
        while queue:
            sn = queue.popleft()
            topo.append(sn)
            for dep_sn in dag_adj[sn]:
                in_deg[dep_sn] -= 1
                if in_deg[dep_sn] == 0:
                    queue.append(dep_sn)

        # --- Bitset propagation (reverse topological) -------------------
        super_inf = [0] * n_supers
        leaf_inf = defaultdict(int)

        for cell, idx in out_idx.items():
            sn = cell_to_super.get(cell)
            if sn is not None:
                super_inf[sn] |= 1 << idx

        for sn in topo:
            bits = super_inf[sn]
            if not bits:
                continue
            for dep in super_dep_cells[sn]:
                if dep in cell_to_super:
                    dep_sn = cell_to_super[dep]
                    if dep_sn != sn:
                        super_inf[dep_sn] |= bits
                else:
                    leaf_inf[dep] |= bits

        # --- Collect (cell, bitset) pairs and build output lists --------
        all_pairs = []
        for sn in range(n_supers):
            bits = super_inf[sn]
            if bits:
                for cell in super_cells[sn]:
                    all_pairs.append((cell, bits))
        for cell, bits in leaf_inf.items():
            all_pairs.append((cell, bits))

        if not all_pairs:
            return

        all_pairs.sort(key=lambda x: x[0])

        cid = {}
        for cell, _ in all_pairs:
            cid[cell] = cell_id(*cell)
        for cell in sorted_outputs:
            if cell not in cid:
                cid[cell] = cell_id(*cell)

        out_lists = [[] for _ in range(n_out)]
        i = 0
        n_pairs = len(all_pairs)
        while i < n_pairs:
            bits = all_pairs[i][1]
            j = i + 1
            while j < n_pairs and all_pairs[j][1] == bits:
                j += 1
            group_cells = [all_pairs[k][0] for k in range(i, j)]
            group_set = set(group_cells)
            b = bits
            while b:
                low_bit = b & (-b)
                idx = low_bit.bit_length() - 1
                own_cell = sorted_outputs[idx]
                if own_cell in group_set:
                    out_lists[idx].extend(
                        cid[c] for c in group_cells if c != own_cell
                    )
                else:
                    out_lists[idx].extend(cid[c] for c in group_cells)
                b ^= low_bit
            i = j

        for i, cell in enumerate(sorted_outputs):
            deps = out_lists[i]
            if deps:
                self._transitive_deps[cid[cell]] = deps


def iter_dependency_entries(graph):
    """Yield (cell_id, sorted dep cell ids) for every output cell with deps.

    Output cells are visited in a deterministic (column, row) order and each
    cell's result is produced and discarded on the fly.
    """
    if graph._transitive_deps:
        for cid_str in sorted(graph._transitive_deps):
            yield cid_str, graph._transitive_deps[cid_str]
    else:
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


def total_dep_entries(graph):
    """Total number of dependency→cell pairs across all output cells."""
    return sum(len(deps) for deps in graph._transitive_deps.values()) if graph._transitive_deps else 0


def summarize_workbook(graph, name):
    """Print a human-readable summary for one computed workbook."""
    cells = len(graph.output_cells)
    with_deps = count_dependency_entries(graph)
    total = total_dep_entries(graph)
    dropped = len(graph.dropped_refs)
    print(f"  → {name}: {cells} output formulas, "
          f"{with_deps} with deps, {total} dep pairs, "
          f"{dropped} dynamic refs ignored")


def print_skipped_summary(function_file_counts, transpose=False):
    """Print a table of skipped functions broken down by file.

    Default orientation: functions in rows, files in columns.
    With transpose=True: files in rows, functions in columns.
    Total counts are in the bottom row.
    """
    if not function_file_counts:
        return
    files = sorted({f for counts in function_file_counts.values() for f in counts})
    funcs = sorted(function_file_counts)
    if transpose:
        rows = files
        row_ids = files
        cols = funcs
        col_ids = funcs
        row_label = "File"
        first_col_w = max((len(f) for f in files), default=0)
    else:
        rows = funcs
        row_ids = funcs
        cols = files
        col_ids = files
        row_label = "Function"
        first_col_w = max((len(f) for f in funcs), default=0)
    col_ws = [max(len(c), 5) for c in cols]
    header = row_label.ljust(first_col_w) + "  " + "  ".join(c.ljust(w) for c, w in zip(cols, col_ws))
    sep = "=" * len(header)
    print(f"\n{sep}")
    print("Skipped functions by file")
    print(sep)
    print(header)
    print("-" * len(header))
    for rid in rows:
        cells = []
        for cid in col_ids:
            if transpose:
                cells.append(function_file_counts[cid].get(rid, 0))
            else:
                cells.append(function_file_counts[rid].get(cid, 0))
        print(f"{rid.ljust(first_col_w)}  " + "  ".join(str(v).rjust(w) for v, w in zip(cells, col_ws)))
    print("-" * len(header))
    totals = []
    if transpose:
        totals = [sum(function_file_counts[cid].values()) for cid in cols]
    else:
        totals = [sum(function_file_counts[r].get(c, 0) for r in funcs) for c in cols]
    print(f"{'Total'.ljust(first_col_w)}  " + "  ".join(str(v).rjust(w) for v, w in zip(totals, col_ws)))
    print(sep)


def _build_graph_with_timeout(file_path, output_sheet, timeout):
    if timeout <= 0:
        return DependencyGraph(file_path, output_sheet)
    import signal

    def _handler(signum, frame):
        raise TimeoutError(f"exceeded {timeout}s")

    prev = signal.signal(signal.SIGALRM, _handler)
    signal.alarm(timeout)
    try:
        graph = DependencyGraph(file_path, output_sheet)
    except TimeoutError:
        raise
    finally:
        signal.alarm(0)
        signal.signal(signal.SIGALRM, prev)
    return graph


def process_single(file_path, out_path, output_sheet, no_save=False, timeout=0):
    graph = _build_graph_with_timeout(file_path, output_sheet, timeout)
    meta = workbook_meta(
        file_path, output_sheet, datetime.now(timezone.utc).isoformat(),
        len(graph.output_cells), count_dependency_entries(graph), graph.dropped_refs,
    )
    if no_save:
        print(f"Calculated: {file_path}")
        summarize_workbook(graph, Path(file_path).name)
    else:
        write_output_file(out_path, meta, iter_dependency_entries(graph))
    del graph
    return meta


def process_directory(input_dir, output_dir, output_sheet, merge=False, no_save=False, timeout=0, transpose=False):
    """Process all *.xlsx files in *input_dir*.

    Returns 0 on success (even if some files failed), -1 on fatal error.
    With --merge the merged file is streamed workbook-by-workbook, so memory
    stays bounded regardless of how many/ how large the workbooks are.
    With --no-save nothing is written to disk; per-file summaries are printed.
    """
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    if not no_save:
        output_path.mkdir(parents=True, exist_ok=True)

    xlsx_files = sorted(input_path.glob("*.xlsx"))
    if not xlsx_files:
        print(f"No *.xlsx files found in '{input_dir}'.")
        return 0

    batch_t0 = time.time()
    total = len(xlsx_files)
    successful = 0
    failed = 0
    function_file_counts = defaultdict(lambda: defaultdict(int))

    merged_f = None
    first_workbook = True
    if merge and not no_save:
        merged_f = open(output_path / "merged_dependencies.json", "w", encoding="utf-8")
        merged_f.write('{\n  "workbooks": {')

    try:
        for idx, xlsx_file in enumerate(xlsx_files, 1):
            t0 = time.time()
            print(f"[{idx}/{total}] Processing: {xlsx_file.name}")
            try:
                graph = _build_graph_with_timeout(str(xlsx_file), output_sheet, timeout)
                try:
                    meta = workbook_meta(
                        str(xlsx_file), output_sheet, datetime.now(timezone.utc).isoformat(),
                        len(graph.output_cells), count_dependency_entries(graph), graph.dropped_refs,
                    )
                    for ref in graph.dropped_refs:
                        function_file_counts[ref["function"]][xlsx_file.name] += 1
                    if no_save:
                        summarize_workbook(graph, xlsx_file.name)
                    else:
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
            if no_save:
                print(f"  ✓ computed ({elapsed:.2f}s)")
            else:
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
    print_skipped_summary(function_file_counts, transpose)
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
    parser.add_argument(
        "--no-save", action="store_true", default=False,
        help="Calculate only: build the graph and print a summary, write nothing to disk"
    )
    parser.add_argument(
        "--timeout", type=int, default=0,
        help="Max seconds per file (0 = no limit). Skips files that exceed this."
    )
    parser.add_argument(
        "--transpose", action="store_true", default=False,
        help="Transpose the skipped-functions table: files in rows, functions in columns"
    )

    args = parser.parse_args(argv)

    # Batch mode
    input_dir = args.input_dir or (args.input if args.input and Path(args.input).is_dir() else None)
    if input_dir is not None and Path(input_dir).is_dir():
        return process_directory(input_dir, args.output_dir, args.output_sheet,
                                 args.merge, args.no_save, args.timeout, args.transpose)

    # Single-file mode
    try:
        meta = process_single(args.input, args.out, args.output_sheet, args.no_save, args.timeout)
    except TimeoutError as exc:
        print(f"ERROR: skipped — {exc}")
        return 1
    print(f"Cells traced: {meta['output_cells_with_dependencies']}")
    print(f"Dynamic refs ignored: {len(meta['ignored_dynamic_refs'])}")
    if not args.no_save:
        print(f"Written to: {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())