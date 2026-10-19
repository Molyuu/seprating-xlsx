"""Core logic: read workbooks, detect headers/tables, group rows, write outputs.

No column names or counts are hardcoded; the only special names are the
optional metadata columns 文件名称 / 工作表名称.
"""
from __future__ import annotations

import datetime as _dt
import re
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.cell import _writer as _cell_writer

_orig_safe_string = _cell_writer.safe_string


def _precise_safe_string(value):
    # openpyxl writes floats with 16 significant digits; use repr (17) so values
    # round-trip exactly like the source file.
    if isinstance(value, float) and value == value and value not in (float("inf"), float("-inf")):
        return repr(value)
    return _orig_safe_string(value)


_cell_writer.safe_string = _precise_safe_string

FILE_NAME_COL = "文件名称"
SHEET_NAME_COL = "工作表名称"
# Pseudo key sources usable in place of a real column.
INPUT_FILE = "<输入文件名>"
SOURCE_SHEET = "<源工作表名>"

DEFAULT_SUBTOTAL_WORDS = ("总计", "合计", "小计", "total", "subtotal")
EXCEL_EXT = (".xlsx", ".xlsm")


# --------------------------------------------------------------------------- data
@dataclass
class Cell:
    value: Any
    style: Optional[dict]  # copied style attributes or None


@dataclass
class Table:
    source_file: Path
    sheet: str
    header_row: int
    header: List[Cell]
    rows: List[List[Cell]]
    col_widths: Dict[int, float]

    @property
    def names(self) -> List[str]:
        return [_text(c.value) for c in self.header]


@dataclass
class Group:
    file_name: str  # raw (unsanitized) output file name
    sheet_name: str  # raw output sheet name
    table: Table
    rows: List[List[Cell]] = field(default_factory=list)
    subtotal_rows: int = 0


@dataclass
class Options:
    file_key: str = FILE_NAME_COL
    sheet_key: str = SHEET_NAME_COL
    include_meta: bool = False
    stamp_filename: bool = False
    stamp_row: bool = False
    timezone: str = "Asia/Singapore"
    subtotal_words: Tuple[str, ...] = DEFAULT_SUBTOTAL_WORDS


@dataclass
class Analysis:
    tables: List[Table]
    groups: List[Group]
    columns: List[str]  # union of header names across all tables
    warnings: List[str]


# --------------------------------------------------------------------------- helpers
def _text(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _blank(row: List[Cell]) -> bool:
    return all(_text(c.value) == "" for c in row)


def _style_of(cell) -> Optional[dict]:
    if not cell.has_style:
        return None
    return {
        "font": copy(cell.font), "fill": copy(cell.fill), "border": copy(cell.border),
        "alignment": copy(cell.alignment), "protection": copy(cell.protection),
        "number_format": cell.number_format,
    }


def _apply_style(cell, style: Optional[dict]) -> None:
    if style:
        for k, v in style.items():
            setattr(cell, k, copy(v) if k != "number_format" else v)


_BAD_FILE = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
_BAD_SHEET = re.compile(r"[\\/:*?\[\]]")
_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)),
             *(f"LPT{i}" for i in range(1, 10))}


def sanitize_filename(name: str) -> str:
    s = _BAD_FILE.sub("_", name).strip().rstrip(".") or "未命名"
    if s.upper() in _RESERVED:
        s = "_" + s
    return s[:150]


def sanitize_sheet(name: str) -> str:
    s = _BAD_SHEET.sub("_", name).strip().strip("'") or "Sheet"
    return s[:31]


def _unique(name: str, used: set, limit: int) -> str:
    cand, i = name, 2
    while cand.lower() in used:
        suffix = f"_{i}"
        cand = name[: limit - len(suffix)] + suffix
        i += 1
    used.add(cand.lower())
    return cand


def stamp_info(tz_name: str) -> Tuple[str, str]:
    """Return (YYYY-MM-DD, timezone label) for the given IANA zone."""
    try:
        from zoneinfo import ZoneInfo
        now = _dt.datetime.now(ZoneInfo(tz_name))
        return now.strftime("%Y-%m-%d"), tz_name
    except Exception:
        now = _dt.datetime.now().astimezone()
        return now.strftime("%Y-%m-%d"), (tz_name or str(now.tzinfo))


# --------------------------------------------------------------------------- reading
def collect_inputs(paths: List[str]) -> List[Path]:
    out: List[Path] = []
    for p in map(Path, paths):
        if p.is_dir():
            out += sorted(f for f in p.iterdir()
                          if f.suffix.lower() in EXCEL_EXT and not f.name.startswith("~$"))
        elif p.suffix.lower() in EXCEL_EXT:
            out.append(p)
    seen, uniq = set(), []
    for p in out:
        if p.resolve() not in seen:
            seen.add(p.resolve())
            uniq.append(p)
    return uniq


def _read_grid(ws) -> List[List[Cell]]:
    """Read a sheet into a grid, filling merged ranges with their top-left value."""
    max_r, max_c = ws.max_row, ws.max_column
    grid = [[Cell(c.value, _style_of(c)) for c in row]
            for row in ws.iter_rows(min_row=1, max_row=max_r, max_col=max_c)]
    for rng in ws.merged_cells.ranges:
        top = grid[rng.min_row - 1][rng.min_col - 1]
        for r in range(rng.min_row, min(rng.max_row, max_r) + 1):
            for c in range(rng.min_col, min(rng.max_col, max_c) + 1):
                grid[r - 1][c - 1] = Cell(top.value, top.style)
    return grid


def _is_header_like(row: List[Cell]) -> bool:
    vals = [c.value for c in row if _text(c.value)]
    return len(vals) >= 2 and all(isinstance(v, str) for v in vals)


def _detect_header(grid: List[List[Cell]], start: int, scan: int = 30) -> Optional[int]:
    """Pick the row (from `start`) with the most text cells; earliest wins ties."""
    best, best_score = None, 1
    for i in range(start, min(len(grid), start + scan)):
        # Count distinct labels so merged title rows (same text repeated) lose.
        score = len({c.value.strip() for c in grid[i] if isinstance(c.value, str) and c.value.strip()})
        if score > best_score:
            best, best_score = i, score
    return best


def read_tables(path: Path) -> List[Table]:
    wb = load_workbook(path, data_only=True)
    tables: List[Table] = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        grid = _read_grid(ws)
        widths = {i: ws.column_dimensions[get_column_letter(i)].width
                  for i in range(1, ws.max_column + 1)
                  if ws.column_dimensions[get_column_letter(i)].width}
        h = _detect_header(grid, 0)
        if h is None:
            continue
        header, rows, header_idx = grid[h], [], h
        i, after_blank = h + 1, False
        while i < len(grid):
            row = grid[i]
            if _blank(row):
                after_blank = True
            elif after_blank and _is_header_like(row) and i + 1 < len(grid) \
                    and not _is_header_like(grid[i + 1]) and not _blank(grid[i + 1]):
                # A new table begins after a blank gap.
                tables.append(_mk_table(path, ws.title, header_idx, header, rows, widths))
                header, rows, header_idx, after_blank = row, [], i, False
            else:
                rows.append(row)
                after_blank = False
            i += 1
        tables.append(_mk_table(path, ws.title, header_idx, header, rows, widths))
    return [t for t in tables if t.rows]


def _mk_table(path, sheet, idx, header, rows, widths) -> Table:
    # Trim trailing empty header columns that never hold data.
    width = len(header)
    while width > 0 and not _text(header[width - 1].value) \
            and all(width - 1 >= len(r) or not _text(r[width - 1].value) for r in rows):
        width -= 1
    hdr = list(header[:width])
    for j, c in enumerate(hdr):  # name unnamed columns
        if not _text(c.value):
            hdr[j] = Cell(f"列{j + 1}", c.style)
    return Table(path, sheet, idx + 1, hdr, [r[:width] for r in rows], widths)


# --------------------------------------------------------------------------- grouping
def default_keys(columns: List[str]) -> Tuple[str, str]:
    if FILE_NAME_COL in columns and SHEET_NAME_COL in columns:
        return FILE_NAME_COL, SHEET_NAME_COL
    company = next((c for c in columns if "公司" in c), None) \
        or next((c for c in columns if "company" in c.lower()), None)
    return (company or INPUT_FILE), SOURCE_SHEET


def _key_value(key: str, table: Table, row: List[Cell]) -> str:
    if key == INPUT_FILE:
        return table.source_file.stem
    if key == SOURCE_SHEET:
        return table.sheet
    names = table.names
    if key in names:
        return _text(row[names.index(key)].value)
    return ""


def _is_subtotal(row: List[Cell], words) -> bool:
    return any(w.lower() in _text(c.value).lower() for c in row for w in words)


def group_tables(tables: List[Table], opt: Options) -> Tuple[List[Group], List[str]]:
    groups: List[Group] = []
    warnings: List[str] = []
    by_key: Dict[Tuple[str, str], Group] = {}
    for t in tables:
        names = t.names
        for key, label in ((opt.file_key, "文件"), (opt.sheet_key, "工作表")):
            if key not in (INPUT_FILE, SOURCE_SHEET) and key not in names:
                warnings.append(f"{t.source_file.name}/{t.sheet}: 缺少列「{key}」,"
                                f"{label}名改用{'输入文件名' if label == '文件' else '源工作表名'}")
        fk = opt.file_key if opt.file_key in names or opt.file_key in (INPUT_FILE, SOURCE_SHEET) else INPUT_FILE
        sk = opt.sheet_key if opt.sheet_key in names or opt.sheet_key in (INPUT_FILE, SOURCE_SHEET) else SOURCE_SHEET
        current: Optional[Group] = None
        for row in t.rows:
            f, s = _key_value(fk, t, row), _key_value(sk, t, row)
            sub = _is_subtotal(row, opt.subtotal_words)
            # Rows missing key values (e.g. subtotal rows) stay with the group above.
            if current is not None and (not f or not s):
                current.rows.append(row)
                current.subtotal_rows += sub
                continue
            f, s = f or t.source_file.stem, s or t.sheet
            k = (f, s)
            if current is None or (current.file_name, current.sheet_name) != k:
                if k in by_key:
                    current = by_key[k]
                    if current.table is not t and current.table.names != names:
                        warnings.append(f"分组 {f}/{s} 出现在表头不同的表中,按首个表头输出")
                    else:
                        warnings.append(f"分组 {f}/{s} 不连续,已合并到同一工作表")
                else:
                    current = Group(f, s, t)
                    by_key[k] = current
                    groups.append(current)
            current.rows.append(row)
            current.subtotal_rows += sub
    return groups, warnings


def analyze(paths: List[Path], opt: Optional[Options] = None) -> Analysis:
    """Read inputs and group them. If opt is None, default keys are chosen."""
    tables: List[Table] = []
    warnings: List[str] = []
    for p in paths:
        try:
            tables += read_tables(p)
        except Exception as e:  # keep batch going
            warnings.append(f"无法读取 {p.name}: {e}")
    columns: List[str] = []
    for t in tables:
        columns += [n for n in t.names if n not in columns]
    if opt is None:
        opt = Options()
        opt.file_key, opt.sheet_key = default_keys(columns)
    groups, w = group_tables(tables, opt)
    return Analysis(tables, groups, columns, warnings + w)


# --------------------------------------------------------------------------- planning/writing
@dataclass
class OutputPlan:
    path: Path
    sheets: List[Tuple[str, Group]]


def plan_outputs(groups: List[Group], out_dir: Path, opt: Options) -> List[OutputPlan]:
    date, _ = stamp_info(opt.timezone)
    suffix = "_" + date.replace("-", "") if opt.stamp_filename else ""
    plans: Dict[str, OutputPlan] = {}
    sheet_used: Dict[str, set] = {}
    for g in groups:
        fname = sanitize_filename(g.file_name) + suffix
        key = fname.lower()
        if key not in plans:
            plans[key] = OutputPlan(out_dir / f"{fname}.xlsx", [])
            sheet_used[key] = set()
        sname = _unique(sanitize_sheet(g.sheet_name), sheet_used[key], 31)
        plans[key].sheets.append((sname, g))
    return list(plans.values())


def _output_columns(table: Table, opt: Options) -> List[int]:
    """Indexes of source columns to copy (metadata columns are handled separately)."""
    return [i for i, n in enumerate(table.names) if n not in (FILE_NAME_COL, SHEET_NAME_COL)]


def write_outputs(groups: List[Group], out_dir: Path, opt: Options) -> List[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    date, tz = stamp_info(opt.timezone)
    written = []
    for plan in plan_outputs(groups, out_dir, opt):
        wb = Workbook()
        wb.remove(wb.active)
        for sname, g in plan.sheets:
            ws = wb.create_sheet(sname)
            cols = _output_columns(g.table, opt)
            meta = [g.file_name, g.sheet_name] if opt.include_meta else []
            hdr_style = g.table.header[cols[0]].style if cols else None
            r = 1
            if opt.stamp_row:
                ws.cell(r, 1, f"生成时间: {date}, {tz}")
                r += 1
            for j, name in enumerate([FILE_NAME_COL, SHEET_NAME_COL][: len(meta)]):
                _apply_style(ws.cell(r, j + 1, name), hdr_style)
            for j, ci in enumerate(cols, start=len(meta) + 1):
                src = g.table.header[ci]
                _apply_style(ws.cell(r, j, src.value), src.style)
            ws.freeze_panes = ws.cell(r + 1, 1)
            r += 1
            for row in g.rows:
                for j, v in enumerate(meta, start=1):
                    ws.cell(r, j, v)
                for j, ci in enumerate(cols, start=len(meta) + 1):
                    src = row[ci] if ci < len(row) else Cell(None, None)
                    _apply_style(ws.cell(r, j, src.value), src.style)
                r += 1
            for j, ci in enumerate(cols, start=len(meta) + 1):
                w = g.table.col_widths.get(ci + 1)
                if w:
                    ws.column_dimensions[get_column_letter(j)].width = w
            for j in range(1, len(meta) + 1):
                ws.column_dimensions[get_column_letter(j)].width = 14
        wb.save(plan.path)
        written.append(plan.path)
    return written
