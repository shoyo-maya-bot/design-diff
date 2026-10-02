"""Excel から値・文章・表を取り出す（openpyxl。計算結果の値を読む）."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from ..domain.model import text


def open_book(path: Path) -> Any:
    if not path.is_file():
        raise FileNotFoundError(f"Excel が見つかりません: {path}")
    return load_workbook(path, data_only=True)


def _merged_value(ws: Worksheet, row: int, col: int) -> Any:
    v = ws.cell(row, col).value
    if v is not None:
        return v
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return ws.cell(rng.min_row, rng.min_col).value
    return None


def _norm_label(v: Any) -> str:
    return text(v).rstrip(":：").replace("\n", "").replace(" ", "").replace("　", "")


def find_label(ws: Worksheet, label: str) -> tuple[int, int] | None:
    want = _norm_label(label)
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and _norm_label(c.value) == want:
                return c.row, c.column
    return None


def read_label(ws: Worksheet, label: str, offset: tuple[int, int] = (0, 1)) -> str | None:
    """ラベルのセルから offset（行, 列）離れたセルの値. ラベルが無ければ None."""
    pos = find_label(ws, label)
    if pos is None:
        return None
    return text(_merged_value(ws, pos[0] + offset[0], pos[1] + offset[1]))


def read_block(ws: Worksheet, label: str, offset: tuple[int, int] = (1, 0)) -> str | None:
    """ラベルの下に続く行（空行まで）を改行でつないだ文章."""
    pos = find_label(ws, label)
    if pos is None:
        return None
    r, c = pos[0] + offset[0], pos[1] + offset[1]
    lines = []
    while r <= ws.max_row:
        v = text(_merged_value(ws, r, c))
        if not v:
            break
        lines.append(v)
        r += 1
    return "\n".join(lines)


def read_table(ws: Worksheet, header: str, title: str | None = None) -> tuple[list[str], list[dict[str, str]]] | None:
    """見出しの列名 header を含む行を見出し行とする表. title があれば、そのラベルの次の行を見出し行とする.

    行は全列が空の行で終わる. 結合セルは左上の値で埋める.
    """
    if title:
        pos = find_label(ws, title)
        if pos is None:
            return None
        head_row = pos[0] + 1
    else:
        pos = find_label(ws, header)
        if pos is None:
            return None
        head_row = pos[0]
    cols: list[tuple[int, str]] = []
    for c in range(1, ws.max_column + 1):
        name = _norm_label(_merged_value(ws, head_row, c))
        if name and name not in [n for _, n in cols]:
            cols.append((c, name))
    if not cols:
        return None
    rows: list[dict[str, str]] = []
    r = head_row + 1
    while r <= ws.max_row:
        values = {name: text(_merged_value(ws, r, c)) for c, name in cols}
        if not any(values.values()):
            break
        rows.append(values)
        r += 1
    return [n for _, n in cols], rows
