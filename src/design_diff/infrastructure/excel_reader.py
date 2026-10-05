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


def _own_value(ws: Worksheet, row: int, col: int) -> str:
    """結合の左上以外は空として読む（結合で埋めない）セルの値."""
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col and (row, col) != (rng.min_row, rng.min_col):
            return ""
    return _norm_label(ws.cell(row, col).value)


MAX_HEADER_ROWS = 3


def _count_header_rows(ws: Worksheet, head_row: int) -> int:
    """見出しが何段か（結合セルから判断する）.

    次の行も見出しとみなすのは、次のどちらかのとき:
    - 見出し行から始まる結合セルが、次の行まで縦に続いている（例: 2 段の見出しで縦に結合した「項目名」）
    - 見出し行で横に結合したセル（例: 「型情報」）の下の各列に、それぞれ別の値がある（例: 「型」「桁」）
    データ行も同じように横に結合しているだけの表は 1 段のままです.
    """
    n = 1
    while n < MAX_HEADER_ROWS:
        last, nxt = head_row + n - 1, head_row + n
        if nxt > ws.max_row:
            break
        sub = False
        for rng in ws.merged_cells.ranges:
            if not head_row <= rng.min_row <= last:
                continue
            if rng.max_row >= nxt:
                sub = True
                break
            if rng.max_row == last and rng.max_col > rng.min_col:
                below = [_own_value(ws, nxt, c) for c in range(rng.min_col, rng.max_col + 1)]
                if all(below) and len(set(below)) == len(below):
                    sub = True
                    break
        if not sub:
            break
        n += 1
    return n


def read_table(
    ws: Worksheet, header: str, title: str | None = None, header_rows: int | None = None
) -> tuple[list[str], list[dict[str, str]]] | None:
    """見出しの列名 header を含む行を見出し行とする表. title があれば、そのラベルの次の行を見出し行とする.

    見出しが複数段（結合したグループ見出し + 子見出し）のときは「親/子」の列名にする.
    段数は header_rows で指定する. 省略時は結合セルから判断する（_count_header_rows）.
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
    depth = header_rows or _count_header_rows(ws, head_row)
    cols: list[tuple[int, str]] = []
    for c in range(1, ws.max_column + 1):
        parts: list[str] = []
        for r in range(head_row, head_row + depth):
            v = _norm_label(_merged_value(ws, r, c))
            if v and (not parts or parts[-1] != v):
                parts.append(v)
        name = "/".join(parts)
        if name and name not in [n for _, n in cols]:
            cols.append((c, name))
    if not cols:
        return None
    rows: list[dict[str, str]] = []
    r = head_row + depth
    while r <= ws.max_row:
        values = {name: text(_merged_value(ws, r, c)) for c, name in cols}
        if not any(values.values()):
            break
        rows.append(values)
        r += 1
    return [n for _, n in cols], rows
