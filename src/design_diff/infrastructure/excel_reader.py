"""Excel から値・文章・表を取り出す（openpyxl。計算結果の値を読む）.

数式セルは、Excel が保存した計算結果（キャッシュ）を読む. ツール（openpyxl など）で作った Excel は計算結果を
持たないため、そのままでは数式セルが空になる. その場合は数式の文字列（例: =SUM(B4:B5)）を値として読み、
警告を出す（Excel で開いて保存し直すと計算結果で比べられる）.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from ..domain.model import text


def open_book(path: Path) -> Any:
    return load_book(path)[0]


def load_book(path: Path) -> tuple[Any, list[str]]:
    """計算結果の値で読んだブックと、警告（計算結果が保存されていない数式セルがあったとき）."""
    if not path.is_file():
        raise FileNotFoundError(f"Excel が見つかりません: {path}")
    wb = load_workbook(path, data_only=True)
    missing = _fill_missing_cached_values(path, wb)
    if not missing:
        return wb, []
    examples = "、".join(missing[:3]) + (" など" if len(missing) > 3 else "")
    return wb, [
        f"{path.name}: 計算結果が保存されていない数式セルが {len(missing)} 個あります（{examples}）。"
        "これらは計算結果ではなく数式の文字列（例: =SUM(B4:B5)）で比べました。ツールで作った Excel によくあります。"
        "Excel で開いて保存し直すと、計算結果で比べられます"
    ]


def _formula_text(cell: Any) -> str | None:
    if getattr(cell, "data_type", None) != "f" or cell.value is None:
        return None
    v = cell.value
    if isinstance(v, str):
        return v
    text_attr = getattr(v, "text", None)  # 配列数式（ArrayFormula）
    return text_attr if isinstance(text_attr, str) else str(v)


def _fill_missing_cached_values(path: Path, wb: Any) -> list[str]:
    """計算結果が無い数式セルに、数式の文字列を入れる. 入れたセルの位置（シート!セル）の一覧を返す."""
    formulas = load_workbook(path, read_only=True)
    missing: list[str] = []
    try:
        for fws in formulas.worksheets:
            if not hasattr(fws, "iter_rows") or fws.title not in wb.sheetnames:
                continue
            ws = wb[fws.title]
            for row in fws.iter_rows():
                for c in row:
                    f = _formula_text(c)
                    if f is None:
                        continue
                    target = ws.cell(c.row, c.column)
                    if target.value is None:
                        target.value = f
                        missing.append(f"{fws.title}!{target.coordinate}")
    finally:
        formulas.close()
    return missing


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
