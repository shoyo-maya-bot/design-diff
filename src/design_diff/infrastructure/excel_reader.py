"""Excel から値・文章・表を取り出す（openpyxl。計算結果の値を読む）.

セルの値は表示形式どおりの文字列で比べる（devtools_common.excel.render_value）. 例えば表示形式 `000` の数値 1 は
「001」、`#,##0` の 1500 は「1,500」、日付は `yyyy/mm/dd` なら「2026/04/05」. 文字列で入力した「001」と、
表示形式でゼロ埋めした数値 1 は同じ値とみなす（Excel の画面で同じに見えるため）.

数式セルは、Excel が保存した計算結果（キャッシュ）を読む. ツール（openpyxl など）で作った Excel は計算結果を
持たないため、そのままでは数式セルが空になる. その場合は数式の文字列（例: =SUM(B4:B5)）を値として読み、
警告を出す（Excel で開いて保存し直すと計算結果で比べられる）.
"""

from __future__ import annotations

import weakref
from pathlib import Path
from typing import Any

from devtools_common.excel import MergeMap, header_depth, header_names, load_workbook_with_formula_fallback, render_cell
from openpyxl.worksheet.worksheet import Worksheet

from ..domain.model import text

MAX_HEADER_ROWS = 3


def open_book(path: Path) -> Any:
    return load_book(path)[0]


def load_book(path: Path) -> tuple[Any, list[str]]:
    """計算結果の値で読んだブックと、警告（計算結果が保存されていない数式セルがあったとき）."""
    wb, missing = load_workbook_with_formula_fallback(path)
    if not missing:
        return wb, []
    examples = "、".join(missing[:3]) + (" など" if len(missing) > 3 else "")
    return wb, [
        f"{path.name}: 計算結果が保存されていない数式セルが {len(missing)} 個あります（{examples}）。"
        "これらは計算結果ではなく数式の文字列（例: =SUM(B4:B5)）で比べました。ツールで作った Excel によくあります。"
        "Excel で開いて保存し直すと、計算結果で比べられます"
    ]


_MERGES: weakref.WeakKeyDictionary[Any, MergeMap] = weakref.WeakKeyDictionary()


def _merges(ws: Worksheet) -> MergeMap:
    m = _MERGES.get(ws)
    if m is None:
        m = _MERGES[ws] = MergeMap.from_worksheet(ws)
    return m


def cell_text(ws: Worksheet, row: int, col: int) -> str:
    """セルの値を表示形式どおりの文字列で（前後の空白は除く）. 結合で埋めない."""
    return text(render_cell(ws.cell(row, col))[0])


def _merged_value(ws: Worksheet, row: int, col: int) -> str:
    """結合の左上の値で埋めたセルの値（表示形式どおりの文字列）."""
    return cell_text(ws, *_merges(ws).anchor(row, col))


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
    return _merged_value(ws, pos[0] + offset[0], pos[1] + offset[1])


def read_block(ws: Worksheet, label: str, offset: tuple[int, int] = (1, 0)) -> str | None:
    """ラベルの下に続く行（空行まで）を改行でつないだ文章."""
    pos = find_label(ws, label)
    if pos is None:
        return None
    r, c = pos[0] + offset[0], pos[1] + offset[1]
    lines = []
    while r <= ws.max_row:
        v = _merged_value(ws, r, c)
        if not v:
            break
        lines.append(v)
        r += 1
    return "\n".join(lines)


def _own_value(ws: Worksheet, row: int, col: int) -> str:
    """結合の左上以外は空として読む（結合で埋めない）セルの値."""
    return "" if _merges(ws).is_covered(row, col) else _norm_label(cell_text(ws, row, col))


def _count_header_rows(ws: Worksheet, head_row: int) -> int:
    """見出しが何段か（結合セルから判断する. devtools_common.excel.header_depth）.

    次の行も見出しとみなすのは、次のどちらかのとき:
    - 見出し行から始まる結合セルが次の行まで縦に続いていて、次の行が文字だけ（例: 縦に結合した「項目名」の横に「型」「桁」）
    - 見出し行で横に結合したセル（例: 「型情報」）の下の各列に、どれも文字の値がある（例: 「型」「桁」）
    データ行も同じように横に結合しているだけの表や、結合の下が数値の表は 1 段のままです.
    """
    return header_depth(
        _merges(ws),
        lambda r, c: _own_value(ws, r, c),
        head_row,
        first_col=1,
        last_col=ws.max_column,
        last_row=ws.max_row,
        max_rows=MAX_HEADER_ROWS,
    )


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
    names = header_names(lambda r, c: _norm_label(_merged_value(ws, r, c)), head_row, depth, range(1, ws.max_column + 1))
    cols: list[tuple[int, str]] = []
    for c, name in zip(range(1, ws.max_column + 1), names, strict=True):
        if name and name not in [n for _, n in cols]:
            cols.append((c, name))
    if not cols:
        return None
    rows: list[dict[str, str]] = []
    r = head_row + depth
    while r <= ws.max_row:
        values = {name: _merged_value(ws, r, c) for c, name in cols}
        if not any(values.values()):
            break
        rows.append(values)
        r += 1
    return [n for _, n in cols], rows
