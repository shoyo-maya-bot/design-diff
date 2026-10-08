"""Markdown のレポート: セルの値・注意の文の HTML（`<` `>` `&`）をエスケープし、表の列をずらさない."""

from __future__ import annotations

import html

import pytest
from devtools_common.review_results import table_cells, unescape_cell

from design_diff.application import render
from design_diff.domain.model import Change, DiffResult


def _result(before: str, after: str) -> DiffResult:
    r = DiffResult("新旧の差分", "変更前 a.xlsx", "変更後 b.xlsx")
    r.changes = [Change("cell_changed", "IF0001 / レイアウト", "備考", before, after)]
    return r


def _row(md: str) -> list[str]:
    line = next(x for x in md.splitlines() if x.startswith("| IF0001"))
    return [unescape_cell(c) for c in table_cells(line)]


@pytest.mark.parametrize(
    ("before", "after"),
    [
        ("<script>alert(1)</script>", "a & b"),
        ("<!-- 以下を隠す", "-->"),
        ("<b>太字</b>", "x > y"),
        ("a\\|b", "C:\\work\\"),  # バックスラッシュと | も列をずらさない
    ],
)
def test_cell_values_are_escaped_and_round_trip(before: str, after: str):
    md = render.markdown(_result(before, after))
    assert "<script" not in md and "<!--" not in md and "<b>" not in md
    cells = _row(md)
    assert len(cells) == 4
    assert [html.unescape(c) for c in cells[2:]] == [before, after]


def test_newline_in_cell_stays_br():
    md = render.markdown(_result("1行目\n2行目", ""))
    assert "1行目<br>2行目" in md


def test_notes_and_warnings_are_escaped():
    r = _result("x", "y")
    r.notes = ["シート「<IF>」を読み飛ばしました"]
    r.warnings = ["数式 =A1<B1 の計算結果がありません"]
    r.alerts = [("duplicate_column", "列「<備考>」が 2 列\nあります")]
    md = render.markdown(r)
    assert "&lt;IF&gt;" in md and "=A1&lt;B1" in md and "&lt;備考&gt;" in md
    assert "<IF>" not in md and "<備考>" not in md
    assert "> **注意**: 列「&lt;備考&gt;」が 2 列 あります" in md  # 引用の行を途切れさせない
