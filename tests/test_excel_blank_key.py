"""表の差分で、キー列が空の行を比べずに警告すること（件数と行番号を示す）."""

from __future__ import annotations

from pathlib import Path

from test_robustness import TABLE_CONF, book, table_cells

from design_diff.application import excel_diff
from design_diff.domain.excel_diff import blank_key_rows


def test_blank_key_rows():
    rows = [{"項番": "1"}, {"項番": ""}, {"項番": " "}, {"項番": "2"}]
    assert blank_key_rows(["項番"], rows) == [1, 2]
    assert blank_key_rows(["a", "b"], [{"a": "", "b": "x"}, {"a": "", "b": ""}]) == [1]


def test_blank_key_rows_are_warned_with_count_and_row_numbers(tmp_path: Path):
    header = ["項番", "名前", "型"]
    before = book(tmp_path / "b.xlsx", table_cells(header, [[1, "a", "数値"], [None, "b", "文字列"], [2, "c", "数値"]]))
    after = book(
        tmp_path / "a.xlsx", table_cells(header, [[1, "a", "数値"], [None, "b2", "文字列"], [None, "x", "日付"], [2, "c", "数値"]])
    )
    result = excel_diff.run(TABLE_CONF, before, after)
    alerts = [w for rule, w in result.alerts if rule == "blank_key"]
    assert len(alerts) == 2, result.alerts
    assert "変更前" in alerts[0] and "1 行" in alerts[0] and "行 3）" in alerts[0]
    assert "変更後" in alerts[1] and "2 行" in alerts[1] and "行 3、4）" in alerts[1]
    # 空キーでない行に差分が無ければ、差分は 0 件（空キーの行は比べていないことを警告で知らせる）
    assert [c for c in result.changes if c.category not in ("outside_changed",)] == []


def test_no_blank_key_no_alert(tmp_path: Path):
    header = ["項番", "名前"]
    b = book(tmp_path / "b.xlsx", table_cells(header, [[1, "a"]]))
    assert not [r for r, _ in excel_diff.run(TABLE_CONF, b, b).alerts if r == "blank_key"]
