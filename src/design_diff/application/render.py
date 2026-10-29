"""差分のレポート（Markdown はレビューにそのまま貼れる形、JSON などは共通のレポート形式）."""

from __future__ import annotations

from devtools_common import report
from devtools_common.report import Finding, Severity

from ..domain.model import DiffResult
from ..domain.spec_diff import CATEGORIES as SPEC_CATEGORIES

EXCEL_CATEGORIES = {
    "sheet_added": "追加されたシート",
    "sheet_removed": "削除されたシート",
    "sheet_renamed": "シート名の変更",
    "value_changed": "値の変更",
    "label_added": "追加された項目",
    "label_removed": "削除された項目",
    "label_not_found": "どちらにも無いラベル（設定の見直しが必要）",
    "table_not_found": "どちらにも無い表（設定の見直しが必要）",
    "table_added": "追加された表",
    "table_removed": "削除された表",
    "column_added": "追加された列",
    "column_removed": "削除された列",
    "key_column_missing": "キー列が無い（設定の見直しが必要）",
    "duplicate_key": "キーの重複",
    "row_added": "追加された行",
    "row_removed": "削除された行",
    "cell_changed": "セルの変更",
    "outside_changed": "比較対象外のセルの変更（設定に無いセル）",
    "outside_cell_changed": "比較対象外のセルの変更（設定に無いセル）",
}
ALERT_LABELS = {
    "duplicate_column": "同じ名前の列",
    "table_truncated": "表の続きを比べていない",
    "duplicate_pair_key": "pair_by のキーが重複したシート",
    "sheet_too_large": "シートの値が多すぎて一部だけ読んだ",
    "blank_key": "キー列が空の行（比べていない）",
    "duplicate_api_id": "同じ API ID のシート・操作（最初のものだけ比べた）",
    "not_a_number": "数値として読めない値（比べていない）",
}
LABELS = {**EXCEL_CATEGORIES, **SPEC_CATEGORIES}
CONFIG_PROBLEMS = {"label_not_found", "table_not_found", "key_column_missing", "duplicate_key"}


def _html(s: str) -> str:
    """Markdown に入れる文字列の HTML（`&` `<` `>`）をエスケープする. セルの値の `<script>` や `<!--` を HTML として解釈させない."""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _cell(s: str) -> str:
    """表のセル. HTML をエスケープし、`\\` と `|` をエスケープして（devtools_common.report と同じ）、改行は `<br>` にする."""
    return _html(s).replace("\\", "\\\\").replace("|", "\\|").replace("\n", "<br>")


def _line(s: str) -> str:
    """表の外の 1 行（注意・注記・見出し）. HTML をエスケープし、改行は空白にする（引用の行が途切れないように）."""
    return _html(s).replace("\r", " ").replace("\n", " ")


def markdown(result: DiffResult) -> str:
    lines = [
        f"# {_line(result.title)}",
        "",
        f"- 左: {_line(result.left_label)}",
        f"- 右: {_line(result.right_label)}",
        f"- 差分: {len(result.changes)} 件",
        "",
    ]
    for w in result.warnings:
        lines.append(f"> **注意**: {_line(w)}")
    for _, w in result.alerts:
        lines.append(f"> **注意**: {_line(w)}")
    for n in result.notes:
        lines.append(f"> {_line(n)}")
    if result.notes or result.warnings or result.alerts:
        lines.append("")
    if not result.changes:
        return "\n".join(lines + ["差分はありません。", ""])
    lines += ["## 種類ごとの件数", "", "| 種類 | 件数 |", "|---|---|"]
    groups = result.by_category()
    lines += [f"| {LABELS.get(c, c)} | {len(v)} |" for c, v in groups.items()]
    for cat, changes in groups.items():
        lines += ["", f"## {LABELS.get(cat, cat)}", "", "| 場所 | 項目 | 左 | 右 |", "|---|---|---|---|"]
        lines += [f"| {_cell(c.location)} | {_cell(c.field)} | {_cell(c.before)} | {_cell(c.after)} |" for c in changes]
    return "\n".join(lines) + "\n"


def to_findings(result: DiffResult) -> list[Finding]:
    return (
        [
            Finding(
                c.category,
                Severity.ERROR if c.category in CONFIG_PROBLEMS else Severity.WARNING,
                LABELS.get(c.category, c.category),
                location=c.location,
                field=c.field,
                expected=c.before,
                actual=c.after,
            )
            for c in result.changes
        ]
        + [Finding("formula_cache_missing", Severity.WARNING, w) for w in result.warnings]
        + [Finding(rule, Severity.WARNING, w) for rule, w in result.alerts]
        + [Finding("note", Severity.INFO, n) for n in result.notes]
    )


def render(result: DiffResult, fmt: str, tool: str, version: str) -> str:
    if fmt == "markdown":
        return markdown(result)
    return report.render(
        to_findings(result), fmt, title=result.title, meta={"changes": len(result.changes)}, tool=tool, tool_version=version
    )
