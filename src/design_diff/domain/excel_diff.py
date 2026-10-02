"""Excel 設計書の新旧比較（抽出済みの値どうしを比べる）.

比較方法（設定ファイルで項目ごとに選ぶ）:
- label: ラベルのセルからの相対位置の値（例: 「ジョブ ID」の右隣）
- block: ラベルの下に続く複数行を連結した文章（例: 「処理概要」の下の段落）
- table: 見出し行のある表。キー列で行を対応付け、追加・削除・変更された行とセルを出す
"""

from __future__ import annotations

from .model import Change, text

Table = tuple[list[str], list[dict[str, str]]]  # (列名の並び, 行の並び)


def compare_value(location: str, name: str, before: str | None, after: str | None) -> list[Change]:
    if before is None and after is None:
        return [Change("label_not_found", location, name)]
    if before is None:
        return [Change("label_added", location, name, "", text(after))]
    if after is None:
        return [Change("label_removed", location, name, text(before), "")]
    if text(before) != text(after):
        return [Change("value_changed", location, name, text(before), text(after))]
    return []


def compare_table(location: str, key: list[str], before: Table, after: Table) -> list[Change]:
    b_cols, b_rows = before
    a_cols, a_rows = after
    changes: list[Change] = []
    for col in b_cols:
        if col not in a_cols:
            changes.append(Change("column_removed", location, col))
    for col in a_cols:
        if col not in b_cols:
            changes.append(Change("column_added", location, col))
    missing_key = [k for k in key if k not in b_cols or k not in a_cols]
    if missing_key:
        changes.append(Change("key_column_missing", location, "・".join(missing_key)))
        return changes

    def index(rows: list[dict[str, str]]) -> dict[str, dict[str, str]]:
        out: dict[str, dict[str, str]] = {}
        for r in rows:
            k = " / ".join(text(r.get(c)) for c in key)
            if k.replace(" / ", ""):
                if k in out:
                    changes.append(Change("duplicate_key", location, k))
                out[k] = r
        return out

    b_idx, a_idx = index(b_rows), index(a_rows)
    common_cols = [c for c in b_cols if c in a_cols and c not in key]
    for k, r in b_idx.items():
        if k not in a_idx:
            changes.append(Change("row_removed", location, k, _summary(r, b_cols), ""))
    for k, r in a_idx.items():
        if k not in b_idx:
            changes.append(Change("row_added", location, k, "", _summary(r, a_cols)))
            continue
        for c in common_cols:
            bv, av = text(b_idx[k].get(c)), text(r.get(c))
            if bv != av:
                changes.append(Change("cell_changed", location, f"{k} の {c}", bv, av))
    return changes


def _summary(row: dict[str, str], cols: list[str]) -> str:
    return "、".join(f"{c}={text(row.get(c))}" for c in cols if text(row.get(c)))
