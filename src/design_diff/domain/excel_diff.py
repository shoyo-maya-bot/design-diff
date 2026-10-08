"""Excel 設計書の新旧比較（抽出済みの値どうしを比べる）.

比較方法（設定ファイルで項目ごとに選ぶ）:
- label: ラベルのセルからの相対位置の値（例: 「ジョブ ID」の右隣）
- block: ラベルの下に続く複数行を連結した文章（例: 「処理概要」の下の段落）
- table: 見出し行のある表。キー列で行を対応付け、追加・削除・変更された行とセルを出す
  （キーが重複した行は、同じキーの中で出現順に対応付ける. キー列が空の行は対応付けられないので比べない. blank_key_rows で警告する）
- 設定に無いセル: 行の並びを対応付けて（行の挿入・削除でずれても）値の違うセルを出す
"""

from __future__ import annotations

from difflib import SequenceMatcher

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

    reported: set[str] = set()

    def index(rows: list[dict[str, str]]) -> dict[str, dict[str, str]]:
        """キー → 行. 同じキーの 2 行目からは「キー（2 件目）」として別の行にする（後の行で上書きしない）."""
        out: dict[str, dict[str, str]] = {}
        count: dict[str, int] = {}
        for r in rows:
            k = " / ".join(text(r.get(c)) for c in key)
            if k.replace(" / ", ""):
                count[k] = count.get(k, 0) + 1
                if count[k] == 2 and k not in reported:
                    reported.add(k)
                    changes.append(Change("duplicate_key", location, k))
                out[k if count[k] == 1 else f"{k}（{count[k]} 件目）"] = r
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


def blank_key_rows(key: list[str], rows: list[dict[str, str]]) -> list[int]:
    """キー列がすべて空の行の位置（0 始まり）. compare_table はこの行を比べない（対応付けのキーが無いため）."""
    return [i for i, r in enumerate(rows) if not any(text(r.get(c)) for c in key)]


def column_letter(col: int) -> str:
    """列番号（1 始まり）→ Excel の列名（1 → A、27 → AA）."""
    s = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        s = chr(ord("A") + rem) + s
    return s


OutsideRow = tuple[int, dict[int, str]]  # (行番号, 列番号 → 値)


def compare_outside(before: list[OutsideRow], after: list[OutsideRow]) -> list[tuple[str, str, str]]:
    """設定に無いセルどうしを比べ、違うセルの (セル番地, 変更前, 変更後) を返す.

    行の並びを内容で対応付ける（difflib）ので、表の行が増えて下の欄がずれただけなら差分にしない.
    セル番地は変更後の位置（削除された行は変更前の位置）.
    """

    def key(row: OutsideRow) -> tuple[tuple[int, str], ...]:
        return tuple(sorted(row[1].items()))

    out: list[tuple[str, str, str]] = []

    def cells(b: OutsideRow | None, a: OutsideRow | None) -> None:
        bv, av = (b[1] if b else {}), (a[1] if a else {})
        row = a[0] if a else b[0]  # type: ignore[index]
        for c in sorted(set(bv) | set(av)):
            if bv.get(c, "") != av.get(c, ""):
                out.append((f"{column_letter(c)}{row}", bv.get(c, ""), av.get(c, "")))

    sm = SequenceMatcher(a=[key(r) for r in before], b=[key(r) for r in after], autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            continue
        bs, as_ = before[i1:i2], after[j1:j2]
        for i in range(max(len(bs), len(as_))):
            cells(bs[i] if i < len(bs) else None, as_[i] if i < len(as_) else None)
    return out


def _summary(row: dict[str, str], cols: list[str]) -> str:
    return "、".join(f"{c}={text(row.get(c))}" for c in cols if text(row.get(c)))
