"""差分の表現（ドメイン層。ファイル形式や出力形式に依存しない）."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class Change:
    """差分 1 件.

    category: 差分の種類（例: value_changed, row_added, field_missing_in_oas）
    location: どこの差分か（シート・表・API ID など）
    field: 項目名（無ければ空）
    before: 変更前（新旧差分）/ 設計書側（仕様差分）の値
    after: 変更後 / OpenAPI 側の値
    """

    category: str
    location: str
    field: str = ""
    before: str = ""
    after: str = ""


@dataclass
class DiffResult:
    title: str
    left_label: str
    right_label: str
    changes: list[Change] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    def by_category(self) -> dict[str, list[Change]]:
        out: dict[str, list[Change]] = {}
        for c in self.changes:
            out.setdefault(c.category, []).append(c)
        return out


def text(value: Any) -> str:
    """比較用に値を文字列へ（None と空白だけの値は空、前後の空白は無視、改行は保つ）."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()
