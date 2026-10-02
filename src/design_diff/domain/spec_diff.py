"""API 定義書（Excel）と OpenAPI の比較. どちらも正規化した形にしてから比べる.

正規化した形（ApiSpec）:
    {"API0001": {"method": "GET", "path": "/api/v1/x/{id}",
                 "request": {"項目名": {"type": "string", "required": True, "max_length": 50, "location": "body"}},
                 "response": {...同じ形...},
                 "statuses": ["200", "404"]}}
"""

from __future__ import annotations

from typing import Any

from .model import Change

ApiSpec = dict[str, dict[str, Any]]

CATEGORIES = {
    "api_missing_in_oas": "設計書にあるが OpenAPI に無い API",
    "api_missing_in_doc": "OpenAPI にあるが設計書に無い API",
    "method_mismatch": "HTTP メソッドの不一致",
    "path_mismatch": "パスの不一致",
    "field_missing_in_oas": "設計書にあるが OpenAPI に無い項目",
    "field_missing_in_doc": "OpenAPI にあるが設計書に無い項目",
    "type_mismatch": "型の不一致",
    "required_mismatch": "必須の不一致",
    "length_mismatch": "桁（最大長）の不一致",
    "location_mismatch": "受け渡し位置（path / query / body）の不一致",
    "status_missing_in_oas": "設計書にあるが OpenAPI に無い応答ステータス",
    "status_missing_in_doc": "OpenAPI にあるが設計書に無い応答ステータス",
}


def compare(doc: ApiSpec, oas: ApiSpec) -> list[Change]:
    changes: list[Change] = []
    for api_id in sorted(doc):
        if api_id not in oas:
            changes.append(Change("api_missing_in_oas", api_id))
    for api_id in sorted(oas):
        if api_id not in doc:
            changes.append(Change("api_missing_in_doc", api_id))
    for api_id in sorted(set(doc) & set(oas)):
        d, o = doc[api_id], oas[api_id]
        if d.get("method", "").upper() != o.get("method", "").upper():
            changes.append(Change("method_mismatch", api_id, "", d.get("method", ""), o.get("method", "")))
        if _norm_path(d.get("path", "")) != _norm_path(o.get("path", "")):
            changes.append(Change("path_mismatch", api_id, "", d.get("path", ""), o.get("path", "")))
        for part in ("request", "response"):
            changes += _fields(f"{api_id} {part}", d.get(part, {}), o.get(part, {}))
        ds, os_ = set(map(str, d.get("statuses", []))), set(map(str, o.get("statuses", [])))
        if ds:  # 設計書に応答の一覧があるときだけ比べる
            changes += [Change("status_missing_in_oas", api_id, s) for s in sorted(ds - os_)]
            changes += [Change("status_missing_in_doc", api_id, s) for s in sorted(os_ - ds)]
    return changes


def _fields(location: str, d: dict[str, dict[str, Any]], o: dict[str, dict[str, Any]]) -> list[Change]:
    out: list[Change] = []
    for name in d:
        if name not in o:
            out.append(Change("field_missing_in_oas", location, name))
    for name in o:
        if name not in d:
            out.append(Change("field_missing_in_doc", location, name))
    for name in [n for n in d if n in o]:
        df, of = d[name], o[name]
        for key, cat in (("type", "type_mismatch"), ("location", "location_mismatch")):
            if df.get(key) and of.get(key) and df[key] != of[key]:
                out.append(Change(cat, location, name, str(df[key]), str(of[key])))
        if df.get("required") is not None and of.get("required") is not None and bool(df["required"]) != bool(of["required"]):
            out.append(Change("required_mismatch", location, name, _yn(df["required"]), _yn(of["required"])))
        if df.get("max_length") and df.get("max_length") != of.get("max_length"):
            out.append(Change("length_mismatch", location, name, str(df["max_length"]), str(of.get("max_length") or "（指定なし）")))
    return out


def _norm_path(p: str) -> str:
    import re

    return re.sub(r"\{[^}]+\}", "{}", p.rstrip("/")) or "/"


def _yn(v: Any) -> str:
    return "必須" if v else "任意"
