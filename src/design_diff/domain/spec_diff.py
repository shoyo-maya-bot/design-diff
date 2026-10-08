"""API 定義書（Excel）と OpenAPI の比較. どちらも正規化した形にしてから比べる.

正規化した形（ApiSpec）:
    {"API0001": {"method": "GET", "path": "/api/v1/x/{id}",
                 "request": {"項目名": {"type": "string", "required": True, "max_length": 50, "location": "body"}},
                 "response": {...同じ形...},
                 "statuses": ["200", "404"]}}

    OpenAPI の oneOf / anyOf の候補の項目には "variant"（例: "oneOf 候補2"）が付く. required が None の項目は必須を比べない.

項目は「名前と受け渡し位置（location）」で識別する. path パラメーターの id と body の id のように、同じ名前で位置が違う項目は
別の項目として持つ（辞書のキーは 2 つ目から「名前 [位置]」で、その項目の "name" に元の名前）. 読み取り側は add_field で加える.
比べるときは、名前が両側で 1 つずつならそのまま対応付け（位置が違えば location_mismatch）、
同じ名前が複数あれば位置の同じもの同士を対応付ける.
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


def add_field(fields: dict[str, dict[str, Any]], name: str, entry: dict[str, Any]) -> str:
    """項目を名前と受け渡し位置で識別して加える（同じ名前・同じ位置なら置き換える）. 加えたキーを返す."""
    loc = _loc(entry)
    key = next((k for k, f in fields.items() if _name(k, f) == name and _loc(f) == loc), None)
    if key is None:
        key = name if name not in fields else f"{name} [{entry.get('location') or '-'}]"
        while key in fields:
            key += "'"
    fields[key] = {**entry, "name": name} if key != name else dict(entry)
    return key


def _name(key: str, f: dict[str, Any]) -> str:
    return str(f.get("name") or key)


def _loc(f: dict[str, Any]) -> str:
    return str(f.get("location") or "").strip().lower()


def _match(d: dict[str, dict[str, Any]], o: dict[str, dict[str, Any]]) -> tuple[list[tuple[str, str]], list[str], list[str]]:
    """項目の対応付け → (対応した (設計書のキー, OpenAPI のキー), 設計書にだけある, OpenAPI にだけある)."""
    o_by_name: dict[str, list[str]] = {}
    for k, f in o.items():
        o_by_name.setdefault(_name(k, f), []).append(k)
    d_by_name: dict[str, list[str]] = {}
    for k, f in d.items():
        d_by_name.setdefault(_name(k, f), []).append(k)
    pairs: list[tuple[str, str]] = []
    only_d: list[str] = []
    for name, dks in d_by_name.items():
        oks = list(o_by_name.get(name, []))
        if len(dks) == 1 and len(oks) == 1:
            pairs.append((dks[0], oks[0]))
            continue
        rest_d: list[str] = []
        for dk in dks:
            m = next((ok for ok in oks if _loc(o[ok]) == _loc(d[dk])), None)
            if m is None:
                m = next((ok for ok in oks if not _loc(o[ok]) or not _loc(d[dk])), None)
            if m is None:
                rest_d.append(dk)
            else:
                pairs.append((dk, m))
                oks.remove(m)
        if len(rest_d) == 1 and len(oks) == 1:  # 残りが 1 つずつなら位置の違う同じ項目とみなす
            pairs.append((rest_d[0], oks.pop()))
            rest_d = []
        only_d += rest_d
    paired_o = {ok for _, ok in pairs}
    only_o = [k for k in o if k not in paired_o]
    order = {k: i for i, k in enumerate(d)}
    pairs.sort(key=lambda p: order[p[0]])
    only_d.sort(key=lambda k: order[k])
    return pairs, only_d, only_o


def _label(key: str, f: dict[str, Any], d: dict[str, dict[str, Any]], o: dict[str, dict[str, Any]]) -> str:
    """表示する項目名. 同じ名前の項目が複数あるときは「名前（位置）」."""
    name = _name(key, f)
    count = max(sum(_name(k, x) == name for k, x in side.items()) for side in (d, o))
    return f"{name}（{f.get('location') or '-'}）" if count > 1 else name


def _fields(location: str, d: dict[str, dict[str, Any]], o: dict[str, dict[str, Any]]) -> list[Change]:
    out: list[Change] = []
    pairs, only_d, only_o = _match(d, o)
    for k in only_d:
        out.append(Change("field_missing_in_oas", location, _label(k, d[k], d, o)))
    for k in only_o:
        out.append(Change("field_missing_in_doc", location, _label(k, o[k], d, o), "", str(o[k].get("variant") or "")))
    for dk, ok in pairs:
        df, of = d[dk], o[ok]
        name = _label(dk, df, d, o)
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
