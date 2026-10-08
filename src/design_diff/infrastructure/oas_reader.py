"""OpenAPI 3.x（YAML / JSON）を正規化した ApiSpec にする.

スキーマの組み合わせの読み方:
- allOf: 全部の properties と required をつないで 1 つのスキーマとして読む（$ref も、入れ子の allOf も展開する）
- oneOf / anyOf: 全候補の項目を合わせて（和集合で）読む. 候補の項目には "variant"（例: "oneOf 候補1・2"）を付け、
  候補があることをメモ（notes）に出す. 必須は「全候補で必須なら必須、どの候補でも任意なら任意、それ以外は比べない（None）」.
  null との oneOf / anyOf（null を許すだけの書き方）は、null でない方のスキーマとして読み、メモは出さない
- discriminator: 候補の選び分けには使わない. 判別用の項目名をメモに添えるだけ（項目自体は各候補の properties から読む）
- 自分自身を参照するスキーマ（木構造など）は、2 回目に出てきたところで展開をやめ、メモに出す
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

METHODS = ("get", "post", "put", "patch", "delete")
VARIANT_KEYS = ("oneOf", "anyOf")
MAX_DEPTH = 20


def _resolve_once(doc: dict, node: Any) -> Any:
    ref = node["$ref"]
    if not ref.startswith("#/"):
        raise OpenApiError(f"外部ファイルの $ref には対応していません: {ref}")
    target: Any = doc
    for part in ref[2:].split("/"):
        part = part.replace("~1", "/").replace("~0", "~")  # JSON Pointer のエスケープ
        if not isinstance(target, dict) or part not in target:
            raise OpenApiError(f"$ref の参照先がありません: {ref}")
        target = target[part]
    return target


def _resolve(doc: dict, node: Any, depth: int = 0) -> Any:
    if isinstance(node, dict) and "$ref" in node and depth < MAX_DEPTH:
        return _resolve(doc, _resolve_once(doc, node), depth + 1)
    return node


def _ref_chain(doc: dict, node: Any) -> list[str]:
    refs: list[str] = []
    while isinstance(node, dict) and "$ref" in node and len(refs) < MAX_DEPTH:
        refs.append(node["$ref"])
        node = _resolve_once(doc, node)
    return refs


def _merge_props(props: dict[str, Any], more: dict[str, Any]) -> None:
    for name, prop in more.items():
        props[name] = {"allOf": [props[name], prop]} if name in props else prop


def _shape(doc: dict, schema: Any, depth: int = 0) -> dict[str, Any]:
    """$ref を解決し、allOf をつないだ 1 つのスキーマ. oneOf / anyOf は "_variants"（[(種類, 候補の並び)]）に集める."""
    s = _resolve(doc, schema)
    if not isinstance(s, dict):
        return {}
    out = {k: v for k, v in s.items() if k not in ("allOf", *VARIANT_KEYS)}
    variants: list[tuple[str, list]] = [(k, s[k]) for k in VARIANT_KEYS if isinstance(s.get(k), list)]
    if isinstance(s.get("allOf"), list) and depth < MAX_DEPTH:
        props: dict[str, Any] = {}
        required: list[str] = []
        for sub in s["allOf"]:
            sh = _shape(doc, sub, depth + 1)
            _merge_props(props, sh.get("properties") or {})
            required += sh.get("required") or []
            variants += sh.get("_variants") or []
            for k, v in sh.items():
                if k not in ("properties", "required", "_variants"):
                    out.setdefault(k, v)  # 自分の値 → 先の allOf の値 の順に優先
        _merge_props(props, s.get("properties") or {})
        required += s.get("required") or []
        if props:
            out["properties"] = props
        if required:
            out["required"] = required
    if variants:
        out["_variants"] = variants
    return _collapse_nullable(doc, out, depth)


def _collapse_nullable(doc: dict, shape: dict[str, Any], depth: int) -> dict[str, Any]:
    """oneOf: [X, {type: null}] のような「null を許すだけ」の書き方は X として読む."""
    v = shape.get("_variants")
    if not v or len(v) != 1 or shape.get("type") or shape.get("properties"):
        return shape
    options = [_shape(doc, o, depth + 1) for o in v[0][1]]
    non_null = [o for o in options if o.get("type") != "null"]
    if len(options) > 1 and len(non_null) == 1:
        return {**non_null[0], **{k: val for k, val in shape.items() if k != "_variants"}}
    return shape


def _type(doc: dict, shape: dict[str, Any]) -> str | None:
    """型. OpenAPI 3.1 の type: [string, "null"] は string. oneOf / anyOf で候補の型が違うときは None（比べない）."""
    t = shape.get("type")
    if isinstance(t, list):
        ts = [x for x in t if x != "null"]
        return ts[0] if len(ts) == 1 else None
    if t:
        return t
    if shape.get("properties"):
        return "object"
    if shape.get("_variants"):
        types = {_type(doc, _shape(doc, o)) for _, options in shape["_variants"] for o in options}
        types.discard("null")
        return types.pop() if len(types) == 1 else None
    return "object"


def _max_length(doc: dict, shape: dict[str, Any]) -> Any:
    return shape.get("maxLength") or _shape(doc, shape.get("items")).get("maxLength")


def _object_fields(doc: dict, shape: dict[str, Any], where: str, notes: list[str], depth: int = 0) -> dict[str, dict[str, Any]]:
    """項目名 → {"schema": 項目のスキーマ, "required": True/False/None, "variant": 候補の印}."""
    required = set(shape.get("required") or [])
    fields = {name: {"schema": prop, "required": name in required, "variant": ""} for name, prop in (shape.get("properties") or {}).items()}
    for kind, options in shape.get("_variants") or []:
        option_fields: list[dict[str, dict[str, Any]]] = []
        summaries: list[str] = []
        for i, opt in enumerate(options, 1):
            osh = _shape(doc, opt)
            of = _object_fields(doc, osh, where, notes, depth + 1) if depth < MAX_DEPTH else {}
            option_fields.append(of)
            summaries.append(f"候補{i} = " + (", ".join(of) if of else f"型 {_type(doc, osh) or '不明'}"))
        union: dict[str, dict[str, Any]] = {}
        for i, of in enumerate(option_fields, 1):
            for name, f in of.items():
                u = union.setdefault(name, {"schema": f["schema"], "req": [], "idx": []})
                u["req"].append(f["required"])
                u["idx"].append(i)
        mixed: list[str] = []
        for name, u in union.items():
            in_all = len(u["idx"]) == len(options)
            if name in fields:  # 親の properties にもある項目は親の定義を使う
                if in_all and all(r is True for r in u["req"]):
                    fields[name]["required"] = True
                continue
            if name in required or (in_all and all(r is True for r in u["req"])):
                req: bool | None = True
            elif all(r is False for r in u["req"]):
                req = False
            else:
                req = None
                mixed.append(name)
            fields[name] = {"schema": u["schema"], "required": req, "variant": f"{kind} 候補{'・'.join(map(str, u['idx']))}"}
        note = (
            f"OpenAPI の {where} に {kind} があります（候補 {len(options)} 個: {' / '.join(summaries)}）。項目は全候補を合わせて比べました"
        )
        if mixed:
            note += f"。候補によって必須が違う項目（{', '.join(mixed)}）は必須を比べていません"
        disc = shape.get("discriminator")
        if isinstance(disc, dict) and disc.get("propertyName"):
            note += f"。判別用の項目（discriminator）: {disc['propertyName']}"
        notes.append(note + "。どの候補の項目か、設計書の書き方と合っているかを人が確認してください")
    return fields


def _props(
    doc: dict,
    schema: Any,
    location: str,
    label: str,
    notes: list[str],
    prefix: str = "",
    seen: frozenset[str] = frozenset(),
    depth: int = 0,
) -> dict[str, dict[str, Any]]:
    refs = _ref_chain(doc, schema)
    if (set(refs) & seen) or depth > MAX_DEPTH:
        notes.append(f"OpenAPI の {label} の {prefix.rstrip('.')} は循環参照のため、中の項目は展開していません")
        return {}
    seen = seen | set(refs)
    shape = _shape(doc, schema or {})
    if _type(doc, shape) == "array":
        return _props(doc, shape.get("items"), location, label, notes, prefix, seen, depth + 1)
    where = f"{label} の {prefix.rstrip('.')}" if prefix else label
    out: dict[str, dict[str, Any]] = {}
    for name, f in _object_fields(doc, shape, where, notes).items():
        prop = _shape(doc, f["schema"])
        key = f"{prefix}{name}"
        t = _type(doc, prop)
        out[key] = {"type": t, "required": f["required"], "max_length": _max_length(doc, prop), "location": location}
        if f["variant"]:
            out[key]["variant"] = f["variant"]
        if prop.get("type") == "object" or prop.get("properties") or (prop.get("_variants") and t != "array"):
            out.update(_props(doc, f["schema"], location, label, notes, f"{key}.", seen, depth + 1))
    return out


def read(path: Path) -> dict[str, dict[str, Any]]:
    return read_with_notes(path)[0]


class OpenApiError(ValueError):
    """OpenAPI を読めない（YAML / JSON の構文の誤り、形式の誤り）. 入力の誤りとして扱う（CLI は終了コード 2）."""


def read_with_notes(path: Path) -> tuple[dict[str, dict[str, Any]], list[str]]:
    """ApiSpec と、人が確認すべきこと（oneOf / anyOf・循環参照）のメモ.

    YAML の構文の誤りや、辞書であるべきところが辞書でないなどの構造の誤りは OpenApiError（ValueError）にする.
    """
    try:
        doc = yaml.safe_load(path.read_text(encoding="utf-8"))
    except yaml.YAMLError as e:
        raise OpenApiError(f"OpenAPI の YAML / JSON を読めません（{_brief(e)}）: {path}") from e
    except UnicodeDecodeError as e:
        raise OpenApiError(f"OpenAPI を UTF-8 として読めません: {path}") from e
    if not isinstance(doc, dict) or not isinstance(doc.get("paths"), dict):
        raise OpenApiError(f"OpenAPI の形式ではありません（paths がありません）: {path}")
    try:
        return _read(doc)
    except OpenApiError:
        raise
    except (AttributeError, TypeError, KeyError, IndexError, ValueError) as e:
        raise OpenApiError(f"OpenAPI の構造が想定と違います（{type(e).__name__}: {e}）: {path}") from e


def _brief(e: yaml.YAMLError) -> str:
    mark = getattr(e, "problem_mark", None)
    problem = getattr(e, "problem", None) or type(e).__name__
    return f"{mark.line + 1} 行 {mark.column + 1} 列: {problem}" if mark is not None else str(problem)


def _media_type(key: Any) -> str:
    """メディアタイプの本体（パラメーター「; charset=…」を除き、小文字）."""
    return str(key).split(";", 1)[0].strip().lower()


def is_json_media_type(key: Any) -> bool:
    """JSON のメディアタイプか. application/json、*+json（application/problem+json など）、パラメーター付きも含む."""
    mt = _media_type(key)
    return mt == "application/json" or (mt.startswith("application/") and mt.endswith("+json"))


def json_schema(content: Any) -> Any:
    """content（メディアタイプ → Media Type Object）から比べるスキーマを選ぶ.

    1. application/json（パラメーター付きも）2. *+json 3. */* の順. JSON が無く */* も無ければ None.
    """
    if not isinstance(content, dict):
        return None
    keys = list(content)
    ranked = (
        [k for k in keys if _media_type(k) == "application/json"]
        + [k for k in keys if is_json_media_type(k) and _media_type(k) != "application/json"]
        + [k for k in keys if _media_type(k) == "*/*"]
    )
    for k in ranked:
        media = content.get(k) or {}
        if isinstance(media, dict) and media.get("schema") is not None:
            return media.get("schema")
    return None


def _read(doc: dict) -> tuple[dict[str, dict[str, Any]], list[str]]:
    spec: dict[str, dict[str, Any]] = {}
    notes: list[str] = []
    for p, item in (doc.get("paths") or {}).items():
        common = item.get("parameters", [])
        for m in METHODS:
            op = item.get(m)
            if not op:
                continue
            api_id = op.get("x-api-id") or op.get("operationId") or f"{m.upper()} {p}"
            request: dict[str, dict[str, Any]] = {}
            for prm in [*common, *op.get("parameters", [])]:
                prm = _resolve(doc, prm)
                schema = _shape(doc, prm.get("schema", {}))
                request[prm["name"]] = {
                    "type": _type(doc, schema) if (schema.get("type") or schema.get("_variants")) else "string",
                    "required": bool(prm.get("required", False)),
                    "max_length": schema.get("maxLength"),
                    "location": prm.get("in", "query"),
                }
            body = _resolve(doc, op.get("requestBody") or {})
            req_schema = json_schema(body.get("content"))
            if req_schema:
                request.update(_props(doc, req_schema, "body", f"{api_id} request", notes))
            responses = op.get("responses") or {}
            ok = next((_resolve(doc, r) for code, r in responses.items() if str(code).startswith("2")), {})
            res_schema = json_schema(ok.get("content"))
            spec[str(api_id)] = {
                "method": m.upper(),
                "path": p,
                "request": request,
                "response": _props(doc, res_schema, "body", f"{api_id} response", notes) if res_schema else {},
                "statuses": sorted(str(c) for c in responses),
            }
    return spec, list(dict.fromkeys(notes))
