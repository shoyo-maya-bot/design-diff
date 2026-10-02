"""OpenAPI 3.x（YAML / JSON）を正規化した ApiSpec にする."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

METHODS = ("get", "post", "put", "patch", "delete")


def _resolve(doc: dict, node: Any, depth: int = 0) -> Any:
    if isinstance(node, dict) and "$ref" in node and depth < 20:
        ref = node["$ref"]
        if not ref.startswith("#/"):
            raise ValueError(f"外部ファイルの $ref には対応していません: {ref}")
        target: Any = doc
        for part in ref[2:].split("/"):
            target = target[part]
        return _resolve(doc, target, depth + 1)
    return node


def _props(doc: dict, schema: Any, location: str, prefix: str = "") -> dict[str, dict[str, Any]]:
    schema = _resolve(doc, schema or {})
    if schema.get("type") == "array":
        return _props(doc, schema.get("items"), location, prefix)
    out: dict[str, dict[str, Any]] = {}
    required = set(schema.get("required", []))
    for name, prop in (schema.get("properties") or {}).items():
        prop = _resolve(doc, prop)
        key = f"{prefix}{name}"
        out[key] = {
            "type": prop.get("type", "object"),
            "required": name in required,
            "max_length": prop.get("maxLength") or (prop.get("items") or {}).get("maxLength"),
            "location": location,
        }
        if prop.get("type") == "object" or "properties" in prop:
            out.update(_props(doc, prop, location, f"{key}."))
    return out


def read(path: Path) -> dict[str, dict[str, Any]]:
    doc = yaml.safe_load(path.read_text(encoding="utf-8"))
    if not isinstance(doc, dict) or "paths" not in doc:
        raise ValueError(f"OpenAPI の形式ではありません（paths がありません）: {path}")
    spec: dict[str, dict[str, Any]] = {}
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
                schema = _resolve(doc, prm.get("schema", {}))
                request[prm["name"]] = {
                    "type": schema.get("type", "string"),
                    "required": bool(prm.get("required", False)),
                    "max_length": schema.get("maxLength"),
                    "location": prm.get("in", "query"),
                }
            body = _resolve(doc, op.get("requestBody") or {})
            json_schema = ((body.get("content") or {}).get("application/json") or {}).get("schema")
            if json_schema:
                request.update(_props(doc, json_schema, "body"))
            responses = op.get("responses") or {}
            ok = next((_resolve(doc, r) for code, r in responses.items() if str(code).startswith("2")), {})
            res_schema = ((ok.get("content") or {}).get("application/json") or {}).get("schema")
            spec[str(api_id)] = {
                "method": m.upper(),
                "path": p,
                "request": request,
                "response": _props(doc, res_schema, "body") if res_schema else {},
                "statuses": sorted(str(c) for c in responses),
            }
    return spec
