"""design-diff — 設計書のレビューで使う差分レポート.

  design-diff excel --type api --before 旧.xlsx --after 新.xlsx          # 設定は config/api.yaml
  design-diff excel --all                                                # config/global.yaml の pairs を一括
  design-diff spec --excel API定義書.xlsx --oas openapi.yaml             # 設定は config/spec.yaml

出力: 既定は Markdown（レビューにそのまま貼れる）。-f json で機械可読。-o でファイルへ。
  --all はペアごとのレポートを output_dir に -f の形式で書き、一覧（-f の形式）を標準出力か -o に出す。
  global.yaml の pairs・output_dir の相対パスは global.yaml のあるフォルダから解決する（実行場所に依らない）。
終了コード: 0 正常（差分の有無にかかわらず）/ 1 --fail-on-diff で差分あり、設定の見直しが必要、
  または --all で比べられなかったペアがある / 2 引数・設定・入力ファイルの誤り（壊れた・形式の違う Excel を含む）
  / 3 内部エラー（想定外の例外）
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from devtools_common import cli as common_cli
from devtools_common import report
from devtools_common.config import load_structured
from devtools_common.report import Finding, Severity

from .. import __version__
from ..application import excel_diff, render, spec_diff
from ..infrastructure.excel_reader import UnreadableBookError

TOOL = "design-diff"
EXTENSIONS = {"markdown": ".md", "json": ".json", "html": ".html", "csv": ".csv"}
CONFIG_DIR_HELP = "設定ファイルの置き場（既定: config/）。サブコマンドの前後どちらに書いてもよい"


def _add_config_dir(parser: argparse.ArgumentParser, *, top: bool) -> None:
    # サブコマンドの後に書いても受け付ける. サブコマンド側は既定値を持たない（前に書いた値を上書きしない）
    parser.add_argument("--config-dir", type=Path, default=Path("config") if top else argparse.SUPPRESS, help=CONFIG_DIR_HELP)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog=TOOL, description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    _add_config_dir(p, top=True)
    sub = p.add_subparsers(dest="command", required=True)

    e = sub.add_parser("excel", help="Excel 設計書の新旧差分")
    e.add_argument("--type", help="設定の種類（config/<type>.yaml）。例: api / batch / if")
    e.add_argument("--before", type=Path)
    e.add_argument("--after", type=Path)
    e.add_argument(
        "--all",
        action="store_true",
        help="config/global.yaml の pairs をすべて比較し、output_dir に -f の形式で書き出す（-o には一覧を書く）",
    )
    e.add_argument(
        "--outside",
        choices=excel_diff.OUTSIDE_MODES,
        default="summary",
        help="設定に無いセル（承認者欄など）の変更: summary=シートごとに件数と場所（既定）/ all=1 セルずつ / off=比べない",
    )
    e.add_argument("--fail-on-diff", action="store_true")
    _add_config_dir(e, top=False)
    common_cli.add_common_args(e)

    s = sub.add_parser("spec", help="API 定義書（Excel）と OpenAPI の差分")
    s.add_argument("--excel", type=Path, required=True)
    s.add_argument("--oas", type=Path, required=True)
    s.add_argument("--config", type=Path, help="既定: config/spec.yaml")
    s.add_argument("--fail-on-diff", action="store_true")
    _add_config_dir(s, top=False)
    common_cli.add_common_args(s)
    return p


def _config(path: Path) -> dict:
    if not path.is_file():
        raise common_cli.UsageError(f"設定ファイルが見つかりません: {path}")
    return load_structured(path) or {}


def _exit(result, fail_on_diff: bool) -> int:
    config_problem = any(c.category in render.CONFIG_PROBLEMS for c in result.changes)
    if config_problem or (fail_on_diff and result.changes):
        return common_cli.ExitCode.FINDINGS
    return common_cli.ExitCode.OK


def _unique_dest(out_dir: Path, stem: str, ext: str, used: set[str]) -> Path:
    """出力名が他のペアと重なれば -2, -3 … を付ける（大文字小文字は区別しない. Windows・macOS に合わせる）."""
    name, n = f"diff_{stem}", 1
    while f"{name}{ext}".casefold() in used:
        n += 1
        name = f"diff_{stem}-{n}"
    used.add(f"{name}{ext}".casefold())
    return out_dir / f"{name}{ext}"


def _resolve(base: Path, value: str, what: str, *, must_exist: bool) -> Path:
    """global.yaml に書いた相対パスを、global.yaml のあるフォルダから解決する（実行場所に依らない）.

    互換: global.yaml 基準で見つからず、実行場所からの相対なら見つかるときは、そちらを使って警告する
    （以前は実行場所からの相対だった）. 絶対パスはそのまま.
    """
    p = Path(value)
    if p.is_absolute():
        return p
    resolved = base / p
    legacy = Path.cwd() / p
    exists = resolved.exists() if must_exist else resolved.parent.exists() or resolved.exists()
    if not exists and legacy.exists() and legacy.resolve() != resolved.resolve():
        print(
            f"design-diff: global.yaml の {what} {value} は global.yaml のフォルダ（{base.as_posix()}）からの相対では見つからないため、"
            "実行場所からの相対として使いました（実行場所からの相対は今後なくなります。global.yaml からの相対パスに書き換えてください）",
            file=sys.stderr,
        )
        return legacy
    return resolved


def cmd_all(args) -> int:
    """global.yaml の pairs をすべて比べる. 欠けたペア・読めないペアは警告して続け、最後に終了コード 1 で知らせる.

    pairs の before / after と output_dir の相対パスは global.yaml のあるフォルダから解決する.
    """
    g_path = args.config_dir / "global.yaml"
    g = _config(g_path)
    base = g_path.resolve().parent
    pairs = g.get("pairs") or []
    confs: dict[str, dict] = {}
    for i, pair in enumerate(pairs):  # 設定の誤りは比べ始める前にまとめて止める（終了コード 2）
        if not isinstance(pair, dict) or not all(pair.get(k) for k in ("type", "before", "after")):
            raise common_cli.UsageError(f"global.yaml: pairs[{i}] に type・before・after が必要です")
        if pair["type"] not in confs:
            conf = _config(args.config_dir / f"{pair['type']}.yaml")
            try:
                excel_diff.validate_config(conf, f"{pair['type']}.yaml")
            except ValueError as e:
                raise common_cli.UsageError(str(e)) from None
            confs[pair["type"]] = conf
    out_dir = _resolve(base, str(g.get("output_dir", "output")), "output_dir", must_exist=False)
    out_dir.mkdir(parents=True, exist_ok=True)
    ext = EXTENSIONS[args.format]
    used: set[str] = set()
    findings: list[Finding] = []
    rc = common_cli.ExitCode.OK
    for pair in pairs:
        before = _resolve(base, str(pair["before"]), "before", must_exist=True)
        after = _resolve(base, str(pair["after"]), "after", must_exist=True)
        where = f"{before.as_posix()} → {after.as_posix()}"
        missing = [p for p in (before, after) if not p.is_file()]
        if missing:
            msg = "見つかりません: " + "、".join(p.as_posix() for p in missing)
            print(f"design-diff: {where}: {msg}（このペアは比べずに続けます）", file=sys.stderr)
            findings.append(Finding("pair_missing", Severity.ERROR, msg, location=where))
            rc = max(rc, common_cli.ExitCode.FINDINGS)
            continue
        try:
            result = excel_diff.run(confs[pair["type"]], before, after, outside=args.outside)
        except Exception as e:  # noqa: BLE001 — 1 ペアの失敗で全体を止めない
            print(f"design-diff: {where}: {type(e).__name__}: {e}（このペアは比べずに続けます）", file=sys.stderr)
            findings.append(Finding("pair_failed", Severity.ERROR, f"{type(e).__name__}: {e}", location=where))
            rc = max(rc, common_cli.ExitCode.FINDINGS)
            continue
        dest = _unique_dest(out_dir, after.stem, ext, used)
        dest.write_text(render.render(result, args.format, TOOL, __version__), encoding="utf-8", newline="\n")
        findings.append(
            Finding(
                "pair_compared",
                Severity.INFO,
                f"差分 {len(result.changes)} 件 → {dest.as_posix()}",
                location=where,
                actual=str(len(result.changes)),
            )
        )
        rc = max(rc, _exit(result, args.fail_on_diff))
    text = report.render(
        findings,
        args.format,
        title="design-diff 一括比較（--all）",
        meta={"pairs": len(pairs), "output_dir": out_dir.as_posix()},
        tool=TOOL,
        tool_version=__version__,
    )
    common_cli.write_output(text, args.output)
    return rc


def cmd_excel(args) -> int:
    if args.all:
        return cmd_all(args)
    if not (args.type and args.before and args.after):
        raise common_cli.UsageError("--type・--before・--after を指定するか、--all を使ってください")
    conf = _config(args.config_dir / f"{args.type}.yaml")
    try:
        excel_diff.validate_config(conf, f"{args.type}.yaml")
    except ValueError as e:
        raise common_cli.UsageError(str(e)) from None
    for f in (args.before, args.after):
        if not f.is_file():
            raise common_cli.UsageError(f"見つかりません: {f}")
    try:
        result = excel_diff.run(conf, args.before, args.after, outside=args.outside)
    except UnreadableBookError as e:
        raise common_cli.UsageError(str(e)) from None
    common_cli.write_output(render.render(result, args.format, TOOL, __version__), args.output)
    return _exit(result, args.fail_on_diff)


def cmd_spec(args) -> int:
    conf = _config(args.config or args.config_dir / "spec.yaml")
    for f in (args.excel, args.oas):
        if not f.is_file():
            raise common_cli.UsageError(f"見つかりません: {f}")
    try:
        result = spec_diff.run(conf, args.excel, args.oas)
    except UnreadableBookError as e:
        raise common_cli.UsageError(str(e)) from None
    except (KeyError, ValueError) as e:
        raise common_cli.UsageError(f"設定または OpenAPI の誤り: {e}") from None
    common_cli.write_output(render.render(result, args.format, TOOL, __version__), args.output)
    return _exit(result, args.fail_on_diff)


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    common_cli.configure_from_args(args)
    return cmd_excel(args) if args.command == "excel" else cmd_spec(args)


def entrypoint() -> None:
    sys.exit(common_cli.run_main(main))
