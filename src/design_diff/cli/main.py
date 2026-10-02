"""design-diff — 設計書のレビューで使う差分レポート.

  design-diff excel --type api --before 旧.xlsx --after 新.xlsx          # 設定は config/api.yaml
  design-diff excel --all                                                # config/global.yaml の pairs を一括
  design-diff spec --excel API定義書.xlsx --oas openapi.yaml             # 設定は config/spec.yaml

出力: 既定は Markdown（レビューにそのまま貼れる）。-f json で機械可読。-o でファイルへ。
終了コード: 0 正常（差分の有無にかかわらず）/ 1 --fail-on-diff で差分あり、または設定の見直しが必要 / 2 引数・設定の誤り / 3 実行時エラー
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from devtools_common import cli as common_cli
from devtools_common.config import load_structured

from .. import __version__
from ..application import excel_diff, render, spec_diff

TOOL = "design-diff"


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog=TOOL, description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    p.add_argument("--config-dir", type=Path, default=Path("config"), help="設定ファイルの置き場（既定: config/）")
    sub = p.add_subparsers(dest="command", required=True)

    e = sub.add_parser("excel", help="Excel 設計書の新旧差分")
    e.add_argument("--type", help="設定の種類（config/<type>.yaml）。例: api / batch / if")
    e.add_argument("--before", type=Path)
    e.add_argument("--after", type=Path)
    e.add_argument("--all", action="store_true", help="config/global.yaml の pairs をすべて比較し、output_dir に書き出す")
    e.add_argument("--fail-on-diff", action="store_true")
    common_cli.add_common_args(e)

    s = sub.add_parser("spec", help="API 定義書（Excel）と OpenAPI の差分")
    s.add_argument("--excel", type=Path, required=True)
    s.add_argument("--oas", type=Path, required=True)
    s.add_argument("--config", type=Path, help="既定: config/spec.yaml")
    s.add_argument("--fail-on-diff", action="store_true")
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


def cmd_excel(args) -> int:
    if args.all:
        g = _config(args.config_dir / "global.yaml")
        out_dir = Path(g.get("output_dir", "output"))
        out_dir.mkdir(parents=True, exist_ok=True)
        rc = common_cli.ExitCode.OK
        for pair in g.get("pairs") or []:
            conf = _config(args.config_dir / f"{pair['type']}.yaml")
            try:
                excel_diff.validate_config(conf, f"{pair['type']}.yaml")
            except ValueError as e:
                raise common_cli.UsageError(str(e)) from None
            result = excel_diff.run(conf, Path(pair["before"]), Path(pair["after"]))
            dest = out_dir / f"diff_{Path(pair['after']).stem}.md"
            dest.write_text(render.markdown(result), encoding="utf-8", newline="\n")
            print(f"{dest}: 差分 {len(result.changes)} 件")
            rc = max(rc, _exit(result, args.fail_on_diff))
        return rc
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
    result = excel_diff.run(conf, args.before, args.after)
    common_cli.write_output(render.render(result, args.format, TOOL, __version__), args.output)
    return _exit(result, args.fail_on_diff)


def cmd_spec(args) -> int:
    conf = _config(args.config or args.config_dir / "spec.yaml")
    for f in (args.excel, args.oas):
        if not f.is_file():
            raise common_cli.UsageError(f"見つかりません: {f}")
    try:
        result = spec_diff.run(conf, args.excel, args.oas)
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
