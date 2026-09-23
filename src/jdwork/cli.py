"""Command line entry point for jdwork."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

from . import __version__
from .config import resolve_config_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="jdw", description="ERP 和店铺表格标准化工作流")
    parser.add_argument("--version", action="version", version=__version__)
    commands = parser.add_subparsers(dest="command", required=True)

    normalize = commands.add_parser("normalize", help="将 raw 表格标准化到 norm")
    normalize.add_argument("-c", "--config", help="配置文件路径，默认 config/norm.json")
    normalize.add_argument("--check", action="store_true", help="只检查配置和输入，不生成输出")

    filtering = commands.add_parser("filter", help="按规则筛选 norm 店铺表")
    filtering.add_argument("-c", "--config", help="配置文件路径，默认 config/filter.json")
    filtering.add_argument("--batch-id", help="14 位批次号，例如 20260923150000")
    filtering.add_argument("--check", action="store_true", help="只检查配置和标准化输入")

    backfill = commands.add_parser("backfill", help="将人工修改回填到 raw 店铺表")
    backfill.add_argument("-c", "--config", help="配置文件路径，默认 config/filter.json")
    backfill.add_argument("-i", "--input", required=True, help="filter/filter-{batch_id}.xlsx")
    backfill.add_argument("--dry-run", action="store_true", help="只预览变更，不保存 raw")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "normalize":
            from . import normalize

            config_path = resolve_config_path(args.config, "norm")
            normalize.check(config_path) if args.check else normalize.run(config_path)
        elif args.command == "filter":
            from . import filtering

            config_path = resolve_config_path(args.config, "filter")
            filtering.check(config_path) if args.check else filtering.run(config_path, args.batch_id)
        elif args.command == "backfill":
            from . import backfill

            config_path = resolve_config_path(args.config, "filter")
            backfill.run(config_path, Path(args.input).resolve(), args.dry_run)
        else:  # argparse enforces this; retained for type checkers.
            raise ValueError(f"未知命令: {args.command}")
        return 0
    except Exception as exc:
        print(f"失败: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
