"""Non-UI (CLI) entry point.

Usage
-----
# Phase 1
python main.py phase1 --csv data/input/sample_data/sample.csv

# Phase 2
python main.py phase2 \\
    --mapping  data/work/image_mapping.json \\
    --template data/input/template.pptx \\
    --output   data/output/ppt_dataset.json
"""
from __future__ import annotations

import argparse
import logging
import sys

from core.config_manager import ConfigManager
from core.phase1_builder import Phase1Pipeline
from core.phase2_replace import Phase2Pipeline

CONFIG_PATH = "config/settings.yaml"


def _setup_logging(cfg: ConfigManager) -> None:
    logging.basicConfig(level=cfg.logging_level, format=cfg.logging_format)


def cmd_phase1(args: argparse.Namespace) -> int:
    pipeline = Phase1Pipeline(config_path=args.config)
    result = pipeline.run(csv_path=args.csv)
    print(result["message"])
    return 0 if result["status"] == "success" else 1


def cmd_phase2(args: argparse.Namespace) -> int:
    pipeline = Phase2Pipeline(config_path=args.config)
    result = pipeline.run(
        image_mapping_path=args.mapping,
        template_pptx_path=args.template,
        output_json_path=args.output,
    )
    print(result["message"])
    return 0 if result["status"] == "success" else 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="PPTX自動生成ツール CLI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--config", default=CONFIG_PATH, help="設定ファイルパス")
    sub = parser.add_subparsers(dest="phase", required=True)

    p1 = sub.add_parser("phase1", help="CSV → SQLite")
    p1.add_argument("--csv", required=True, help="入力CSVファイルパス")

    p2 = sub.add_parser("phase2", help="SQLite + JSON → VBA実行データセット")
    p2.add_argument("--mapping", required=True, help="JSON画像マッピングパス")
    p2.add_argument("--template", required=True, help="テンプレートPPTXパス")
    p2.add_argument(
        "--output",
        default="data/output/ppt_dataset.json",
        help="出力JSONパス",
    )

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    cfg = ConfigManager(args.config)
    _setup_logging(cfg)

    if args.phase == "phase1":
        sys.exit(cmd_phase1(args))
    elif args.phase == "phase2":
        sys.exit(cmd_phase2(args))


if __name__ == "__main__":
    main()
