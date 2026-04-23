#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Phase 1 実行スクリプト — CSV → SQLite（＋ Option B: JSON出力）

使用例:
  python phase1_main.py --csv-path data/input/sample.csv
  python phase1_main.py --csv-path data/input/sample.csv --export-json
  python phase1_main.py --csv-path data/input/sample.csv --db-path data/work/database.db --export-json
"""
import argparse
import os
import sys

# Python モジュール検索パスを設定（Embedded Python対応）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Phase1 は標準ライブラリのみ使用
from python.csv_to_sqlite import CSVToSQLite
from python.validator import ValidationError


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description='Phase 1: CSV → SQLite（外部パッケージ不要）'
    )
    parser.add_argument('--csv-path', required=True, help='入力CSVファイルパス')
    parser.add_argument(
        '--db-path',
        default='data/work/database.db',
        help='SQLite出力パス（デフォルト: data/work/database.db）',
    )
    parser.add_argument(
        '--export-json',
        action='store_true',
        help='VBA用 cases.json も出力する（Option B）',
    )
    parser.add_argument(
        '--json-path',
        default='data/work/cases.json',
        help='cases.json 出力パス（--export-json 時のみ有効）',
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    converter = CSVToSQLite(db_path=args.db_path)

    try:
        rows = converter.run(csv_path=args.csv_path)
    except ValidationError as e:
        print(f'✗ バリデーションエラー: {e}', file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f'✗ エラー: {e}', file=sys.stderr)
        sys.exit(1)

    print('✓ Phase1 完了')
    print(f'  - SQLiteDB : {os.path.abspath(args.db_path)}')
    print(f'  - ケース数  : {len(rows)} 件')

    if args.export_json:
        try:
            converter.export_json(rows=rows, json_path=args.json_path)
            print(f'  - cases.json: {os.path.abspath(args.json_path)}')
        except Exception as e:
            print(f'✗ JSON出力エラー: {e}', file=sys.stderr)
            sys.exit(1)


if __name__ == '__main__':
    main()
