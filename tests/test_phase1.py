"""Tests for Phase 1: CSV → SQLite (python/ modules)."""
import os
import json
import sqlite3
import tempfile
import textwrap

import pytest

from python.validator import (
    ValidationError,
    validate_columns,
    validate_row,
    normalise_case_id,
    REQUIRED_COLUMNS,
)
from python.csv_to_sqlite import CSVToSQLite

# ---------------------------------------------------------------------------
# Sample data
# ---------------------------------------------------------------------------
SAMPLE_CSV = textwrap.dedent("""\
    ケース番号,設定条件,追加条件,その他,設定ファイル名,最大高さ,90%高さ,ガスホールドアップ
    1a,flat底面,z=40 mm,,input_202604_case1a,0.125,0.088,0.718
    1b,flat底面,z=60 mm,,input_202604_case1b,0.138,0.096,0.742
    1c,curved底面,z=40 mm,,input_202604_case1c,0.119,0.082,0.705
""")


def _write_csv(content: str) -> str:
    fd, path = tempfile.mkstemp(suffix='.csv')
    with os.fdopen(fd, 'w', encoding='utf-8') as f:
        f.write(content)
    return path


# ---------------------------------------------------------------------------
# validator.py
# ---------------------------------------------------------------------------
class TestValidateColumns:
    def test_all_required_present(self):
        validate_columns(REQUIRED_COLUMNS)  # should not raise

    def test_missing_column_raises(self):
        cols = [c for c in REQUIRED_COLUMNS if c != '最大高さ']
        with pytest.raises(ValidationError, match='最大高さ'):
            validate_columns(cols)

    def test_extra_columns_ok(self):
        validate_columns(REQUIRED_COLUMNS + ['余分な列'])


class TestValidateRow:
    def _row(self, **overrides):
        base = {
            'ケース番号': '1a',
            '設定ファイル名': 'input_case1a',
            '設定条件': 'flat底面',
            '最大高さ': '0.125',
            '90%高さ': '0.088',
            'ガスホールドアップ': '0.718',
        }
        base.update(overrides)
        return base

    def test_valid_row(self):
        validate_row(self._row(), lineno=2)  # should not raise

    def test_empty_case_id_raises(self):
        with pytest.raises(ValidationError, match='ケース番号'):
            validate_row(self._row(**{'ケース番号': ''}), lineno=2)

    def test_non_numeric_raises(self):
        with pytest.raises(ValidationError, match='最大高さ'):
            validate_row(self._row(**{'最大高さ': 'ABC'}), lineno=5)

    def test_empty_numeric_raises(self):
        with pytest.raises(ValidationError, match='90%高さ'):
            validate_row(self._row(**{'90%高さ': ''}), lineno=3)


class TestNormaliseCaseId:
    def test_plain(self):
        assert normalise_case_id('1a') == '1a'

    def test_with_prefix(self):
        assert normalise_case_id('ケース1a') == '1a'

    def test_uppercase(self):
        assert normalise_case_id('Case1B') == '1b'

    def test_numeric_only(self):
        assert normalise_case_id('2') == '2'


# ---------------------------------------------------------------------------
# csv_to_sqlite.py
# ---------------------------------------------------------------------------
class TestCSVToSQLite:
    def test_run_creates_db(self, tmp_path):
        csv_path = tmp_path / 'sample.csv'
        csv_path.write_text(SAMPLE_CSV, encoding='utf-8')
        db_path = str(tmp_path / 'work' / 'database.db')

        converter = CSVToSQLite(db_path)
        rows = converter.run(str(csv_path))

        assert len(rows) == 3
        assert os.path.isfile(db_path)

    def test_db_contains_rows(self, tmp_path):
        csv_path = tmp_path / 'sample.csv'
        csv_path.write_text(SAMPLE_CSV, encoding='utf-8')
        db_path = str(tmp_path / 'work' / 'database.db')

        converter = CSVToSQLite(db_path)
        converter.run(str(csv_path))

        con = sqlite3.connect(db_path)
        rows = con.execute('SELECT case_id FROM cases ORDER BY case_id').fetchall()
        con.close()
        assert [r[0] for r in rows] == ['1a', '1b', '1c']

    def test_upsert_no_duplicate(self, tmp_path):
        csv_path = tmp_path / 'sample.csv'
        csv_path.write_text(SAMPLE_CSV, encoding='utf-8')
        db_path = str(tmp_path / 'work' / 'database.db')

        converter = CSVToSQLite(db_path)
        converter.run(str(csv_path))
        converter.run(str(csv_path))  # second run should upsert

        con = sqlite3.connect(db_path)
        count = con.execute('SELECT COUNT(*) FROM cases').fetchone()[0]
        con.close()
        assert count == 3

    def test_file_not_found_raises(self, tmp_path):
        converter = CSVToSQLite(str(tmp_path / 'database.db'))
        with pytest.raises(ValidationError, match='見つかりません'):
            converter.run('/no/such/file.csv')

    def test_missing_column_raises(self, tmp_path):
        bad_csv = textwrap.dedent("""\
            ケース番号,設定条件,設定ファイル名,最大高さ,ガスホールドアップ
            1a,flat底面,input_case1a,0.125,0.718
        """)
        csv_path = tmp_path / 'bad.csv'
        csv_path.write_text(bad_csv, encoding='utf-8')
        converter = CSVToSQLite(str(tmp_path / 'database.db'))
        with pytest.raises(ValidationError, match='必須列'):
            converter.run(str(csv_path))

    def test_duplicate_case_id_raises(self, tmp_path):
        dup_csv = textwrap.dedent("""\
            ケース番号,設定条件,追加条件,その他,設定ファイル名,最大高さ,90%高さ,ガスホールドアップ
            1a,flat底面,,,input_case1a,0.125,0.088,0.718
            1a,flat底面,,,input_case1a,0.125,0.088,0.718
        """)
        csv_path = tmp_path / 'dup.csv'
        csv_path.write_text(dup_csv, encoding='utf-8')
        converter = CSVToSQLite(str(tmp_path / 'database.db'))
        with pytest.raises(ValidationError, match='重複'):
            converter.run(str(csv_path))

    def test_export_json(self, tmp_path):
        csv_path = tmp_path / 'sample.csv'
        csv_path.write_text(SAMPLE_CSV, encoding='utf-8')
        db_path = str(tmp_path / 'work' / 'database.db')
        json_path = str(tmp_path / 'work' / 'cases.json')

        converter = CSVToSQLite(db_path)
        rows = converter.run(str(csv_path))
        converter.export_json(rows, json_path)

        assert os.path.isfile(json_path)
        with open(json_path, encoding='utf-8') as f:
            data = json.load(f)
        assert len(data) == 3
        assert data[0]['case_id'] == '1a'
        assert data[0]['max_height'] == pytest.approx(0.125)
