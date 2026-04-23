"""Tests for Phase 1: CSV → SQLite."""
import os
import sqlite3
import tempfile
import textwrap

import pytest

from core.phase1_builder import (
    CSVReader,
    DataValidator,
    Phase1Pipeline,
    SQLiteWriter,
)
from core.validator import ValidationError

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
SAMPLE_CSV_CONTENT = textwrap.dedent("""\
    ケース番号,設定条件,追加条件,その他,設定ファイル名,最大高さ,90%高さ,ガスホールドアップ
    1a,flat底面,z=40 mm,,input_202604_case1a,0.125,0.088,0.718
    1b,flat底面,z=60 mm,,input_202604_case1b,0.138,0.096,0.742
""")


def _write_csv(content: str, suffix=".csv") -> str:
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(content)
    return path


# ---------------------------------------------------------------------------
# CSVReader
# ---------------------------------------------------------------------------
class TestCSVReader:
    def test_read_valid(self):
        path = _write_csv(SAMPLE_CSV_CONTENT)
        try:
            df = CSVReader().read(path)
            assert len(df) == 2
            assert "ケース番号" in df.columns
        finally:
            os.unlink(path)

    def test_file_not_found(self):
        with pytest.raises(ValidationError, match="not found"):
            CSVReader().read("/nonexistent/path.csv")


# ---------------------------------------------------------------------------
# DataValidator
# ---------------------------------------------------------------------------
class TestDataValidator:
    def _df(self, content=SAMPLE_CSV_CONTENT):
        path = _write_csv(content)
        try:
            return CSVReader().read(path)
        finally:
            os.unlink(path)

    def test_valid_passes(self):
        DataValidator().validate(self._df())  # should not raise

    def test_missing_column(self):
        bad_csv = textwrap.dedent("""\
            ケース番号,設定条件,設定ファイル名,最大高さ,ガスホールドアップ
            1a,flat底面,input_202604_case1a,0.125,0.718
        """)
        with pytest.raises(ValidationError, match="missing required columns"):
            DataValidator().validate(self._df(bad_csv))

    def test_duplicate_case_id(self):
        dup_csv = textwrap.dedent("""\
            ケース番号,設定条件,追加条件,その他,設定ファイル名,最大高さ,90%高さ,ガスホールドアップ
            1a,flat底面,,,input_202604_case1a,0.125,0.088,0.718
            1a,flat底面,,,input_202604_case1a,0.125,0.088,0.718
        """)
        with pytest.raises(ValidationError, match="Duplicate"):
            DataValidator().validate(self._df(dup_csv))

    def test_non_numeric_raises(self):
        bad_csv = textwrap.dedent("""\
            ケース番号,設定条件,追加条件,その他,設定ファイル名,最大高さ,90%高さ,ガスホールドアップ
            1a,flat底面,,,input_202604_case1a,ABC,0.088,0.718
        """)
        with pytest.raises(ValidationError, match="non-numeric"):
            DataValidator().validate(self._df(bad_csv))


# ---------------------------------------------------------------------------
# SQLiteWriter
# ---------------------------------------------------------------------------
class TestSQLiteWriter:
    def test_write_and_read_back(self):
        csv_path = _write_csv(SAMPLE_CSV_CONTENT)
        try:
            df = CSVReader().read(csv_path)
            with tempfile.TemporaryDirectory() as tmp:
                db_path = os.path.join(tmp, "test.db")
                writer = SQLiteWriter(db_path)
                count = writer.write(df)
                assert count == 2
                con = sqlite3.connect(db_path)
                rows = con.execute("SELECT case_id FROM cases ORDER BY case_id").fetchall()
                con.close()
                assert [r[0] for r in rows] == ["1a", "1b"]
        finally:
            os.unlink(csv_path)

    def test_upsert_overwrites(self):
        csv_path = _write_csv(SAMPLE_CSV_CONTENT)
        try:
            df = CSVReader().read(csv_path)
            with tempfile.TemporaryDirectory() as tmp:
                db_path = os.path.join(tmp, "test.db")
                writer = SQLiteWriter(db_path)
                writer.write(df)
                writer.write(df)  # second write should upsert, not duplicate
                con = sqlite3.connect(db_path)
                count = con.execute("SELECT COUNT(*) FROM cases").fetchone()[0]
                con.close()
                assert count == 2
        finally:
            os.unlink(csv_path)


# ---------------------------------------------------------------------------
# Phase1Pipeline (integration)
# ---------------------------------------------------------------------------
class TestPhase1Pipeline:
    def test_full_run(self, tmp_path):
        csv_path = tmp_path / "sample.csv"
        csv_path.write_text(SAMPLE_CSV_CONTENT, encoding="utf-8")

        cfg_path = _make_config(tmp_path)
        pipeline = Phase1Pipeline(config_path=cfg_path)
        result = pipeline.run(csv_path=str(csv_path))

        assert result["status"] == "success"
        assert result["rows_written"] == 2
        assert os.path.isfile(result["db_path"])

    def test_run_returns_error_on_bad_csv(self, tmp_path):
        cfg_path = _make_config(tmp_path)
        pipeline = Phase1Pipeline(config_path=cfg_path)
        result = pipeline.run(csv_path="/no/such/file.csv")
        assert result["status"] == "error"


def _make_config(tmp_path) -> str:
    import yaml

    cfg = {
        "paths": {
            "db_path": str(tmp_path / "work" / "database.db"),
            "input_dir": str(tmp_path / "input"),
            "work_dir": str(tmp_path / "work"),
            "output_dir": str(tmp_path / "output"),
        },
        "sqlite": {"timeout": 30, "check_same_thread": False},
        "pptx": {
            "required_shapes": [
                "TXT_ConfigFile",
                "TXT_Conditions",
                "TXT_Dispersion",
                "TXT_GasHU",
                "IMG_Main",
                "IMG_Graph1",
                "IMG_Graph2",
            ]
        },
        "logging": {"level": "WARNING", "format": "%(message)s"},
    }
    cfg_path = tmp_path / "settings.yaml"
    cfg_path.write_text(yaml.dump(cfg), encoding="utf-8")
    return str(cfg_path)
