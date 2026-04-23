"""Phase 1: CSV → SQLite.

Classes
-------
CSVReader        – Read and normalise the input CSV.
DataValidator    – Validate schema and data types.
SQLiteWriter     – Write/upsert rows into the SQLite database.
Phase1Pipeline   – Top-level orchestrator (called from UI, cmd, or tests).
"""
from __future__ import annotations

import logging
import os
import re
import sqlite3
from datetime import datetime
from typing import Any

import pandas as pd

from core.config_manager import ConfigManager
from core.validator import ValidationError, coerce_float, require_columns, require_file

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Column name constants (match the CSV header)
# ---------------------------------------------------------------------------
COL_CASE_ID = "ケース番号"
COL_CONFIG_FILE = "設定ファイル名"
COL_CONDITIONS = "設定条件"
COL_ADDITIONAL = "追加条件"
COL_OTHER = "その他"
COL_MAX_HEIGHT = "最大高さ"
COL_DISP_HEIGHT = "90%高さ"
COL_GAS_HOLDUP = "ガスホールドアップ"

REQUIRED_COLUMNS = [
    COL_CASE_ID,
    COL_CONFIG_FILE,
    COL_CONDITIONS,
    COL_MAX_HEIGHT,
    COL_DISP_HEIGHT,
    COL_GAS_HOLDUP,
]

DDL_CASES = """
CREATE TABLE IF NOT EXISTS cases (
    case_id           TEXT PRIMARY KEY,
    config_file       TEXT,
    conditions        TEXT,
    max_height        REAL,
    dispersion_height REAL,
    gas_holdup        REAL,
    created_at        TEXT
);
"""

DDL_IMAGE_MAPPINGS = """
CREATE TABLE IF NOT EXISTS image_mappings (
    case_id   TEXT PRIMARY KEY,
    img_main  TEXT,
    img_graph1 TEXT,
    img_graph2 TEXT,
    FOREIGN KEY (case_id) REFERENCES cases(case_id)
);
"""


# ---------------------------------------------------------------------------
# CSVReader
# ---------------------------------------------------------------------------
class CSVReader:
    """Read a CSV file and return a normalised DataFrame."""

    def read(self, csv_path: str) -> pd.DataFrame:
        abs_path = require_file(csv_path, "CSV")
        logger.info("Reading CSV: %s", abs_path)
        df = pd.read_csv(abs_path, dtype=str, encoding="utf-8-sig")
        df.columns = [c.strip() for c in df.columns]
        # Drop fully-empty rows
        df = df.dropna(how="all").reset_index(drop=True)
        # Strip whitespace from all string cells
        df = df.apply(lambda col: col.str.strip() if col.dtype == object else col)
        return df


# ---------------------------------------------------------------------------
# DataValidator
# ---------------------------------------------------------------------------
class DataValidator:
    """Validate a normalised DataFrame before DB write."""

    def validate(self, df: pd.DataFrame) -> None:
        require_columns(list(df.columns), REQUIRED_COLUMNS)
        self._check_duplicates(df)
        self._check_numerics(df)

    def _check_duplicates(self, df: pd.DataFrame) -> None:
        dupes = df[df.duplicated(subset=[COL_CASE_ID], keep=False)][COL_CASE_ID].unique()
        if len(dupes):
            raise ValidationError(f"Duplicate case IDs found: {list(dupes)}")

    def _check_numerics(self, df: pd.DataFrame) -> None:
        for col in (COL_MAX_HEIGHT, COL_DISP_HEIGHT, COL_GAS_HOLDUP):
            for _, val in df[col].items():
                if pd.isna(val) or str(val).strip() == "":
                    continue
                coerce_float(val, col)


# ---------------------------------------------------------------------------
# SQLiteWriter
# ---------------------------------------------------------------------------
class SQLiteWriter:
    """Create / update the SQLite database."""

    def __init__(self, db_path: str, timeout: int = 30) -> None:
        os.makedirs(os.path.dirname(db_path), exist_ok=True)
        self._db_path = db_path
        self._timeout = timeout

    def write(self, df: pd.DataFrame) -> int:
        """Upsert rows from *df* into the cases table. Returns row count."""
        con = sqlite3.connect(self._db_path, timeout=self._timeout)
        try:
            con.execute(DDL_CASES)
            con.execute(DDL_IMAGE_MAPPINGS)
            con.commit()

            rows_written = 0
            now = datetime.now().isoformat(sep=" ", timespec="seconds")
            for _, row in df.iterrows():
                case_id = self._normalise_case_id(str(row[COL_CASE_ID]))
                if not case_id:
                    continue
                conditions = self._normalise_text(row.get(COL_CONDITIONS, ""))
                max_h = self._to_float_or_none(row.get(COL_MAX_HEIGHT))
                disp_h = self._to_float_or_none(row.get(COL_DISP_HEIGHT))
                gas_hu = self._to_float_or_none(row.get(COL_GAS_HOLDUP))
                config_file = str(row.get(COL_CONFIG_FILE, "")).strip()

                con.execute(
                    """
                    INSERT INTO cases
                        (case_id, config_file, conditions, max_height,
                         dispersion_height, gas_holdup, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(case_id) DO UPDATE SET
                        config_file       = excluded.config_file,
                        conditions        = excluded.conditions,
                        max_height        = excluded.max_height,
                        dispersion_height = excluded.dispersion_height,
                        gas_holdup        = excluded.gas_holdup,
                        created_at        = excluded.created_at
                    """,
                    (case_id, config_file, conditions, max_h, disp_h, gas_hu, now),
                )
                rows_written += 1
            con.commit()
            logger.info("Wrote %d rows to %s", rows_written, self._db_path)
            return rows_written
        finally:
            con.close()

    @staticmethod
    def _normalise_case_id(raw: str) -> str:
        """Convert '1a', 'ケース1a', 'case1a' → '1a'."""
        raw = raw.strip()
        m = re.search(r"(\d+[a-zA-Z]+|\d+)", raw)
        return m.group(1).lower() if m else ""

    @staticmethod
    def _normalise_text(raw: Any) -> str:
        if pd.isna(raw):
            return ""
        text = str(raw).strip()
        # Unify line endings
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        return text

    @staticmethod
    def _to_float_or_none(raw: Any) -> float | None:
        if pd.isna(raw) or str(raw).strip() == "":
            return None
        try:
            return float(str(raw).strip())
        except ValueError:
            return None


# ---------------------------------------------------------------------------
# Phase1Pipeline
# ---------------------------------------------------------------------------
class Phase1Pipeline:
    """Orchestrate Phase 1: CSV → SQLite."""

    def __init__(self, config_path: str = "config/settings.yaml") -> None:
        self._cfg = ConfigManager(config_path)
        self._reader = CSVReader()
        self._validator = DataValidator()
        self._writer = SQLiteWriter(self._cfg.db_path, self._cfg.sqlite_timeout)

    def run(self, csv_path: str) -> dict:
        """
        Execute Phase 1.

        Returns
        -------
        dict with keys: status, message, db_path, rows_written
        """
        logger.info("=== Phase 1 start ===")
        try:
            df = self._reader.read(csv_path)
            self._validator.validate(df)
            rows = self._writer.write(df)
            msg = f"Phase 1 完了: {rows} 件を {self._cfg.db_path} に書き込みました。"
            logger.info(msg)
            return {
                "status": "success",
                "message": msg,
                "db_path": self._cfg.db_path,
                "rows_written": rows,
            }
        except ValidationError as e:
            logger.error("Validation error: %s", e)
            return {"status": "error", "message": f"バリデーションエラー: {e}"}
        except Exception as e:
            logger.exception("Unexpected error in Phase 1")
            return {"status": "error", "message": f"予期しないエラー: {e}"}
