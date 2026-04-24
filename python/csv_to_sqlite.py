"""CSV → SQLite converter — standard library only (csv, sqlite3, json, re, datetime, os)."""
import csv
import json
import os
import sqlite3
from datetime import datetime

from python.validator import (
    NUMERIC_COLUMNS,
    ValidationError,
    normalise_case_id,
    validate_columns,
    validate_row,
)

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
    case_id    TEXT PRIMARY KEY,
    img_main   TEXT,
    img_graph1 TEXT,
    img_graph2 TEXT,
    FOREIGN KEY (case_id) REFERENCES cases(case_id)
);
"""


class CSVToSQLite:
    def __init__(self, db_path: str) -> None:
        self.db_path = db_path

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def run(self, csv_path: str) -> list:
        """Read CSV, validate, write to SQLite. Returns list of row dicts."""
        rows = self._read_csv(csv_path)
        self._write_sqlite(rows)
        return rows

    def export_json(self, rows: list, json_path: str) -> None:
        """Export case rows as JSON for VBA consumption (Option B)."""
        os.makedirs(os.path.dirname(os.path.abspath(json_path)), exist_ok=True)
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(rows, f, ensure_ascii=False, indent=2)

    # ------------------------------------------------------------------
    # CSV reading
    # ------------------------------------------------------------------

    def _read_csv(self, csv_path: str) -> list:
        if not os.path.isfile(csv_path):
            raise ValidationError(f'CSVファイルが見つかりません: {csv_path}')

        rows = []
        seen_ids: set = set()

        with open(csv_path, 'r', encoding='utf-8-sig', newline='') as f:
            reader = csv.DictReader(f)
            # Strip whitespace from column names
            if reader.fieldnames:
                reader.fieldnames = [c.strip() for c in reader.fieldnames]
            validate_columns(reader.fieldnames or [])

            for lineno, raw in enumerate(reader, start=2):
                # Strip values; skip fully empty rows
                row = {k.strip(): (str(v).strip() if v is not None else '')
                       for k, v in raw.items() if k}
                if not any(row.values()):
                    continue

                validate_row(row, lineno)

                case_id = normalise_case_id(row['ケース番号'])
                if case_id in seen_ids:
                    raise ValidationError(
                        f'行{lineno}: ケース番号の重複 → "{row["ケース番号"]}"'
                    )
                seen_ids.add(case_id)

                rows.append({
                    'case_id': case_id,
                    'config_file': row.get('設定ファイル名', ''),
                    'conditions': row.get('設定条件', '').replace('\r\n', '\n').replace('\r', '\n'),
                    'max_height': float(row['最大高さ']),
                    'dispersion_height': float(row['90%高さ']),
                    'gas_holdup': float(row['ガスホールドアップ']),
                })

        return rows

    # ------------------------------------------------------------------
    # SQLite writing
    # ------------------------------------------------------------------

    def _write_sqlite(self, rows: list) -> None:
        abs_db = os.path.abspath(self.db_path)
        os.makedirs(os.path.dirname(abs_db), exist_ok=True)

        con = sqlite3.connect(abs_db, timeout=30)
        try:
            con.execute(DDL_CASES)
            con.execute(DDL_IMAGE_MAPPINGS)
            con.commit()

            now = datetime.now().isoformat(sep=' ', timespec='seconds')
            for r in rows:
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
                    (r['case_id'], r['config_file'], r['conditions'],
                     r['max_height'], r['dispersion_height'], r['gas_holdup'], now),
                )
            con.commit()
        finally:
            con.close()
