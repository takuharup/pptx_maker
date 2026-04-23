"""Validation primitives — standard library only."""
import re

REQUIRED_COLUMNS = [
    'ケース番号',
    '設定ファイル名',
    '設定条件',
    '最大高さ',
    '90%高さ',
    'ガスホールドアップ',
]

NUMERIC_COLUMNS = ('最大高さ', '90%高さ', 'ガスホールドアップ')


class ValidationError(Exception):
    pass


def validate_columns(fieldnames: list) -> None:
    missing = [c for c in REQUIRED_COLUMNS if c not in fieldnames]
    if missing:
        raise ValidationError(f'CSVに必須列がありません: {missing}')


def validate_row(row: dict, lineno: int) -> None:
    case_id = str(row.get('ケース番号', '')).strip()
    if not case_id:
        raise ValidationError(f'行{lineno}: "ケース番号" が空です')

    for col in NUMERIC_COLUMNS:
        val = str(row.get(col, '')).strip()
        if not val:
            raise ValidationError(f'行{lineno}: "{col}" が空です')
        try:
            float(val)
        except ValueError:
            raise ValidationError(f'行{lineno}: "{col}" が数値ではありません: {val!r}')


def normalise_case_id(raw: str) -> str:
    """Convert '1a', 'ケース1a', 'Case1A' → '1a'."""
    raw = str(raw).strip()
    m = re.search(r'(\d+[a-zA-Z]+|\d+)', raw)
    return m.group(1).lower() if m else raw.lower()
