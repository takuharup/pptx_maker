"""Base validation utilities shared across phases."""
from __future__ import annotations

import os
from typing import Any


class ValidationError(Exception):
    pass


def require_file(path: str, label: str = "File") -> str:
    """Resolve to absolute path and assert existence."""
    abs_path = os.path.abspath(path)
    if not os.path.isfile(abs_path):
        raise ValidationError(f"{label} not found: {abs_path}")
    return abs_path


def require_columns(df_columns: list[str], required: list[str]) -> None:
    missing = [c for c in required if c not in df_columns]
    if missing:
        raise ValidationError(f"CSV missing required columns: {missing}")


def coerce_float(value: Any, column: str) -> float:
    """Parse a value to float; raise ValidationError on failure."""
    try:
        return float(str(value).strip())
    except (ValueError, TypeError):
        raise ValidationError(
            f"Column '{column}' contains non-numeric value: {value!r}"
        )
