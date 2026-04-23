"""Load and expose application configuration from settings.yaml."""
from __future__ import annotations

import os
import yaml


class ConfigManager:
    def __init__(self, config_path: str = "config/settings.yaml") -> None:
        self._config_path = os.path.abspath(config_path)
        self._base_dir = os.path.dirname(os.path.dirname(self._config_path))
        with open(self._config_path, encoding="utf-8") as f:
            self._cfg = yaml.safe_load(f)

    def _abs(self, rel: str) -> str:
        return os.path.join(self._base_dir, rel)

    @property
    def db_path(self) -> str:
        return self._abs(self._cfg["paths"]["db_path"])

    @property
    def input_dir(self) -> str:
        return self._abs(self._cfg["paths"]["input_dir"])

    @property
    def work_dir(self) -> str:
        return self._abs(self._cfg["paths"]["work_dir"])

    @property
    def output_dir(self) -> str:
        return self._abs(self._cfg["paths"]["output_dir"])

    @property
    def required_shapes(self) -> list[str]:
        return self._cfg["pptx"]["required_shapes"]

    @property
    def sqlite_timeout(self) -> int:
        return self._cfg["sqlite"]["timeout"]

    @property
    def logging_level(self) -> str:
        return self._cfg["logging"]["level"]

    @property
    def logging_format(self) -> str:
        return self._cfg["logging"]["format"]
