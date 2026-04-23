"""Phase 1 panel: CSV → SQLite."""
from __future__ import annotations

from PyQt5.QtWidgets import (
    QGroupBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from ui.widgets import FilePickerWidget


class Phase1Panel(QGroupBox):
    """Panel containing all Phase 1 controls."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__("Phase 1 — CSV → SQLite", parent)

        layout = QVBoxLayout(self)

        self.csv_picker = FilePickerWidget(
            "CSVファイル：",
            placeholder="まとめ用エクセルファイル_*.csv",
            file_filter="CSV Files (*.csv);;All Files (*)",
        )
        layout.addWidget(self.csv_picker)

        self.run_btn = QPushButton("Phase 1 実行")
        self.run_btn.setFixedHeight(36)
        layout.addWidget(self.run_btn)
