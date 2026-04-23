"""Phase 2 panel: SQLite + JSON image mapping → VBA dataset."""
from __future__ import annotations

from PyQt5.QtWidgets import (
    QGroupBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from ui.widgets import FilePickerWidget


class Phase2Panel(QGroupBox):
    """Panel containing all Phase 2 controls."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__("Phase 2 — VBA実行データセット生成", parent)

        layout = QVBoxLayout(self)

        self.mapping_picker = FilePickerWidget(
            "JSON画像マッピング：",
            placeholder="image_mapping.json",
            file_filter="JSON Files (*.json);;All Files (*)",
        )
        layout.addWidget(self.mapping_picker)

        self.template_picker = FilePickerWidget(
            "テンプレートPPTX：",
            placeholder="template.pptx",
            file_filter="PowerPoint Files (*.pptx);;All Files (*)",
        )
        layout.addWidget(self.template_picker)

        self.output_picker = FilePickerWidget(
            "出力JSON：",
            placeholder="data/output/ppt_dataset.json",
            file_filter="JSON Files (*.json);;All Files (*)",
            save_mode=True,
        )
        layout.addWidget(self.output_picker)

        self.run_btn = QPushButton("Phase 2 実行")
        self.run_btn.setFixedHeight(36)
        layout.addWidget(self.run_btn)
