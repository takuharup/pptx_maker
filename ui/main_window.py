"""Main application window."""
from __future__ import annotations

import os

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QLabel,
    QMainWindow,
    QSplitter,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

from ui.callbacks import on_phase1_run, on_phase2_run
from ui.phase1_panel import Phase1Panel
from ui.phase2_panel import Phase2Panel
from ui.widgets import LogWidget

_DEFAULT_CONFIG = "config/settings.yaml"
_DEFAULT_OUTPUT = os.path.join("data", "output", "ppt_dataset.json")


class MainWindow(QMainWindow):
    def __init__(self, config_path: str = _DEFAULT_CONFIG) -> None:
        super().__init__()
        self._config_path = config_path
        self.setWindowTitle("PPTX自動生成ツール")
        self.resize(800, 600)
        self._build_ui()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        title = QLabel("PPTX 自動生成ツール")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 16pt; font-weight: bold; padding: 4px;")
        root.addWidget(title)

        splitter = QSplitter(Qt.Vertical)
        root.addWidget(splitter, stretch=1)

        # --- top: controls ---
        controls = QWidget()
        ctrl_layout = QVBoxLayout(controls)
        ctrl_layout.setContentsMargins(0, 0, 0, 0)

        self._phase1 = Phase1Panel()
        ctrl_layout.addWidget(self._phase1)

        self._phase2 = Phase2Panel()
        # Pre-fill default output path
        self._phase2.output_picker.setText(
            os.path.abspath(_DEFAULT_OUTPUT)
        )
        ctrl_layout.addWidget(self._phase2)

        splitter.addWidget(controls)

        # --- bottom: log ---
        self._log = LogWidget()
        splitter.addWidget(self._log)
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 1)

        # Status bar
        self.setStatusBar(QStatusBar())

        # Wire buttons
        self._phase1.run_btn.clicked.connect(self._run_phase1)
        self._phase2.run_btn.clicked.connect(self._run_phase2)

    def _run_phase1(self) -> None:
        on_phase1_run(
            csv_path=self._phase1.csv_picker.text(),
            config_path=self._config_path,
            log_widget=self._log,
        )

    def _run_phase2(self) -> None:
        on_phase2_run(
            image_mapping_path=self._phase2.mapping_picker.text(),
            template_pptx_path=self._phase2.template_picker.text(),
            output_json_path=self._phase2.output_picker.text(),
            config_path=self._config_path,
            log_widget=self._log,
        )
