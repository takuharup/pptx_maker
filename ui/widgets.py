"""Reusable custom widgets."""
from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTextEdit,
    QWidget,
)


class FilePickerWidget(QWidget):
    """A label + line-edit + browse button for picking a file path."""

    def __init__(
        self,
        label: str,
        placeholder: str = "",
        file_filter: str = "All Files (*)",
        save_mode: bool = False,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._filter = file_filter
        self._save_mode = save_mode

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        lbl = QLabel(label)
        lbl.setFixedWidth(180)
        layout.addWidget(lbl)

        self._edit = QLineEdit()
        self._edit.setPlaceholderText(placeholder)
        layout.addWidget(self._edit, stretch=1)

        btn = QPushButton("参照…")
        btn.setFixedWidth(72)
        btn.clicked.connect(self._browse)
        layout.addWidget(btn)

    def _browse(self) -> None:
        if self._save_mode:
            path, _ = QFileDialog.getSaveFileName(self, "保存先を選択", "", self._filter)
        else:
            path, _ = QFileDialog.getOpenFileName(self, "ファイルを選択", "", self._filter)
        if path:
            self._edit.setText(path)

    def text(self) -> str:
        return self._edit.text().strip()

    def setText(self, text: str) -> None:
        self._edit.setText(text)


class LogWidget(QTextEdit):
    """Scrolling read-only log output area."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setReadOnly(True)
        self.setLineWrapMode(QTextEdit.WidgetWidth)

    def append_info(self, msg: str) -> None:
        self.append(f'<span style="color:#000;">{msg}</span>')
        self._scroll_bottom()

    def append_success(self, msg: str) -> None:
        self.append(f'<span style="color:#1a7a00;">{msg}</span>')
        self._scroll_bottom()

    def append_error(self, msg: str) -> None:
        self.append(f'<span style="color:#cc0000;">{msg}</span>')
        self._scroll_bottom()

    def _scroll_bottom(self) -> None:
        sb = self.verticalScrollBar()
        sb.setValue(sb.maximum())
