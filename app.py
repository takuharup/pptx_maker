"""UI entry point.  Run this to launch the PyQt5 GUI."""
import logging
import sys

from PyQt5.QtWidgets import QApplication

from core.config_manager import ConfigManager
from ui.main_window import MainWindow

CONFIG_PATH = "config/settings.yaml"


def _setup_logging(cfg: ConfigManager) -> None:
    logging.basicConfig(level=cfg.logging_level, format=cfg.logging_format)


def main() -> None:
    cfg = ConfigManager(CONFIG_PATH)
    _setup_logging(cfg)

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow(config_path=CONFIG_PATH)
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
