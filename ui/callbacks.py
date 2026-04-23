"""UI callbacks: bridge between button presses and core pipelines.

Each function accepts UI widgets as plain arguments and returns nothing;
it writes results back via the log widget.  No core logic lives here.
"""
from __future__ import annotations

import logging

from core.phase1_builder import Phase1Pipeline
from core.phase2_replace import Phase2Pipeline

logger = logging.getLogger(__name__)


def on_phase1_run(
    csv_path: str,
    config_path: str,
    log_widget,
) -> None:
    if not csv_path:
        log_widget.append_error("CSVファイルパスを指定してください。")
        return

    log_widget.append_info("Phase 1 を開始しています...")
    try:
        pipeline = Phase1Pipeline(config_path=config_path)
        result = pipeline.run(csv_path=csv_path)
        if result["status"] == "success":
            log_widget.append_success(result["message"])
        else:
            log_widget.append_error(result["message"])
    except Exception as e:
        log_widget.append_error(f"予期しないエラー: {e}")
        logger.exception("Phase 1 UI callback error")


def on_phase2_run(
    image_mapping_path: str,
    template_pptx_path: str,
    output_json_path: str,
    config_path: str,
    log_widget,
) -> None:
    missing = []
    if not image_mapping_path:
        missing.append("JSON画像マッピングパス")
    if not template_pptx_path:
        missing.append("テンプレートPPTXパス")
    if not output_json_path:
        missing.append("出力JSONパス")
    if missing:
        log_widget.append_error(f"以下の入力が未設定です: {', '.join(missing)}")
        return

    log_widget.append_info("Phase 2 を開始しています...")
    try:
        pipeline = Phase2Pipeline(config_path=config_path)
        result = pipeline.run(
            image_mapping_path=image_mapping_path,
            template_pptx_path=template_pptx_path,
            output_json_path=output_json_path,
        )
        if result["status"] == "success":
            log_widget.append_success(result["message"])
        else:
            log_widget.append_error(result["message"])
    except Exception as e:
        log_widget.append_error(f"予期しないエラー: {e}")
        logger.exception("Phase 2 UI callback error")
