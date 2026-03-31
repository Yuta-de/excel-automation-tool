# ====================
# CLI（コマンド実行用）
# ====================

import argparse
from core.main_logic import main
from core.config_loader import load_config, get_project_root
from core.logger_setup import setup_logger

from core.config_validator import validate_config_file_exists, validate_config
from pathlib import Path

def run_cli():
    try:
        BASE_DIR = get_project_root()
        config_path = BASE_DIR / "config.ini"
        validate_config_file_exists(str(config_path))
        config = load_config(str(config_path))
        validate_config(config)
    except Exception as e:
        print(f"設定エラー：{e}")
        return

    setup_logger(config["LOG"]["log_file"])
    main(config)

def run_gui():
    from gui.gui_app import gui_main
    gui_main()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="売上レポート自動生成ツール")
    parser.add_argument("--gui", action="store_true", help="GUIモードで起動する")
    args = parser.parse_args()

    if args.gui:
        run_gui()
    else:
        run_cli()