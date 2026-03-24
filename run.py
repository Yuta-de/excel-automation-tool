# ====================
# CLI（コマンド実行用）
# ====================

import argparse
from core.main_logic import main
from core.config_loader import load_config
from core.logger_setup import setup_logger

def run_cli():
    config = load_config()
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