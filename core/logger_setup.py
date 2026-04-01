# ====================
# logging初期化
# ====================

import logging
from pathlib import Path

# デフォルトの設定にハンドラーを追加できるようにした関数
def setup_logger(log_file: str, extra_handlers: list[logging.Handler] | None = None) -> None:
    log_dir = Path(log_file).parent
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # loggingの二重登録を防止
    root_logger = logging.getLogger()
    if root_logger.handlers:
        root_logger.handlers.clear()
    
    handlers=[
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler()
    ]
    
    # ハンドラーを追加する
    if extra_handlers:
        handlers.extend(extra_handlers)


    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=handlers
    )