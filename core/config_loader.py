# ====================
# config.ini 読み込み
# ====================
# 読むだけにする

import configparser
from pathlib import Path

def get_project_root() -> Path:
    return Path(__file__).resolve().parent.parent

def load_config(path : str = "config.ini") -> configparser.ConfigParser:
    config = configparser.ConfigParser()
    config.read(path, encoding="utf-8")
    return config