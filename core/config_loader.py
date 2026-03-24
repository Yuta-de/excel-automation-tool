# ====================
# config.ini 読み込み
# ====================

import configparser

def load_config(path="config.ini") -> configparser.ConfigParser:
    config = configparser.ConfigParser()
    config.read(path, encoding="utf-8")
    return config