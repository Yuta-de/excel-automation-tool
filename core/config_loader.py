# ====================
# config.ini 読み込み
# ====================
# 読むだけにする

import configparser

def load_config(path : str = "config.ini") -> configparser.ConfigParser:
    config = configparser.ConfigParser()
    config.read(path, encoding="utf-8")
    return config