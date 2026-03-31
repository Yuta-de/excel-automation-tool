# config.iniの内容を検証するためのロジックのみ

from pathlib import Path

REQUIRED_SECTIONS = {
    "PATH": ["input_folder", "template_path", "output_path"],
    "LOG": ["log_file"]
}

def validate_config_file_exists(path: str) -> None:
    if not Path(path).exists():
        raise FileNotFoundError(f"設定ファイルが見つかりません：{path}")
        
def validate_config(config) -> None :
    for section, keys in REQUIRED_SECTIONS.items():
        if section not in config:
            raise ValueError(f"必須セクションがありません：[{section}]")
        for key in keys:
            if key not in config[section]:
                raise ValueError(f"必須キーがありません：[{section}] {key}")
            if not config[section][key].strip():
                raise ValueError(f"値が空です：[{section}] {key}")