"""
全局配置项
"""
import json
import os


def load_config(default_config: dict, config_json_path: str = 'config.json') -> dict:
    if os.path.isfile(config_json_path):
        with open(config_json_path, encoding='utf-8') as f:
            file_config = json.load(f)
            default_config.update(file_config)
    return default_config


def create_example_config(default_config: dict):
    with open(r'config.json.example', 'w', encoding='utf-8') as fp:
        json.dump(default_config, fp, ensure_ascii=False, indent=4)

