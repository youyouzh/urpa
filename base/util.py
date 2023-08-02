"""
工具封装，定义一些通用函数
"""
import hashlib
import json
import os
import win32clipboard
import win32con

from base.log import logger


def report_error(message: str):
    # 错误消息上报，可以接入http
    logger.error(message)


def log_info(message: str):
    logger.info(message)


def log_warning(message: str):
    logger.warning(message)


def log_error(message: str):
    logger.error(message)


# 字符串hash
def md5_encrypt(content: str):
    md5_hash = hashlib.md5()
    if content:
        content = content.encode("utf8")
    md5_hash.update(content)
    return md5_hash.hexdigest()


# 将配置加载到config中
def load_config(config=None):
    # 加载文件配置表并合并
    if config is None:
        config = {}

    config_json_path = 'config.json'
    if os.path.isfile(config_json_path):
        with open(config_json_path, encoding='utf-8') as f:
            file_config = json.load(f)
            config.update(file_config)
    return config


def win32_clipboard_text(text: str):
    # 复制文本内容到剪贴板
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
    win32clipboard.CloseClipboard()


def win32_clipboard_dict(content_dict: dict):
    # 复制二进制kv
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    for k, v in content_dict.items():
        win32clipboard.SetClipboardData(int(k), v)
    win32clipboard.CloseClipboard()
