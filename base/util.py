"""
工具封装，定义一些通用函数
"""
import hashlib
import json
import os
import win32clipboard
import win32con

from base.log import logger


class MessageSendException(Exception):

    def __init__(self, message: str, extract_data=None):
        logger.error(message + 'extract data: {}'.format(extract_data))
        self.message = message


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
