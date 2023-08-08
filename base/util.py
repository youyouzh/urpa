"""
工具封装，定义一些通用函数
"""
import hashlib
import json
import os
from ctypes import sizeof, c_uint, c_long, c_int, c_bool, Structure

import win32clipboard
import win32con

from base.log import logger


class DROPFILES(Structure):
    _fields_ = [
        ("pFiles", c_uint),
        ("x", c_long),
        ("y", c_long),
        ("fNC", c_int),
        ("fWide", c_bool),
    ]


pDropFiles = DROPFILES()
pDropFiles.pFiles = sizeof(DROPFILES)
pDropFiles.fWide = True
mate_data = bytes(pDropFiles)


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


# 将文件复制到剪贴板，参考：<https://blog.51cto.com/u_11866025/5833952>
def win32_clipboard_files(paths: list):
    files = ("\0".join(paths)).replace("/", "\\")
    data = files.encode("U16")[2:] + b"\0\0"
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_HDROP, mate_data + data)
    win32clipboard.CloseClipboard()
