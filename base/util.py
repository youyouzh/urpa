"""
工具封装，定义一些通用函数
"""
import datetime
import hashlib
import os
import random
import re
import sys
import time
from ctypes import sizeof, c_uint, c_long, c_int, c_bool, Structure

import pyautogui
import win32clipboard
import win32con

from base.config import CONFIG
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


def get_save_file_path(filename: str):
    # 处理特殊字符替换，windows不允许文件名出现这些特殊字符
    filename = re.sub(r"[\\/?*<>|\":]+", '-', filename)
    date_folder = datetime.datetime.now().strftime('%Y-%m-%d')
    save_dir = date_folder
    save_path = os.path.join(date_folder, filename)
    if os.path.exists(save_path):
        # 文件名相同，进行重命名，不检测md5去重了
        logger.info('The same filename is exist. path: {}'.format(save_path))
        filename = filename.replace('.', '{}-{}.'.format(str(time.time()), str(random.randint(1000, 9000))))
        save_path = os.path.join(save_dir, filename)
    return save_path


def get_screenshot_path():
    time_str = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
    save_dir = os.path.join('screenshot')
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)
    return os.path.join(save_dir, 'screenshot-{}.png'.format(time_str))


def get_screenshot():
    screenshot = pyautogui.screenshot()  # 截取全屏
    screenshot_path = get_screenshot_path()
    screenshot.save(screenshot_path)  # 保存截图为 PNG 文件
    return screenshot_path


def get_upload_absolute_path(filepaths: list):
    # 拼接完整的上传绝对路径
    absolute_paths = []
    for filepath in filepaths:
        absolute_path = os.path.join(CONFIG['upload_path'], filepath)
        absolute_path = os.path.abspath(absolute_path)
        absolute_paths.append(absolute_path)
    return absolute_paths


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


# Pyinstaller 可以将资源文件一起bundle到exe中，当exe在运行时，会生成一个临时文件夹，程序可通过sys._MEIPASS访问临时文件夹中的资源
# 编辑spec文件，datas选项包含一个元组[('resources', 'resources')]，第一个元素表示实际资源文件夹名称，第二个元素表示临时文件夹名称
def get_resource_path(filename=None, dir_tag='resources'):
    # 生成资源文件目录访问路径
    if getattr(sys, 'frozen', False):  # 是否Bundle Resource
        base_path = sys._MEIPASS  # 解包exe运行的临时目录
    else:
        base_path = os.path.dirname(__file__).replace(r'\base', '')
    base_path = os.path.join(base_path, dir_tag)
    if filename:
        base_path = os.path.join(base_path, filename)
    return base_path


def get_current_dir():
    if getattr(sys, 'frozen', False):  # 是否Bundle Resource
        # exe运行时获取当前文件夹
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        # 不能使用 __file__
        return os.path.dirname(os.path.abspath(sys.argv[0]))


# 使用命令行一次性打开多个微信客户端，需要没有已登录的微信
def open_multi_wechat(path, count):
    open_command = r'start "" "{}"'.format(path)
    full_command = open_command
    for index in range(count - 1):
        full_command += ' & ' + open_command
    os.system(full_command)
