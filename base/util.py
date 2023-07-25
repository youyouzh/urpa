"""
工具封装，定义一些通用函数
"""
import hashlib

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
