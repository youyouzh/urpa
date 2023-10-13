import os
import sys
import logging
import time
from logging import handlers

ROTATION = 0
INFINITE = 1


logger = logging.getLogger('urpa')


# 日志等级过滤
class MaxLevelFilter(logging.Filter):
    """Filters (lets through) all messages with level < LEVEL"""
    def __init__(self, level):
        super().__init__()
        self.level = level

    def filter(self, record):
        # "<" instead of "<=": since logger.setLevel is inclusive, this should be exclusive
        return record.levelno < self.level


def get_log_path(name='log') -> str:
    # 初始化log
    log_path = os.path.join(os.getcwd(), 'log')
    if not os.path.isdir(log_path):
        os.makedirs(log_path)
    return os.path.join(log_path, name + time.strftime('-%Y-%m-%d-%H-%M-%S', time.localtime(time.time())) + '.log')


def config_file_logger(logging_instance, log_level, log_file, log_type=ROTATION,
                       max_size=1073741824, print_console=True, generate_wf_file=False):
    """
    config logging instance
    :param logging_instance: logger实例
    :param log_level: the log level
    :param log_file: log file path
    :param log_type:
        Two type: ROTATION and INFINITE

        log.ROTATION will let logfile switch to a new one (30 files at most).
        When logger reaches the 30th logfile, will overwrite from the
        oldest to the most recent.

        log.INFINITE will write on the logfile infinitely
    :param max_size: str max size
    :param print_console: Decide whether or not print to console
    :param generate_wf_file: Decide whether or not decide warning or fetal log file
    :return: none
    """
    # config object property
    logging_instance.setLevel(log_level)

    # '%(asctime)s - %(levelname)s - %(filename)s:%(lineno)s - %(message)s'
    # log format
    formatter = logging.Formatter(
        '%(levelname)s:\t %(asctime)s * '
        '[%(process)d:%(thread)x] [%(filename)s:%(lineno)s]\t %(message)s'
    )

    # print to console
    if print_console:
        # 避免默认basicConfig已经注册了root的StreamHandler，会重复输出日志，先移除掉
        for handler in logging.getLogger().handlers:
            if handler.name is None and isinstance(handler, logging.StreamHandler):
                logging.getLogger().removeHandler(handler)
        # DEBUG INFO 输出到 stdout
        stdout_handler = logging.StreamHandler(sys.stdout)
        stdout_handler.setFormatter(formatter)
        stdout_handler.setLevel(log_level)
        stdout_handler.addFilter(MaxLevelFilter(logging.WARNING))

        # WARNING 以上输出到 stderr
        stderr_handler = logging.StreamHandler(sys.stderr)
        stderr_handler.setFormatter(formatter)
        stderr_handler.setLevel(max(log_level, logging.WARNING))

        logging_instance.addHandler(stdout_handler)
        logging_instance.addHandler(stderr_handler)

    # set RotatingFileHandler
    if log_type == ROTATION:
        rf_handler = handlers.RotatingFileHandler(log_file, 'a', max_size, 30, encoding='utf-8')
    else:
        rf_handler = logging.FileHandler(log_file, 'a', encoding='utf-8')

    rf_handler.setFormatter(formatter)
    rf_handler.setLevel(log_level)

    # generate warning and fetal log to wf file
    if generate_wf_file:
        # add warning and fetal handler
        file_wf = str(log_file) + '.wf'
        warn_handler = logging.FileHandler(file_wf, 'a', encoding='utf-8')
        warn_handler.setLevel(logging.WARNING)
        warn_handler.setFormatter(formatter)
        logging_instance.addHandler(warn_handler)

    logging_instance.addHandler(rf_handler)


default_log_path = get_log_path()
config_file_logger(logger, logging.INFO, default_log_path, print_console=True)


# 读取日志最后N行
def get_last_n_logs(num: int):
    """
    读取大文件的最后几行
    :param num: 读取行数
    :return:
    """
    num = int(num)
    blk_size_max = 4096
    n_lines = []
    with open(default_log_path, 'rb') as fp:
        fp.seek(0, os.SEEK_END)
        cur_pos = fp.tell()
        while cur_pos > 0 and len(n_lines) < num:
            blk_size = min(blk_size_max, cur_pos)
            fp.seek(cur_pos - blk_size, os.SEEK_SET)
            blk_data = fp.read(blk_size)
            assert len(blk_data) == blk_size
            lines = blk_data.split(b'\n')

            # adjust cur_pos
            if len(lines) > 1 and len(lines[0]) > 0:
                n_lines[0:0] = lines[1:]
                cur_pos -= (blk_size - len(lines[0]))
            else:
                n_lines[0:0] = lines
                cur_pos -= blk_size
            fp.seek(cur_pos, os.SEEK_SET)

    if len(n_lines) > 0 and len(n_lines[-1]) == 0:
        del n_lines[-1]

    last_lines = []
    for line in n_lines[-num:]:
        last_lines.append(line.decode('utf-8'))
    return last_lines
