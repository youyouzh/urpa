import json
import os.path
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from typing import List

import smbclient
from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.cron import CronTrigger

from base.config import load_config
from base.log import logger

CONFIG = {
    # 邮箱服务地址
    'email_server_host': 'smtp.qq.com',
    'email_server_port': 465,

    # 使用哪个邮箱进行发送
    'from_email_address': 'xxx@qq.com',

    # 发送人名称
    'from_email_name': 'xxxx',

    # QQ邮箱的密钥
    'from_email_password': 'xxxx',

    # nas服务器，不包含"\\"
    'nas_server': r'dog',

    # nas存储访问用户名
    'nas_username': r'uusama',

    # nas存储访问密码
    'nas_password': 'uusama',

    # 任务配置路径
    'task_config_path': r'E:\xx\send_task.json',

    # 发送文件的nas跟目录
    'send_email_file_root_path': r'\\dog\xxx\gme',

    # 启动http服务端口
    'http_server_port': 8032
}
load_config(CONFIG)

# 注册nas存储的全局验证session
smbclient.register_session(CONFIG['nas_server'], CONFIG['nas_username'], CONFIG['nas_password'])
scheduler = BlockingScheduler()
LOAD_SEND_TASK_JOB_ID = 'LOAD_SEND_TASK_JOB'


# 任务配置，定义对象，代替dict更加直观
class TaskConfig(object):

    def __init__(self, config: dict):
        self.index = config.get('index', 1)
        # 任务调度表达式
        self.crontab = config.get('crontab')
        # 客户名称
        self.custom_name = config.get('custom_name')
        # 报表名称
        self.report_name = config.get('report_name')
        # 收件人邮箱地址
        self.email_address = config.get('email_address')
        # 收件人名称，不填时取邮箱地址
        self.email_name = config.get('email_name', self.email_address)
        # 邮件标题，不填时取文件名
        self.email_title = config.get('email_title', '')
        # 邮件内容，不填时没有内容
        self.email_content = config.get('email_content', '')

    def __str__(self):
        return json.dumps(self.__dict__, ensure_ascii=False)

    def __getitem__(self, item):
        return getattr(self, item)

    def __setitem__(self, key, value):
        setattr(self, key, value)


def list_nas_dir(dir_path, filter_dir=False):
    if not smbclient.path.isdir(dir_path):
        logger.error('The dir is not exist: {}'.format(dir_path))
        return []
    filepaths = []
    for file in smbclient.listdir(dir_path):
        full_path = os.path.join(dir_path, file)
        if filter_dir and smbclient.path.isdir(full_path):
            continue
        filepaths.append(full_path)
    return filepaths


# 登录qq邮箱服务器
def login_qq_email_server():
    """
    解决smtplib.SMTPServerDisconnected: Connection unexpectedly closed
    查看是否连接端口异常，比如腾讯企业邮箱为465，此时需要用 smtplib.SMTP_SSL(smtp_host) 连接，其中 smtp_host = ‘smtp.exmail.qq.com’。
    若端口为587，则需要 在 smtp.login() 前加上 smtp.starttls()。若端口为465，则不需要TLS连接，加上反而会报错。
    如果以上情况均正常，则考虑是否附件过大，因为像腾讯企业邮箱的普通附件大小限制为50M，超出即会报错。若出现此问题，需要对附件先分割再发送。
    :return:
    """
    try:
        server = smtplib.SMTP_SSL(CONFIG['email_server_host'], CONFIG['email_server_port'])
        login_result = server.login(CONFIG['from_email_address'], CONFIG['from_email_password'])
        logger.info('login result: {}'.format(login_result))
        return server
    except smtplib.SMTPAuthenticationError as exception:
        logger.error('login exception: '.format(exception))


def build_email_content(to_name, to_email, subject, content='') -> MIMEMultipart:
    # 1定义一个可以添加正文和附件的邮件消息对象
    message = MIMEMultipart()

    # 发件人和收件人名称，如果没设置则使用邮箱地址
    to_name = to_name if to_name else to_email
    from_name = CONFIG['from_email_name'] if CONFIG['from_email_name'] else CONFIG['from_email_address']
    # 构建发送人，接收人，邮件主题
    message['From'] = formataddr((from_name, CONFIG['from_email_address']))
    message['To'] = formataddr((to_name, to_email))
    message['subject'] = subject

    # 构建正文
    message.attach(MIMEText(content, 'plain', 'utf-8'))
    return message


# 在邮件中中添加附件
def add_email_attachment_from_nas(message: MIMEMultipart, filepath):
    if not smbclient.path.isfile(filepath):
        logger.error('The file is not exist: {}'.format(filepath))
        raise Exception('he file is not exist')
    with smbclient.open_file(filepath, mode="rb") as file_handler:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file_handler.read())
        encoders.encode_base64(part)
        filename = os.path.basename(filepath)
        part.add_header('Content-Disposition', 'attachment', filename=("utf-8", "", filename))
        message.attach(part)
        logger.info('add email attachment success: {}'.format(filepath))


def email_send_job(task_config: TaskConfig):
    logger.info('--->begin email send job. task_config: {}'.format(task_config))
    # 检查是否有对应文件
    send_file_dir = CONFIG['send_email_file_root_path']
    send_file_dir = os.path.join(send_file_dir, task_config.custom_name, task_config.report_name)
    if not smbclient.path.isdir(send_file_dir):
        logger.error('The email send file path is not exist: {}'.format(send_file_dir))
        return False

    # 检查该目录下的文件列表，过滤”归档“目录
    sub_files = list_nas_dir(send_file_dir, filter_dir=True)
    if not sub_files:
        logger.warning('There are not any send report file. dir: {}'.format(send_file_dir))
        return False

    email_content = build_email_content(task_config.email_name, task_config.email_address,
                                        task_config.email_title, task_config.email_content)

    # 检查 .ok 文件，表示已经准备好可以发送的文件
    ok_sub_files = []
    source_sub_files = []
    for sub_file in sub_files:
        if sub_file[-3:] != '.ok':
            continue
        source_sub_file = sub_file[:-3]
        # 检查ok文件对应的源文件是否存在
        if source_sub_file not in sub_files:
            logger.warning('The source file is not exist. ok file: {}'.format(sub_file))
            continue
        ok_sub_files.append(sub_file)
        source_sub_files.append(source_sub_file)
    logger.info('The dir: {}, ok sub files: {}'.format(send_file_dir, ok_sub_files))

    if not source_sub_files:
        logger.warning('There are not any valid ok source files. dir: {}'.format(send_file_dir))
        return False

    for sub_file in ok_sub_files:
        if not email_content['subject']:
            # 如果没有设置邮件标题，则取附件名称
            logger.info('add email attachment file: {}'.format(sub_file))
            email_content['subject'] = os.path.basename(sub_file)
        add_email_attachment_from_nas(email_content, sub_file)
    try:
        # 发送时登录，避免过一段时间登录连接中断需要重连
        qq_email_server = login_qq_email_server()
        send_result = qq_email_server.sendmail(CONFIG['from_email_address'], task_config.email_address,
                                               email_content.as_string())
        # 如果正确返回，则说明发送成功
        logger.info('send email finish. send_result: {}'.format(send_result))
    except Exception:
        logger.error('send email exception.', exc_info=True, stack_info=True)
        return False
    # 如果发送成功，则需要把文件移动到“归档”目录
    archive_dir = os.path.join(send_file_dir, '归档')
    if not smbclient.path.isdir(archive_dir):
        # 文件夹不存在则创建
        logger.info('create archive dir: {}'.format(archive_dir))
        smbclient.mkdir(archive_dir)
    # 将发送的文件移动到归档目录
    for sub_file in source_sub_files:
        move_target_path = os.path.join(archive_dir, os.path.basename(sub_file))
        smbclient.replace(sub_file, move_target_path)
        logger.info('move file from {} -> {}'.format(sub_file, move_target_path))
    for sub_file in ok_sub_files:
        move_target_path = os.path.join(archive_dir, os.path.basename(sub_file))
        smbclient.replace(sub_file, move_target_path)
        logger.info('move file from {} -> {}'.format(sub_file, move_target_path))
    return True


# 加载发送任务列表
def ready_email_send_jobs(task_configs: List[TaskConfig]):
    # 获取当前的所有任务
    jobs = scheduler.get_jobs()
    # 遍历并删除所有旧任务
    logger.info('remove old jobs. size: {}'.format(len(jobs)))
    for job in jobs:
        if job.id != LOAD_SEND_TASK_JOB_ID:
            logger.info('remove old job: {}'.format(job.id))
            scheduler.remove_job(job.id)

    # 加载新的邮件发送job
    for task_config in task_configs:
        scheduler.add_job(id=task_config.custom_name + str(task_config.index),
                          func=email_send_job,
                          trigger=CronTrigger.from_crontab(task_config.crontab),
                          args=[task_config])
    logger.info('ready email send jobs finished. size: {}'.format(len(task_configs)))


def load_task_config() -> List[TaskConfig]:
    # nas上的配置文件
    task_config_path = CONFIG['task_config_path']
    if not os.path.isfile(task_config_path):
        logger.error('The task config file is not exist: {}'.format(task_config_path))
        return []
    with open(task_config_path, 'r', encoding='utf-8') as file_handler:
        task_configs = json.load(file_handler)
        logger.info('load task config finish. size: {}, content: {}'.format(len(task_configs), task_configs))
        # 检查配置的参数是否有问题
        valid_task_configs = []
        check_keys = ['crontab', 'custom_name', 'report_name', 'email_address']
        index = 0
        for task_config in task_configs:
            valid_tag = True
            index += 1
            task_config['index'] = int(task_config.get('index', 1)) * 1000 + index
            for check_key in check_keys:
                if not task_config.get(check_key, ''):
                    logger.info('The value is not valid for key: {}, config: {}'.format(check_key, task_config))
                    valid_tag = False
                    break
            if valid_tag:
                valid_task_configs.append(TaskConfig(task_config))
        logger.info('load valid task config finish. size: {}, content: {}'
                    .format(len(valid_task_configs), valid_task_configs))
        return valid_task_configs


# 准备定时加载邮件发送任务的job
def ready_task_config_load_job(task_configs: list):
    scheduler.add_job(id=LOAD_SEND_TASK_JOB_ID,
                      func=ready_email_send_jobs,
                      trigger=CronTrigger.from_crontab('0 * * * *'),
                      args=[task_configs])


def test_send_email_job():
    task_config = {
        "index": 0,
        "send_frequency": "分钟",
        "send_time": "每分钟",
        "custom_name": "测试客户",
        "report_name": "测试报表",
        "email_name": "",
        "email_address": "xxx@qq.com",
        "crontab": "* * * * *"
    }
    task_config = TaskConfig(task_config)
    email_send_job(task_config)


# 立即执行发送任务
def run_email_send_jobs():
    task_configs = load_task_config()
    for task_config in task_configs:
        email_send_job(task_config)


def run_server():
    task_configs = load_task_config()
    # 首次手动加载
    ready_email_send_jobs(task_configs)
    ready_task_config_load_job(task_configs)
    scheduler.start()


# 打包： pyinstaller -F gme.py -p ../ --exclude-module pywin32 --exclude-module gevent --exclude-module flask
# pyinstaller gme.spec
if __name__ == '__main__':
    # test_send_email_job()
    # run_email_send_jobs()
    run_server()
