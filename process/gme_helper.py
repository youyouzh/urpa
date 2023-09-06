import datetime
import json

from flask_apscheduler import APScheduler

from base.config import create_example_config
from base.log import logger
from base.xlsx_util import read_excel_to_dict
from components.flask_app import serve_forever, Response, api, app
from process.gme import login_qq_email_server, build_email_content, CONFIG, email_send_job, TaskConfig, \
    load_task_config, ready_email_send_jobs, ready_task_config_load_job

scheduler = APScheduler()
scheduler.init_app(app)
create_example_config(CONFIG)


# 从xlsx表格中读取发送任务列表
def read_task_config_from_xlsx(xlsx_path: str) -> list:
    task_column_name_map = {
        '编号': 'index',
        '频率': 'send_frequency',
        '发送时间': 'send_time',
        '任务排期': 'crontab',
        '客户名称': 'custom_name',
        'GM报表名称': 'report_name',
        '邮箱': 'email',
    }
    send_tasks = read_excel_to_dict(xlsx_path, task_column_name_map)
    logger.info('send task size: {}'.format(len(send_tasks)))
    with open('send_task.json', 'w', encoding='utf-8') as handler:
        json.dump(send_tasks, handler, ensure_ascii=False, indent=4)
    return send_tasks


# 向Flask注册的API路由
class Api:

    # 查询基本状态
    @staticmethod
    @api.route("/api/jobs", methods=['GET'])
    def query_state():
        return Response.success({
            'time': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        })

    @staticmethod
    @api.route('/api/jobs', methods=['GET'])
    def query_jobs():
        jobs = scheduler.get_jobs()
        return Response.success(jobs)


def test_send_email():
    server = login_qq_email_server()
    to_email_address = '1406558940@qq.com'
    email_content = build_email_content('悠悠', to_email_address, '测试主题', '测试内容')
    send_result = server.sendmail(CONFIG['from_email_address'], to_email_address, email_content.as_string())
    logger.info('send_result: {}'.format(send_result))
    server.quit()


def test_send_email_job():
    task_config = {
        "index": 0,
        "send_frequency": "分钟",
        "send_time": "每分钟",
        "custom_name": "测试客户",
        "report_name": "测试报表",
        "email_name": "",
        "email_address": "1406558940@qq.com",
        "crontab": "* * * * *"
    }
    task_config = TaskConfig(task_config)
    email_send_job(task_config)


def run_server():
    task_configs = load_task_config()
    # 首次手动加载
    ready_email_send_jobs(task_configs)
    ready_task_config_load_job(task_configs)
    scheduler.start()
    serve_forever()


# 打包： pyinstaller -F gme.py -p ../
if __name__ == '__main__':
    # read_task_config_from_xlsx(r'E:\需求相关资料\需求-GM邮件发送\GM清单 的副本.xlsx')
    # test_send_email_job()
    run_server()
