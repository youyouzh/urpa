"""
worker客户端代理，常驻服务，定时和服务端进行心跳检测，自定义一些接口实现
"""
import datetime
import json
import os

from components.flask_app import Response, api, server_forever
import requests
from flask import request, send_file

from base.config import load_config
from base.exception import ParamInvalidException, ControlInvalidException, MessageSendException
from base.log import logger, get_last_n_logs
from base.util import get_save_file_path, get_screenshot, get_localhost_ip
from apscheduler.schedulers.gevent import GeventScheduler
from process.message_sender import MessageSenderManager, Message, WechatTextMessageSender, WecomGroupBotMessageSender, \
    QQTextMessageSender

sender_manager = MessageSenderManager()


CONFIG = {
    # http服务端口
    "http_server_port": 8032,

    # http服务ip地址
    "http_server_host": get_localhost_ip(),

    # 上传文件保存路径
    "upload_path": "upload",

    # 返回失败时发送企微群通知
    "fail_notification_key": '7653c432-a6c7-424f-abf4-e64e3757b7e7',

    # ms-wechat服务列表
    "heartbeat_apis": [
        "http://10.191.0.37:9762/api/chat-message/worker-heartbeat",   # 测试

        # "http://10.201.5.45:10169/api/chat-message/worker-heartbeat",  # 生产
        # "http://10.201.5.46:10169/api/chat-message/worker-heartbeat",  # 生产
        # "http://10.202.5.45:10169/api/chat-message/worker-heartbeat",  # 生产
    ],
}
load_config(CONFIG)
error_message_sender = WecomGroupBotMessageSender()


# 错误消息通知
def error_notification(message):
    notification_key = CONFIG.get('fail_notification_key', '')
    if not notification_key:
        # 如果没有配置则不进行通知
        return False
    message = {
        "channel": "WECOM_GROUP_BOT",
        "fromSubject": notification_key,
        "messageType": "TEXT",
        "messageData": {
            "content": "【统一消息发送】消息发送失败请检查服务.错误信息：" + message,
            "mentionedList": ["zhaohai"]
        },
    }
    message = Message(message)
    try:
        error_message_sender.send(message)
    except:
        logger.error('error notification exception.', exc_info=True)
    return True


# 检查是否有强制更新对话框，登录状态等
def state_check():
    for message_sender in sender_manager.message_senders:
        # 自动跳过微信的强制更新
        if isinstance(message_sender, WechatTextMessageSender):
            message_sender.wechat_app.check_skip_update()


def heartbeat():
    logger.info('^^^^^^^heartbeat begin^^^^^^^^')
    json_data = {
        'host': CONFIG['http_server_host'],
        'port': CONFIG['http_server_port'],
        'supportSenders': []
    }
    unique_senders = []
    for message_sender in sender_manager.message_senders:
        if not message_sender.get_from_subject():
            # 如果发送主体为空，则不上报
            continue
        unique_sender = message_sender.get_channel() + message_sender.get_from_subject()
        if unique_sender not in unique_senders:
            json_data['supportSenders'].append({
                'fromSubject': message_sender.get_from_subject(),
                'channel': message_sender.get_channel(),
                'state': 'ACTIVE',
            })
            unique_senders.append(unique_sender)
    logger.info('request heartbeat api with data: {}'.format(json_data))
    for heartbeat_api in CONFIG['heartbeat_apis']:
        try:
            response = requests.post(heartbeat_api, json=json_data, timeout=10)
        except:
            logger.error('request heartbeat exception. heartbeat_api: {}'.format(heartbeat_api))
            continue

        if response.status_code != 200:
            logger.error('request heartbeat response status code is not 200. api: {}, code: {}, text: {}'
                         .format(heartbeat_api, response.status_code, response.text))
            continue
        logger.info('request heartbeat api success: {} response: {}'.format(heartbeat_api, response.text))
    logger.info('^^^^^^^heartbeat end^^^^^^^^')


def ready_scheduler():
    scheduler = GeventScheduler()
    # 10秒心跳一次
    # scheduler.add_job(id='heartbeat', func=heartbeat, trigger='cron', second='*/10')
    # 1分钟检查一次登录状态
    scheduler.add_job(id='state_check', func=state_check, trigger='cron', hour='*/4')
    scheduler.add_job(id='heartbeat', func=heartbeat, trigger='cron', hour='*/2')
    scheduler.start()


# 向Flask注册的API路由
class Api:

    # 查询基本状态
    @staticmethod
    @api.route("/api/state", methods=['GET'])
    def query_state():
        return Response.success({
            'time': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        })

    @staticmethod
    @api.route("/api/tail-log", methods=['GET'])
    def get_latest_log():
        last_n = request.values.get('n')
        last_n = int(last_n) if last_n else 100
        return '\n'.join(get_last_n_logs(last_n))

    @staticmethod
    @api.route('/api/screenshot', methods=['GET'])
    def screenshot():
        screenshot_path = get_screenshot()
        return send_file(screenshot_path, mimetype='image/png')  # 返回截图文件

    @staticmethod
    @api.route('/api/bot/chat/reload', methods=['GET'])
    def reload_chat_sender():
        sender_manager.ready_message_senders()
        heartbeat()
        return Response.success(sender_manager.senders_to_dict())

    # 封装一些常用的工具命令
    @staticmethod
    @api.route('/api/bot/chat/exec', methods=['GET'])
    def exec_command():
        command = request.values.get('command')
        sender_command = request.values.get('senderCommand')
        logger.info('执行命令 command: {}'.format(command))
        if command == 'heartbeat':
            heartbeat()
        elif command == 'getSenders':
            return Response.success(sender_manager.senders_to_dict())
        elif command == 'senderCommand':
            index = int(request.values.get('index'))
            if len(sender_manager.message_senders) < index:
                logger.warning('给定的index参数超过发送器数量：{}'.format(len(sender_manager.message_senders)))
                return Response.fail('给定的index参数超过发送器数量. index: {}'.format(index))
            match_sender = sender_manager.message_senders[index]
            logger.info('执行命令 command: {}, senderCommand: {}, channel: {}'.format(command, sender_command, index))
            try:
                match_sender.exec_command(sender_command)
            except Exception as e:
                logger.error('执行命令异常: {}'.format(sender_command), stack_info=True, exc_info=True)
                Response.fail('执行命令异常： {}'.format(e))
        return Response.success('exec command success: {}'.format(command))

    @staticmethod
    @api.route('/api/bot/chat/active', methods=['GET'])
    def active_wechat_window():
        index = request.values.get('index')
        index = int(index)
        count = 0
        for message_sender in sender_manager.message_senders:
            if isinstance(message_sender, WechatTextMessageSender):
                if index == count:
                    logger.info('置顶微信窗口: {}'.format(index))
                    message_sender.wechat_app.active(force=True)
                count += 1
        return Response.success()

    # 聊天消息发送
    @staticmethod
    @api.route("/api/bot/wechat/send", methods=['POST'])   # 临时接口，ms-wechat那边写错了
    @api.route("/api/bot/chat/send", methods=['POST'])
    def chat_send_text():
        data = request.get_json()
        # json.loads(request.get_data())
        logger.info('receive send message request. params: {}'.format(json.dumps(data, ensure_ascii=False)))
        try:
            message = Message(data)
            send_results = sender_manager.send_message_with_exception(message)
        except ParamInvalidException as exception:
            logger.error('消息发送参数检查异常', stack_info=True)
            return Response.fail(exception.message)
        except ControlInvalidException as exception:
            logger.error('消息发送元素定位异常', stack_info=True)
            error_notification(exception.message)
            return Response.fail(exception.message)
        except MessageSendException as exception:
            logger.error('消息发送异常', stack_info=True)
            error_notification(exception.message)
            return Response.fail(exception.message)
        except Exception as exception:
            logger.error('服务未知异常', stack_info=True)
            error_notification(exception)
            return Response.fail('服务未知异常：{}'.format(exception))
        logger.info('send message success. message: {}, send_result: {}'.format(message, send_results))
        return Response.success(send_results)

    @staticmethod
    @api.route('/api/file/upload', methods=['POST'])
    def upload_file():
        logger.info('receive upload file request.')
        if 'file' not in request.files:
            return Response.fail('the file param is empty.')

        file = request.files['file']
        if file.filename == '':
            return Response.fail('the filename upload failed, please retry.')

        save_path = get_save_file_path(file.filename)
        # 文件放到上传路径中
        full_save_path = os.path.join(CONFIG.get('upload_path'), save_path)
        if not os.path.isdir(os.path.dirname(full_save_path)):
            # 文件夹不存在则创建
            os.makedirs(os.path.dirname(full_save_path))
        file.save(full_save_path)
        logger.info('upload file success, save path: {}'.format(full_save_path))
        # 返回保存的路径
        return Response.success({'path': save_path})


# 按照业务域，通用账号，worker管理，心跳机制
# 正式环境，使用 `pyinstaller -F worker_agent.py -p ../` 生成exe文件
# pyinstaller worker_agent.spec
if __name__ == "__main__":
    try:
        heartbeat()
        ready_scheduler()
        server_forever(CONFIG.get('http_server_port'))
    except KeyboardInterrupt:
        # 手动触发结束，关闭浏览器等资源
        logger.error('KeyboardInterrupt-Main Thread end.')
