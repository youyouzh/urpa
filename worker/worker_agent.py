"""
worker客户端代理，常驻服务，定时和服务端进行心跳检测，自定义一些接口实现
"""
import datetime
import json
import logging
import os

from flask import Flask, jsonify, request, make_response, send_file
from flask.blueprints import Blueprint
from gevent.pywsgi import WSGIServer

from base.config import CONFIG
from base.log import logger, get_last_n_logs
from base.util import MessageSendException, get_save_file_path, get_screenshot
from process.message_sender import MessageSenderManager, Message


class Response:
    @staticmethod
    def response(code, **kwargs):
        _ret_json = jsonify(kwargs)
        resp = make_response(_ret_json, code)
        resp.headers["Content-Type"] = "application/json; charset=utf-8"
        return resp

    @staticmethod
    def error_400(message: str = None):
        if not isinstance(message, str):
            message = json.dumps(message)
        logger.info(str(message))
        return Response.response(400, **{"errorMessage": message, "code": 400})

    @staticmethod
    def success(data=None):
        logger.info('response with success. data: {}'.format(data))
        return jsonify({'code': 0, 'errorMessage': 'SUCCESS', 'data': data})

    @staticmethod
    def fail(message):
        logger.info('response with fail. message: {}'.format(message))
        return jsonify({'code': 100, 'errorMessage': message})


api = Blueprint("api", __name__)
sender_manager = MessageSenderManager()


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
    def chat_sender_reload():
        senders = sender_manager.ready_message_senders()
        return Response.success('load success message sender size: {}'.format(len(senders)))

    # 聊天消息发送
    @staticmethod
    @api.route("/api/bot/wechat/send", methods=['POST'])   # 临时接口，ms-wechat那边写错了
    @api.route("/api/bot/chat/send", methods=['POST'])
    def chat_send_text():
        data = request.get_json()
        # json.loads(request.get_data())
        logger.info('receive send message request. params: {}'.format(json.dumps(data, ensure_ascii=False)))
        try:
            message = Message()
            message.init_from_json(data)
            sender_manager.send_message_with_exception(message)
        except MessageSendException as exception:
            logger.error('消息发送异常', stack_info=True, exc_info=True)
            return Response.fail(exception.message)
        except Exception as exception:
            logger.error('服务未知异常', stack_info=True, exc_info=True)
            return Response.fail('服务未知异常：{}'.format(exception))
        logger.info('send message success： {}'.format(message))
        return Response.success()

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


def serve_forever():
    # flask app
    app = Flask(__name__)
    app.config['JSON_AS_ASCII'] = False
    app.config['SECRET_KEY'] = 'secret'
    app.register_blueprint(api)

    # bind http server
    # app.run(host='0.0.0.0', port=port)
    # 使用app.run开启的多线程虽然是单个进程下的多线程，但是采用了时间片轮询机制，当线程结束时会释放GIL锁
    # app.run使用flask内置的Web服务器，安全和效率不适合用在生产环境
    # 使用gevent多个协程绑定一个线程，增加高并发支持
    logger.info('begin gevent wsgi server. port: {}'.format(CONFIG.get('http_server_port')))
    wsgi_server = WSGIServer(('0.0.0.0', CONFIG.get('http_server_port')), app,
                             log=logging.getLogger('access'),
                             error_log=logging.getLogger('error'))
    wsgi_server.serve_forever()


def prod_run():
    try:
        serve_forever()
    except KeyboardInterrupt:
        # 手动触发结束，关闭浏览器等资源
        logger.error('KeyboardInterrupt-Main Thread end.')


# 按照业务域，通用账号，worker管理，心跳机制
# 正式环境，使用 `pyinstaller -F worker_agent.py -p ../` 生成exe文件
if __name__ == "__main__":
    prod_run()
