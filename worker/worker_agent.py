"""
worker客户端代理，常驻服务，定时和服务端进行心跳检测，自定义一些接口实现
"""
import hmac
import json
import logging
import threading
from gevent.pywsgi import WSGIServer

from flask import Flask, jsonify, request, make_response
from flask.blueprints import Blueprint

from base.log import logger
from base.util import load_config
from components.wechat_app import WechatApp

CONFIG = load_config()


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
        return Response.response(400, **{"error": {"message": message, "code": 400}})

    @staticmethod
    def success():
        return jsonify({'error': {'code': 0, 'message': 'SUCCESS'}})


api = Blueprint("api", __name__)
wechat_app = WechatApp()


# 向Flask注册的API路由
class Api:

    @staticmethod
    @api.route("/api/bot/wechat/send-file", methods=['POST'])
    def wechat_send_file():
        to_conversations = request.values.get('toConversations')
        file_id = request.values.get('file_id')
        logger.info('send file. to_conversations: {}, file_id: {}'.format(to_conversations, file_id))
        wechat_app.send_file_task(to_conversations, file_id)
        return Response.success()

    @staticmethod
    @api.route("/api/bot/wechat/send-text", methods=['POST'])
    def wechat_send_text():
        to_conversations = request.values.get('toConversations')
        text = request.values.get('messageText')
        logger.info('send content. to_conversations: {}, content: {}'.format(to_conversations, text))
        wechat_app.send_text_task(to_conversations, text)
        return Response.success()


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
# 正式环境，使用 pyinstaller -F worker_agent.py 生成exe文件
if __name__ == "__main__":
    prod_run()
