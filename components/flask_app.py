"""
flask http服务
"""
import json
import logging
from gevent import monkey

from multiprocessing import cpu_count, Process
from flask import Flask, jsonify, make_response
from flask.blueprints import Blueprint
from gevent.pywsgi import WSGIServer
from base.log import logger


# 多线程模式
monkey.patch_all()


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
        return jsonify({'code': 0, 'errorMessage': 'success', 'data': data})

    @staticmethod
    def fail(message):
        logger.info('response with fail. message: {}'.format(message))
        return jsonify({'code': 100, 'errorMessage': message})


api = Blueprint("api", __name__)
# flask app
app = Flask(__name__)
# 指定最大线程数
app.config['MAX_THREADS'] = 50
app.config['EXECUTOR_TYPE'] = 'thread'
app.config['EXECUTOR_MAX_WORKERS'] = 50


def server_forever(port: int = 8032):
    app.config['JSON_AS_ASCII'] = False
    app.config['SECRET_KEY'] = 'secret-key-u'
    app.register_blueprint(api)

    # bind http server
    # app.run(host='0.0.0.0', port=port)
    # 使用app.run开启的多线程虽然是单个进程下的多线程，但是采用了时间片轮询机制，当线程结束时会释放GIL锁
    # app.run使用flask内置的Web服务器，安全和效率不适合用在生产环境
    # 使用gevent多个协程绑定一个线程，增加高并发支持
    logger.info('begin gevent wsgi server. port: {}'.format(port))
    # access_logger = logging.getLogger('access')
    # config_file_logger(logger, logging.INFO, get_log_path('access'), print_console=False)
    wsgi_server = WSGIServer(('0.0.0.0', port), app,
                             log=logging.getLogger(),
                             error_log=logging.getLogger('error'))
    wsgi_server.serve_forever()


# 多进程模式
def server_forever_multi_process(port: int = 8032):
    multi_server = WSGIServer(('0.0.0.0', port), app)
    multi_server.start()

    def start_server_forever():
        multi_server.start_accepting()
        multi_server._stop_event.wait()

    for i in range(cpu_count()):
        p = Process(target=start_server_forever)
        p.start()


def run_server(server_port: int = 8032, multi_process: bool = False):
    """
    运行Flask HTTP服务
    """
    try:
        if multi_process:
            server_forever_multi_process(server_port)
        else:
            server_forever(server_port)
    except KeyboardInterrupt:
        # 手动触发结束，关闭浏览器等资源
        logger.error('KeyboardInterrupt-Main Thread end.')


if __name__ == "__main__":
    run_server(8032)
