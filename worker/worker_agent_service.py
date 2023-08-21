"""
worker客户端代理，常驻服务，定时和服务端进行心跳检测，自定义一些接口实现
"""
import socket

import servicemanager
import win32event
import win32service
import win32serviceutil

from base.log import logger
from worker.worker_agent import serve_forever


# windows系统服务
class WorkerAgentService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'urpa_worker_agent'
    _svc_display_name_ = 'Worker Agent For URPA'

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.is_alive = True

    def stop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False

    def run(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def main(self):
        # 在这里添加您的Python脚本代码
        logger.info('run service: {}'.format(self._svc_name_))
        try:
            serve_forever()
        except KeyboardInterrupt:
            # 手动触发结束，关闭浏览器等资源
            logger.error('KeyboardInterrupt-Main Thread end.')


# 按照业务域，通用账号，worker管理，心跳机制
# 正式环境，使用 `pyinstaller -F worker_agent_service.py -p ../ --hiddenimport win32timezone` 生成exe文件
# 使用命令 worker_agent_service install 安装服务
# 使用命令 worker_agent_service start 启动服务
# 使用命令 worker_agent_service stop 停止服务
if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(WorkerAgentService)
