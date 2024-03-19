"""
微信、企业微信、QQ消息发送流程处理
"""
import json
import os
import time
from typing import List

import cptools
import threading

from base.control_util import active_window
from base.exception import ParamInvalidException, MessageSendException
from components.qq_app import QQApp
from corpwechatbot import CorpWechatBot, AppMsgSender

from base.log import logger
from base.config import load_config
from base.util import open_multi_app
from components.wechat_app import WechatApp

SCREEN_LOCK = threading.Lock()


CONFIG = {
    # 需要启动的微信客户端数量
    "start_wechat_app_count": 3,

    # 需要启动的QQ客户端数量
    "start_qq_app_count": 3,

    # 是否需要启用企微群机器人发送
    "start_wecom_group_bot": True,

    # 是否需要启动企微应用推送
    "start_wecom_app_push": True,

    # 上传文件路径
    "upload_path": "upload",

    # 是否保持会话窗口不关闭
    "qq_is_retain_conversation_window": True,

    # 是否启用在会话窗口中搜索，会更快一些单可能会有问题
    "qq_is_search_in_conversation_window": False,

    # 微信安装目录
    "wechat_path": r"D:\Program Files (x86)\Tencent\WeChat\WeChat.exe",

    # 微信安装目录
    "qq_path": r"D:\Program Files (x86)\Tencent\WeChat\QQ.exe",

    # 企微应用id，这个测试id是自建的企业，需要测试可以找维护人员添加
    "wecom_corp_id": "wwb36926885246afd4",

    # 企微应用key
    "wecom_corp_key": "aHUghPqj3Fd-m4sVuc7Qv9zqo8uhlEl4mgXed5Na-5U",

    # 企微消息推送应用id
    "wecom_agent_id": "3010041",

    # 推送应用名称
    "wecom_corp_name": "RPA助手"
}
load_config(CONFIG)


def check_param(value, message, extra_data=None):
    if not value:
        raise ParamInvalidException(message, extra_data)


def get_upload_absolute_path(filepaths: list):
    # 拼接完整的上传绝对路径
    absolute_paths = []
    for filepath in filepaths:
        absolute_path = os.path.join(CONFIG['upload_path'], filepath)
        absolute_path = os.path.abspath(absolute_path)
        absolute_paths.append(absolute_path)
    return absolute_paths


class Message(object):

    def __init__(self, json_dict: dict):
        # 从请求参数（来自ms-wechat微服务）初始化
        self.message_id = json_dict.get('messageId', '')  # 消息id
        self.from_subject = json_dict.get('fromSubject', '')   # 发送主体
        self.to_conversations = json_dict.get('toConversations', [])  # 发送给哪些会话（群聊）
        self.channel = json_dict.get('channel', '')   # 发送渠道
        self.message_type = json_dict.get('messageType', '')   # 消息类型
        self.message_data = json_dict.get('messageData', {})   # 消息内容

    def __str__(self):
        return json.dumps(self.__dict__, ensure_ascii=False)


# 抽象消息发送类
class MessageSender(object):

    def get_from_subject(self) -> str:
        pass

    def get_channel(self) -> str:
        pass

    def get_message_types(self) -> List[str]:
        return []

    def is_accept(self, message: Message):
        """是否支持处理这种类型的消息"""
        return message.from_subject == self.get_from_subject() and \
            message.channel == self.get_channel() and message.message_type in self.get_message_types()

    def check_valid(self, message: Message):
        """检查参数是否有效"""
        pass

    def lock_screen(self):
        """是否独占屏幕"""
        return True

    def send(self, message: Message):
        """处理发送"""
        pass

    def exec_command(self, command: str):
        """执行命令，快速调用其他一些方法"""
        pass


# 微信文本消息发送
class WechatTextMessageSender(MessageSender):

    def __init__(self, wechat_app: WechatApp):
        self.wechat_app = wechat_app

    def get_from_subject(self) -> str:
        return self.wechat_app.login_user_name

    def get_channel(self) -> str:
        return 'WECHAT'

    def get_message_types(self) -> List[str]:
        return ['TEXT']

    def check_valid(self, message: Message):
        check_param(message.message_data, '消息内容messageData不能为空', message)
        check_param(message.message_data.get('content', ''), '发送内容content不能为空', message)
        check_param(message.to_conversations, 'toConversations参数不能为空', message)
        check_param(self.wechat_app, '微信客户端未启动')

    def send(self, message: Message) -> List[dict]:
        send_results = self.wechat_app.batch_send_message(message.to_conversations,
                                                          message.message_data.get('content'), [])
        return [x.to_dict() for x in send_results]

    def exec_command(self, command: str):
        if command == 'active':
            self.wechat_app.active(force=True)
        elif command == 'check_skip_update':
            self.wechat_app.check_skip_update()
        elif command == 'login_confirm':
            self.wechat_app.check_and_login_after_logout()
            self.wechat_app.login_confirm()
        elif command == 'switch_account':
            self.wechat_app.switch_account()
        else:
            logger.warning('未实现的指令： {}'.format(command))


# 微信文件消息发送
class WechatFileMessageSender(WechatTextMessageSender):

    def get_message_types(self) -> List[str]:
        return ['FILE']

    def check_valid(self, message: Message):
        check_param(message.message_data, '消息内容messageData不能为空', message)
        check_param(message.message_data.get('filePaths', []), 'filePaths参数不能为空', message)
        check_param(message.to_conversations, 'toConversations参数不能为空', message)
        check_param(self.wechat_app, '微信客户端未启动')

    def send(self, message: Message) -> List[dict]:
        absolute_paths = get_upload_absolute_path(message.message_data.get('filePaths'))
        append_text = message.message_data.get('content', '')
        send_results = self.wechat_app.batch_send_message(message.to_conversations, append_text, absolute_paths)
        return [x.to_dict() for x in send_results]


# 微信文件消息发送
class WechatLinkCardMessageSender(WechatTextMessageSender):

    def get_message_types(self) -> List[str]:
        return ['LINK_CARD']

    def check_valid(self, message: Message):
        check_param(message.message_data, '消息内容messageData不能为空', message)
        check_param(message.message_data.get('content', ''), 'content参数不能为空', message)
        check_param(message.message_data.get('shareLink', ''), 'shareLink参数不能为空', message)
        check_param(message.to_conversations, 'toConversations参数不能为空', message)
        check_param(self.wechat_app, '微信客户端未启动')

    def send(self, message: Message) -> List[dict]:
        send_results = self.wechat_app.batch_send_message(message.to_conversations,
                                                          text=message.message_data.get('content'),
                                                          share_link=message.message_data.get('shareLink'))
        return [x.to_dict() for x in send_results]


class QQTextMessageSender(MessageSender):

    def __init__(self, qq_app: QQApp):
        self.qq_app = qq_app

    def get_from_subject(self) -> str:
        return self.qq_app.login_user_name

    def get_channel(self) -> str:
        return 'QQ'

    def get_message_types(self) -> List[str]:
        return ['TEXT']

    def check_valid(self, message: Message):
        check_param(message.message_data, '消息内容messageData不能为空', message)
        check_param(message.message_data.get('content', ''), '发送内容content不能为空', message)
        check_param(message.to_conversations, 'toConversations参数不能为空', message)
        check_param(self.qq_app, '微信客户端未启动')

    def send(self, message: Message) -> List[dict]:
        send_results = self.qq_app.batch_send_message(message.to_conversations,
                                                      text=message.message_data.get('content'),
                                                      check_pre_message=message.message_data.get('checkPreMessage', ''))
        return [x.to_dict() for x in send_results]

    def exec_command(self, command: str):
        if command == 'activeMain':
            active_window(self.qq_app.main_window, True)
        elif command == 'activeConversation':
            active_window(self.qq_app.conversation_window, True)
        elif command == 'check_skip_update':
            self.qq_app.check_skip_update()
        elif command == 'set_online_state':
            self.qq_app.set_online_state()
        elif command == 'click_toolbar_qq_icon':
            self.qq_app.click_toolbar_qq_icon()
        elif command == 'ready_login_qr_code':
            self.qq_app.ready_login_qr_code()
        elif command == 'close_conversation_window':
            self.qq_app.close_conversation_window(force=True)
        elif command == 'exit_account':
            self.qq_app.exit_account()
        elif command == 'switch_account':
            self.qq_app.switch_account()
        elif command == 'close_alert_window':
            self.qq_app.close_alert_window()
        elif command == 'start_new_qq':
            open_multi_app(CONFIG['qq_path'], 1)
        else:
            logger.warning('未实现的指令： {}'.format(command))


# 企微机器人消息发送
class WecomGroupBotMessageSender(MessageSender):

    def get_from_subject(self) -> str:
        return '--'

    def get_channel(self) -> str:
        return 'WECOM_GROUP_BOT'

    def is_accept(self, message: Message):
        return message.channel == self.get_channel()

    def check_valid(self, message: Message):
        check_param(message.from_subject, 'fromSubject参数不能为空', message)
        check_param(message.message_data, '消息内容messageData不能为空', message)

    def lock_screen(self):
        return False

    def send(self, message: Message):
        bot = CorpWechatBot(key=message.from_subject,  # 你的机器人key，通过群聊添加机器人获取
                            log_level=cptools.INFO,  # 设置日志发送等级，INFO, ERROR, WARNING, CRITICAL,可选
                            )
        if message.message_type == 'TEXT':
            bot.send_text(message.message_data.get('content'),
                          mentioned_list=message.message_data.get('mentionedList', []),
                          mentioned_mobile_list=message.message_data.get('mentionedMobileList', []))
        elif message.message_type == 'MARKDOWN':
            bot.send_markdown(message.message_data.get('content'))
        else:
            raise MessageSendException('暂不支持该发送类型: ' + message.message_type)


# 企微应用消息推送
class WecomAppMessageSender(MessageSender):

    def __init__(self, wecom_corp_app: AppMsgSender, from_subject: str):
        self.wecom_corp_app = wecom_corp_app
        self.from_subject = from_subject

    def get_from_subject(self) -> str:
        return self.from_subject

    def get_channel(self) -> str:
        return 'WECOM_APP_MESSAGE'

    def is_accept(self, message: Message):
        return message.channel == self.get_channel()

    def check_valid(self, message: Message):
        check_param(message.to_conversations, 'toConversations参数不能为空', message)

    def lock_screen(self):
        return False

    def send(self, message: Message):
        if message.message_type == 'TEXT':
            self.wecom_corp_app.send_text(content=message.message_data.get('content'),
                                          touser=message.to_conversations)
        elif message.message_type == 'MARKDOWN':
            self.wecom_corp_app.send_markdown(content=message.message_data.get('content'),
                                              touser=message.to_conversations)
        elif message.message_type == 'FILE':
            file_paths = get_upload_absolute_path(message.message_data.get('filePaths'))
            for file_path in file_paths:
                self.wecom_corp_app.send_file(file_path=file_path,
                                              touser=message.to_conversations)
        elif message.message_type == 'CARD':
            self.wecom_corp_app.send_card(title=message.message_data.get('title'),
                                          desp=message.message_data.get('content'),
                                          url=message.message_data.get('url'),
                                          btntxt=message.message_data.get('buttonText'),
                                          touser=message.to_conversations)
        else:
            raise MessageSendException('暂不支持该发送类型: ' + message.message_type)


class MessageSenderManager(object):

    def __init__(self):
        # 初始化消息发送器，如果初始化过程发生异常，可以第二次请求接口进行初始化
        self.wechat_apps: List[WechatApp] = []
        self.qq_apps: List[QQApp] = []
        self.wecom_corp_app = None
        self.message_senders: List[MessageSender] = []
        self.ready_message_senders()

    def ready_message_senders(self):
        self.message_senders = []
        # 初始化消息发送器，如果初始化过程发生异常，可以第二次请求接口进行初始化

        if CONFIG.get('start_wecom_group_bot', True):
            logger.info('启动企微机器人发送器')
            self.message_senders.append(WecomGroupBotMessageSender())

        if CONFIG.get('start_wecom_app_push', True):
            logger.info('启动企微应用消息发送器')
            self.wecom_corp_app = AppMsgSender(corpid=CONFIG['wecom_corp_id'],  # 你的企业id
                                               corpsecret=CONFIG['wecom_corp_key'],  # 你的应用凭证密钥
                                               agentid=CONFIG['wecom_agent_id'],  # 你的应用id
                                               log_level=cptools.INFO,  # 设置日志发送等级，INFO, ERROR, WARNING, CRITICAL,可选
                                               )
            self.message_senders.append(WecomAppMessageSender(self.wecom_corp_app, CONFIG['wecom_corp_name']))

        wechat_app_count = CONFIG.get('start_wechat_app_count', 0)
        if wechat_app_count:
            logger.info('配置启动微信客户端数量: {}'.format(wechat_app_count))
            self.wechat_apps = WechatApp.build_all_wechat_apps()
            if not self.wechat_apps:
                # 如果没有微信窗口，则需要启动客户端
                logger.info('启动微信客户端，数量： {}'.format(wechat_app_count))
                open_multi_app(CONFIG.get('wechat_path'), wechat_app_count)
                time.sleep(5)  # 等待启动完成
                self.wechat_apps = WechatApp.build_all_wechat_apps()
            # 所有微信窗口的发送示例
            for wechat_app in self.wechat_apps:
                self.message_senders.append(WechatTextMessageSender(wechat_app))
                self.message_senders.append(WechatFileMessageSender(wechat_app))
                self.message_senders.append(WechatLinkCardMessageSender(wechat_app))

        qq_app_count = CONFIG.get('start_qq_app_count', 0)
        if qq_app_count:
            logger.info('配置启动QQ客户端数量: {}'.format(qq_app_count))
            self.qq_apps = QQApp.build_all_qq_apps(CONFIG['qq_is_retain_conversation_window'],
                                                   CONFIG['qq_is_search_in_conversation_window'])
            if not self.qq_apps:
                logger.info('启动QQ客户端，数量: {}'.format(qq_app_count))
                open_multi_app(CONFIG.get('qq_path'), qq_app_count)
                time.sleep(5)
                self.qq_apps = QQApp.build_all_qq_apps()
            for qq_app in self.qq_apps:
                self.message_senders.append(QQTextMessageSender(qq_app))
        return self.message_senders

    def get_message_sender(self, message: Message):
        for message_sender in self.message_senders:
            if message_sender.is_accept(message):
                return message_sender
        raise MessageSendException('未找到匹配的消息发送器. channel: {}, fromSubject: {}'
                                   .format(message.channel, message.from_subject))

    def send_message_with_exception(self, message: Message):
        message_sender = self.get_message_sender(message)
        message_sender.check_valid(message)
        logger.info('send message with: {}'.format(message_sender.__class__.__name__))
        if message_sender.lock_screen():
            try:
                SCREEN_LOCK.acquire(timeout=3600)
                return message_sender.send(message)
            finally:
                SCREEN_LOCK.release()
        else:
            return message_sender.send(message)

    def senders_to_dict(self):
        senders = []
        index = 0
        for message_sender in self.message_senders:
            senders.append({
                'channel': message_sender.get_channel(),
                'fromSubject': message_sender.get_from_subject(),
                'type': message_sender.get_message_types(),
                'index': index
            })
            index += 1
        return senders


def benchmark_test():
    sender_manager = MessageSenderManager()
    json_messages = [
        # 微信文本消息发送
        {
            "channel": "WECHAT",
            "fromSubject": "悠悠",
            "toConversations": ["rpa消息测试群"],
            "messageType": "TEXT",
            "messageData": {
                "content": "这是一条测试消息"
            }
        },
        # 微信文件发送
        {
            "channel": "WECHAT",
            "fromSubject": "悠悠",
            "toConversations": ["rpa消息测试群", "文件传输助手"],
            "messageType": "FILE",
            "messageData": {
                "content": "请查看文件",
                "filePaths": ["test-upload-file.txt"]
            }
        },
        # 微信链接卡片
        {
            "channel": "WECHAT",
            "fromSubject": "悠悠",
            "toConversations": ["rpa消息测试群"],
            "messageType": "LINK_CARD",
            "messageData": {
                "content": "请查看分享链接",
                "shareLink": 'https://mp.weixin.qq.com/s/ZLWFNxknX6fgQ60Qbd0C-A'
            }
        },
        # QQ发送文本消息
        {
            "channel": "QQ",
            "fromSubject": "咏春 叶问",
            "messageType": "TEXT",
            "toConversations": ['2307672656', '文件交流'],
            "messageData": {
                "content": "无前置消息检查"
            },
        },
        {
            "channel": "QQ",
            "fromSubject": "咏春 叶问",
            "messageType": "TEXT",
            "toConversations": ['2307672656', '文件交流'],
            "messageData": {
                "checkPreMessage": "无前置消息检查",
                "content": "测试QQ文本消息发送-有前置消息检查"
            },
        },
        # 企微群机器人
        {
            "channel": "WECOM_GROUP_BOT",
            "fromSubject": "7653c432-a6c7-424f-abf4-e64e3757b7e7",
            "messageType": "TEXT",
            "messageData": {
                "content": "企微群机器人测试消息-by zhaohai",
                "mentionedList": ["zhaohai"]
            },
        },
        # 企微群机器人MARKDOWN 文本中@人
        {
            "channel": "WECOM_GROUP_BOT",
            "fromSubject": "7653c432-a6c7-424f-abf4-e64e3757b7e7",
            "messageType": "MARKDOWN",
            "messageData": {
                "content": "测试消息: <font color='warning'>** CJ0408 **</font> -by zh <@zhaohai> 后置消息"
            },
        },
        # 企微应用消息推送，需开通IP白名单，否则不允许调用接口
        {
            "channel": "WECOM_APP_MESSAGE",
            "messageType": "TEXT",
            "toConversations": ["zhaohai"],
            "messageData": {
                "content": "企微应用推送消息-by zhaohai"
            }
        },
    ]
    # json_messages = json_messages[0:1]  # 微信文本消息发送
    json_messages = json_messages[0:-3]
    for json_message in json_messages:
        message = Message(json_message)
        sender_manager.send_message_with_exception(message)


if __name__ == '__main__':
    benchmark_test()
