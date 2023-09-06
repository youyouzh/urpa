"""
微信、企业微信、QQ消息发送流程处理
"""
import json

import cptools
import threading
from corpwechatbot import CorpWechatBot, AppMsgSender

from base.log import logger
from base.config import CONFIG
from base.util import MessageSendException, get_upload_absolute_path
from components.wechat_app import WechatApp

SCREEN_LOCK = threading.Lock()


def check_param(value, message, extra_data=None):
    if not value:
        raise MessageSendException(message, extra_data)


class Message(object):

    def __init__(self):
        self.from_subject = ''  # 发送主体
        self.to_conversations = []  # 发送给哪些会话（群聊）
        self.channel = ''  # 发送渠道
        self.message_type = ''  # 消息类型
        self.message_data = {}  # 消息内容

    def init_from_json(self, json_data: dict):
        # 从请求参数（来自ms-wechat微服务）初始化
        self.from_subject = json_data.get('fromSubject', '')
        self.to_conversations = json_data.get('toConversations', [])
        self.channel = json_data.get('channel', '')
        self.message_type = json_data.get('messageType', '')
        self.message_data = json_data.get('messageData', {})

    def __str__(self):
        return json.dumps(self.__dict__, ensure_ascii=False)


# 抽象消息发送类
class MessageSender(object):

    def is_accept(self, message: Message):
        """是否支持处理这种类型的消息"""
        pass

    def check_valid(self, message: Message):
        """检查参数是否有效"""
        pass

    def lock_screen(self):
        """是否独占屏幕"""
        return True

    def send(self, message: Message):
        """处理发送"""
        pass


# 微信文本消息发送
class WechatTextMessageSender(MessageSender):

    def __init__(self, wechat_app: WechatApp):
        self.wechat_app = wechat_app

    def is_accept(self, message: Message):
        return message.from_subject == self.wechat_app.login_user_name and \
               message.channel == 'WECHAT' and message.message_type == 'TEXT'

    def check_valid(self, message: Message):
        check_param(message.message_data, '消息内容messageData不能为空', message)
        check_param(message.message_data.get('content', ''), '发送内容content不能为空', message)
        check_param(message.to_conversations, 'toConversations参数不能为空', message)
        check_param(self.wechat_app, '微信客户端未启动')

    def send(self, message: Message):
        self.wechat_app.batch_send_task(message.to_conversations, message.message_data.get('content'))


# 微信文件消息发送
class WechatFileMessageSender(MessageSender):

    def __init__(self, wechat_app: WechatApp):
        self.wechat_app = wechat_app

    def is_accept(self, message: Message):
        return message.from_subject == self.wechat_app.login_user_name and \
               message.channel == 'WECHAT' and message.message_type == 'FILE'

    def check_valid(self, message: Message):
        check_param(message.message_data, '消息内容messageData不能为空', message)
        check_param(message.message_data.get('filePaths', []), 'filePaths参数不能为空', message)
        check_param(message.to_conversations, 'toConversations参数不能为空', message)
        check_param(self.wechat_app, '微信客户端未启动')

    def send(self, message: Message):
        for to_conversation in message.to_conversations:
            self.wechat_app.search_switch_conversation(to_conversation)
            self.wechat_app.send_file_message(get_upload_absolute_path(message.message_data.get('filePaths')))


# 企微机器人消息发送
class WecomGroupBotMessageSender(MessageSender):

    def is_accept(self, message: Message):
        return message.channel == 'WECOM_GROUP_BOT'

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
            bot.send_text(message.message_data.get('content'))
        elif message.message_type == 'MARKDOWN':
            bot.send_markdown(message.message_data.get('content'))
        else:
            raise MessageSendException('暂不支持该发送类型: ' + message.message_type, message)


# 企微应用消息推送
class WecomAppMessageSender(MessageSender):

    def __init__(self, wecom_corp_app: AppMsgSender):
        self.wecom_corp_app = wecom_corp_app

    def is_accept(self, message: Message):
        return message.channel == 'WECOM_APP_PUSH'

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
                                          btntxt=message.message_data.get('buttonText'))
        else:
            raise MessageSendException('暂不支持该发送类型: ' + message.message_type, message)


class MessageSenderManager(object):

    def __init__(self):
        # 初始化消息发送器，如果初始化过程发生异常，可以第二次请求接口进行初始化
        self.wechat_apps = None
        self.wecom_corp_app = None
        self.message_senders = []
        self.ready_message_senders()

    def ready_message_senders(self):
        # 初始化消息发送器，如果初始化过程发生异常，可以第二次请求接口进行初始化
        self.wechat_apps = WechatApp.build_all_wechat_apps()
        self.wecom_corp_app = AppMsgSender(corpid=CONFIG['wecom_corp_id'],  # 你的企业id
                                           corpsecret=CONFIG['wecom_corp_key'],  # 你的应用凭证密钥
                                           agentid=CONFIG['wecom_agent_id'],  # 你的应用id
                                           log_level=cptools.INFO,  # 设置日志发送等级，INFO, ERROR, WARNING, CRITICAL,可选
                                           )
        self.message_senders = [
            WecomGroupBotMessageSender(),
            WecomAppMessageSender(self.wecom_corp_app)
        ]
        # 所有微信窗口的发送示例
        for wechat_app in self.wechat_apps:
            self.message_senders.append(WechatTextMessageSender(wechat_app))
            self.message_senders.append(WechatFileMessageSender(wechat_app))
        return self.message_senders

    def get_message_sender(self, message: Message):
        for message_sender in self.message_senders:
            if message_sender.is_accept(message):
                return message_sender
        logger.error('Can not find match message sender: {}'.format(message))
        raise MessageSendException('Can not find match message sender.')

    def send_message_with_exception(self, message: Message):
        message_sender = self.get_message_sender(message)
        message_sender.check_valid(message)
        logger.info('send message with: {}'.format(message_sender.__class__.__name__))
        if message_sender.lock_screen():
            try:
                SCREEN_LOCK.acquire(timeout=3600)
                message_sender.send(message)
            finally:
                SCREEN_LOCK.release()
        else:
            message_sender.send(message)


def benchmark_test():
    sender_manager = MessageSenderManager()
    json_messages = [
        # 微信文本消息发送
        {
            "channel": "WECHAT",
            "fromSubject": "",
            "toConversations": ["文件传输助手"],
            "messageType": "TEXT",
            "messageData": {
                "content": "这是一条urpa测试消息"
            }
        },
        # 微信文件发送
        {
            "channel": "WECHAT",
            "fromSubject": "",
            "toConversations": ["文件传输助手"],
            "messageType": "FILE",
            "messageData": {
                "filePaths": [r"test-upload-file.txt"]
            }
        },
        # 企微群机器人
        {
            "channel": "WECOM_GROUP_BOT",
            "fromSubject": "5e86bea4-a978-4c6a-97dd-35544e52485c",
            "messageType": "TEXT",
            "messageData": {
                "content": "这是一条自动发送的测试消息"
            }
        },
        # 企微应用消息推送，需开通IP白名单，否则不允许调用接口
        {
            "channel": "WECOM_APP_PUSH",
            "messageType": "TEXT",
            "toConversations": ["xxxx"],
            "messageData": {
                "content": "这是一条自动发送的测试消息"
            }
        },
    ]
    json_messages = json_messages[1:2]  # 测试单类发送
    for json_message in json_messages:
        message = Message()
        message.init_from_json(json_message)
        sender_manager.send_message_with_exception(message)


if __name__ == '__main__':
    benchmark_test()
