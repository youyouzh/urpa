import json
import sys
import os
import time

# ntworkd github 参考： <https://github.com/Kaguya233qwq/ntwork>
import ntwork

from concurrent.futures import ThreadPoolExecutor
from base.log import logger

# 发送聊天消息的线程池
thread_pool = ThreadPoolExecutor(max_workers=8)

# 创建企业微信示例，目前测试支持企业微信4.0.8版本，客户端下载地址： <https://dldir1.qq.com/wework/work_weixin/WeCom_4.0.8.6027.exe>
wework = ntwork.WeWork()


class WecomApp(object):
    def __init__(self):
        self.group_map = {}

    def load_group_map(self):
        # 获取群列表，并做映射 conversation_id -> room 映射处理
        # 首次登录时才会获取，建议有值的时候进行缓存，然后从缓存读
        groups = wework.get_rooms()
        group_map_cache_file = 'group-map-cache.json'
        if 'room_list' in groups:
            for group in groups['room_list']:
                self.group_map[group['conversation_id']] = group
            json.dump(self.group_map, open(group_map_cache_file, 'w', encoding='utf-8'), ensure_ascii=False, indent=4)
        else:
            # 尝试从缓存加载，用于企业微信已经登录过的场景
            if os.path.exists(group_map_cache_file):
                self.group_map = json.load(open(group_map_cache_file, 'r', encoding='utf-8'))
            else:
                logger.warn('The cache file is empty.')
        logger.info(json.dumps(self.group_map, ensure_ascii=False))

    def startup(self):
        # 打开pc企业微信, smart: 是否管理已经登录的企业微信
        wework.open(smart=True)
        # 等待登录，会自动登录，一定要退出后重新启动才可以获取到token，如果在启动中是不可以的，会卡住
        wework.wait_login()
        self.load_group_map()

    def handle(self, message):
        # message包含两个字段，data和type，data为消息体
        data = message["data"]
        if data['conversation_id'].startswith("R:"):
            # 群聊消息，R:开头
            self.handle_group(message['data'])
        elif data['conversation_id'].startswith("S:"):
            # 私聊消息，S:开头
            self.handle_private(message['data'])
        else:
            logger.info('unknown type message. conversation_id: ' + data['conversation_id'])

    # 处理群发消息
    def handle_group(self, message):
        send_user_id = message['sender']
        # self_user_id = wework.get_login_info()["user_id"]
        content_type = message['content_type']
        content = message['content']
        at_list = message['at_list']
        group_name = self.group_map.get(message['conversation_id'], {}).get('nickname', '')
        # logger.info('handle group message from: {}, content: {}'.format(group_name, content))

    def handle_private(self, message):
        logger.info('handle private content: ' + message['content'])
        content = message['content']
        sender_name = message['sender_name']
        # logger.info('handle private message from: {}, content: {}'.format(sender_name, content))
        if '悠悠' in sender_name:
            thread_pool.submit(send_text, message['conversation_id'], '测试自动回复')
            # thread_pool.submit(send_text, 'R:182726087590753', '测试自动发送')


def send_text(conversation_id, message):
    wework.send_text(conversation_id, message)


def message_test(we_channel):
    # 测试消息处理
    example_message = {
        'data': {'appinfo': '573564633862363966623566616136301648196260_1216763776', 'at_list': [], 'content': '@bot 你好',
                 'content_type': 2, 'conversation_id': 'R:10822887504240095', 'is_pc': 0, 'local_id': '52',
                 'receiver': '1688854420437671', 'send_time': '1676254010', 'sender': '7881303235927249',
                 'sender_name': '悠悠', 'server_id': '1000604'}, 'type': 11041}
    we_channel.load_group_map()
    we_channel.handle(example_message)
    ntwork.exit_()
    sys.exit()


work_wechat_app = WecomApp()
work_wechat_app.startup()


# 注册消息回调
@wework.msg_register(ntwork.MT_RECV_TEXT_MSG)
def on_recv_text_msg(wework_instance: ntwork.WeWork, message):
    # message 内容参考 message-example.log 文件中的示例
    work_wechat_app.handle(message)


@wework.msg_register(ntwork.MT_ALL)
def on_receive_message(wework_instance: ntwork.WeWork, message):
    # message 内容参考 message-example.log 文件中的示例
    pass
    # logger.info('receive message: {}'.format(json.dumps(message)))


# 收到Crtl+C则结束，关闭进程，没有下面的代码会卡死不会结束
try:
    while True:
        time.sleep(0.5)
except KeyboardInterrupt:
    ntwork.exit_()
    sys.exit()
