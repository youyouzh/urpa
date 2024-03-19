import json
import sys
import os
import time

# ntworkd github 参考： <https://github.com/Kaguya233qwq/ntwork>
# 目前测试支持企业微信4.0.8版本，客户端下载地址： <https://dldir1.qq.com/wework/work_weixin/WeCom_4.0.8.6027.exe>
import ntwork

from base.log import logger
from base.config import load_config
from base.util import get_with_retry

CONFIG = {
    "cache_path": r'cache'
}
load_config(CONFIG)
wework = ntwork.WeWork()


class WecomApp(object):
    def __init__(self):
        self.wework = wework
        self.group_map = {}
        self.contacts = {}
        self.rooms = []

    def ready_contacts(self):
        """
        处理该账号的联系人和群聊记录，用于关联发送人
        :return:
        """
        self.contacts = get_with_retry(wework.get_external_contacts)
        self.rooms = get_with_retry(wework.get_rooms)

        if not self.contacts or not self.rooms:
            return False

        # 将contacts和rooms保存到json文件中，可能需要上报
        with open(os.path.join(CONFIG['cache_path'], 'wework_contacts.json'), 'w', encoding='utf-8') as f:
            json.dump(self.contacts, f, ensure_ascii=False, indent=2)
        with open(os.path.join(CONFIG['cache_path'], 'wework_rooms.json'), 'w', encoding='utf-8') as f:
            json.dump(self.rooms, f, ensure_ascii=False, indent=2)

        # 创建一个空字典来保存结果
        result = {}
        # 遍历列表中的每个字典
        for room in self.rooms['room_list']:
            # 获取聊天室ID
            room_wxid = room['conversation_id']
            # 获取聊天室成员
            room_members = wework.get_room_members(room_wxid)
            # 将聊天室成员保存到结果字典中
            result[room_wxid] = room_members

        # 将结果保存到json文件中
        with open(os.path.join(CONFIG['cache_path'], 'wework_room_members.json'), 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=4)
        logger.info("wework init finished········")
        return True

    def startup(self):
        # 打开pc企业微信, smart: 是否管理已经登录的企业微信
        wework.open(smart=True)
        # 等待登录，会自动登录，一定要退出后重新启动才可以获取到token，如果在启动中是不可以的，会卡住
        wework.wait_login()
        self.ready_contacts()

    def send_text(self, conversation_id, message):
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


if __name__ == '__main__':
    # 收到Crtl+C则结束，关闭进程，没有下面的代码会卡死不会结束
    work_wechat_app = WecomApp()
    work_wechat_app.startup()
    try:
        while True:
            time.sleep(0.5)
    except KeyboardInterrupt:
        ntwork.exit_()
        sys.exit()
