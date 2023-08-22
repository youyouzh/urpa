"""
自动转发微信中的某个卡片消息
"""
import sys
import time
from datetime import datetime

import schedule
import uiautomation as auto

from base.config import load_config
from base.log import logger
from base.util import win32_clipboard_text

auto.uiautomation.SetGlobalSearchTimeout(3)  # 设置全局搜索超时 3

CONFIG = {
    # 转发消息来源
    'forward_message_from_conversation': '文件传输助手',

    # 转发消息索引
    'forward_message_index_from_conversation': -1,

    # 消息转发到哪些对话
    'forward_to_conversations': ['测试群01', '测试群02'],

    # 执行分钟间隔数
    'loop_minutes': 1,

    # 循环秒间隔
    'loop_seconds': 300,
}
load_config(CONFIG)


def report_error(message):
    logger.error('report error: {}'.format(message))


# 从微信聊天记录转发
def forward_from_conversation_message():
    from_conversation = CONFIG['forward_message_from_conversation']
    from_conversation_window = auto.WindowControl(Name=from_conversation, SearchDepth=1, ClassName='ChatWnd')
    from_conversation_window.SetActive()
    logger.info('定位并激活聊天会话窗口：{}'.format(from_conversation))

    # 查找聊天框
    if not from_conversation_window.Exists(1, 1):
        report_error('聊天窗口不存在： {}'.format(from_conversation_window))
        return False

    # 查找消息列表中的小程序卡片
    message_control = from_conversation_window.ListControl(Name='消息')
    if not message_control.Exists(1, 1) or not message_control.GetChildren():
        report_error('未找到聊天消息窗口：{}'.format(from_conversation))
        return False

    forward_message_control = message_control.GetChildren()[CONFIG['forward_message_index_from_conversation']]

    # 右键转发
    logger.info('右键小程序卡片')
    forward_message_control.RightClick()
    time.sleep(0.5)

    # 查找转发按钮
    forward_button_control = from_conversation_window.MenuItemControl(Name='转发...')
    if not forward_button_control.Exists(1, 1):
        report_error('未找到转发按钮，聊天消息框：{}'.format(from_conversation))
        return False

    # 查找转发按钮并点击
    logger.info('点击转发按钮')
    forward_button_control.Click()
    time.sleep(0.5)
    return from_conversation_window


def try_click_forward_menu(main_control):
    # 如果有转发菜单按钮，直接点击，否则点击菜单后转发
    try:
        forward_menu_control = main_control.MenuItemControl(Name='转发')
    except Exception | LookupError as e:
        report_error('没找到转发按钮')
        return False
    forward_menu_control.Click()
    time.sleep(0.5)
    return True


# 从小程序卡片转发
def forward_from_mini_program():
    mini_program_control = auto.PaneControl(Name='创金合信基金', searchDepth=1)
    if not mini_program_control.Exists(1, 1):
        report_error('未找到小程序窗口')
        return False
    mini_program_control.SetActive()
    logger.info('定位和激活小程序窗口')

    # 查找转发菜单然后点击转发
    menu_control = mini_program_control.ButtonControl(Name='菜单')
    if not menu_control.Exists(1, 1):
        report_error('未找到小程序菜单窗口')
        return False
    menu_control.Click()
    time.sleep(0.5)
    logger.info('点击小程序菜单按钮')
    try_click_forward_menu(mini_program_control)
    logger.info('点击小程序转发按钮')
    return auto.WindowControl(searchDepth=1, Name='微信', ClassName='SelectContactWnd')


# 选择转发群进行转发
def select_and_send_forward_conversation(forward_window):
    # 点击多选按钮
    multi_select_button_control = forward_window.ButtonControl(Name='多选')
    if not multi_select_button_control.Exists(1, 1):
        report_error('没有找到多选按钮')
        return False
    multi_select_button_control.Click()
    time.sleep(0.5)
    logger.info('点击多选按钮')

    # 搜索转发的群
    search_control = forward_window.EditControl(Name='搜索')
    if not search_control.Exists(1, 1):
        report_error('未找到搜索安按钮')
        return False
    logger.info('点击搜索按钮')
    send_group_count = 0
    for conversation in CONFIG['forward_to_conversations']:
        search_control.Click()
        time.sleep(0.5)
        win32_clipboard_text(conversation)
        search_control.SendKeys('{Ctrl}a')  # 避免还有旧的搜索
        search_control.SendKeys('{Ctrl}v')
        logger.info('选择搜索结果中的群聊: {}'.format(conversation))

        search_list_control = forward_window.ListControl(Name='请勾选需要添加的联系人')
        if not search_list_control.Exists(1, 1):
            logger.info('未搜索到任何群聊： {}'.format(conversation))
            continue
        logger.info('搜索群聊数量: {}'.format(len(search_list_control.GetChildren())))
        for search_item_control in search_list_control.GetChildren():
            # 第一个阶段是”群聊“
            if not search_item_control.Name or search_item_control.Name != conversation:
                continue

            # 选中群聊
            logger.info('选中群聊： {}'.format(search_item_control.Name))
            search_item_control.Click()
            send_group_count += 1
            time.sleep(0.5)

    if send_group_count == 0:
        report_error('未搜索到需要发送的群聊。')
        return False

    # 点击发送按钮
    logger.info('点击分别发送按钮')
    send_button_control = forward_window.ButtonControl(RegexName='分别发送')
    if not send_button_control.Exists(1, 1):
        report_error('未定位到发送按钮')
        return False
    send_button_control.Click()


def forward_job():
    try:
        logger.info('run loop forward job')
        # forward_from_mini_program()
        forward_window = forward_from_conversation_message()
        select_and_send_forward_conversation(forward_window)
    except KeyboardInterrupt:
        report_error('KeyboardInterrupt exit.')
        sys.exit()
    except Exception as e:
        report_error('Unknown Exception. {}'.format(e))


def main_loop():
    loop_times = 0
    while True:
        now = datetime.now()
        loop_rule = CONFIG['loop_rule']
        now_hour = str(now.hour)
        if now_hour in loop_rule and loop_times % loop_rule[now_hour] == 0:
            forward_job()
            loop_times = 0
        time.sleep(1)
        logger.info('loop times: {}'.format(loop_times))
        loop_times += 1


def main_loop_by_schedule():
    schedule.every(CONFIG['loop_minutes']).minutes.do(forward_job)
    while True:
        # 运行所有可以运行的任务
        schedule.run_pending()
        time.sleep(1)


# 打包： pyinstaller -F wechat_forward.py -p ../
if __name__ == '__main__':
    main_loop()
