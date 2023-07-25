import logging
import os.path
import re
import random
import time
from enum import unique, Enum
from datetime import datetime
from typing import List

import pyautogui
import pyperclip
import uiautomation as auto
from uiautomation import Control
from base.log import logger
from base.selector import select_control, select_parent_control


def parse_time_str(time_str: str):
    # 今天的消息
    now = datetime.now()
    zh_week = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
    try:
        if re.match(r'\d+:\d+', time_str):  # 处理今天的时间格式
            parsed_time = datetime.strptime(time_str, '%H:%M')
            return now.replace(hour=parsed_time.hour, minute=parsed_time.minute)
        elif '昨天' in time_str:  # 处理昨天的时间格式
            parsed_time = datetime.strptime(time_str, '昨天 %H:%M')
            return now.replace(day=now.day - 1, hour=parsed_time.hour, minute=parsed_time.minute)
        elif'星期' in time_str:  # 处理本周的时间格式
            dow = zh_week.index(time_str.split(' ')[0]) + 1  # 获取星期几
            parsed_time = datetime.strptime(time_str.split(' ')[1], '%H:%M')
            return now.replace(day=now.weekday() - dow, hour=parsed_time.hour, minute=parsed_time.minute)
        else:
            # 处理YYYY年MM月DD日的时间格式
            parsed_time = datetime.strptime(time_str, '%Y年%m月%d日 %H:%M')
            return parsed_time
    except ValueError:  # 处理格式错误的时间字符串
        logger.error(f'Time string "{time_str}" is invalid.')
    return None


@unique
class WechatAppState(str, Enum):
    INIT = 'INIT'  # 初始状态
    QR_CODE = 'QR_CODE'  # 等待扫描登录二维码
    LOGIN_CONFIRM = 'LOGIN_CONFIRM'  # 登录确认
    CONVERSATION = 'CONVERSATION'  # 会话列表页
    MESSAGE_RECEIVE = 'MESSAGE_RECEIVE'  # 接收消息中
    MESSAGE_EDIT = 'MESSAGE_EDIT'  # 消息编辑中
    INVALID = 'INVALID'  # 无效，比如窗口已关闭，发生异常等


@unique
class ControlTag(str, Enum):
    MESSAGE_LIST = 'MESSAGE_LIST'
    MESSAGE_INPUT = 'MESSAGE_INPUT'
    MESSAGE_SEND_FILE = 'MESSAGE_SEND_FILE'
    CONVERSATION_LIST = 'CONVERSATION_LIST'
    CONVERSATION_SEARCH = 'CONVERSATION_SEARCH'
    CONVERSATION_ACTIVE_TITLE = 'CONVERSATION_ACTIVE_TITLE'
    CONVERSATION_SEARCH_RESULT = 'CONVERSATION_SEARCH_RESULT'


# 微信客户端封装
class WechatApp(object):

    def __init__(self):
        self.__main_window = None
        self.__state = 'INIT'
        self.__login_user_name = None
        self.__cache_control = {}
        self.__active_conversation = None  # 当前激活会话
        self.__history_message_map = {}

    def cache_control(self, tag: ControlTag, use_cache=True, with_check=True) -> Control | None:
        # 使用缓存
        if use_cache and tag in self.__cache_control and self.__cache_control[tag]:
            return self.__cache_control[tag]

        # 尝试查找控件
        find_control = None
        try:
            if tag == ControlTag.CONVERSATION_LIST:
                # 会话列表控件
                # conversation_list_selector = 'pane:1.pane:1.pane:1.pane:1.pane.pane.pane.list'
                # list_control = select_control(self.__main_window, conversation_list_selector)
                find_control = self.__main_window.ListControl(Name='会话', Depth=8)
            elif tag == ControlTag.CONVERSATION_SEARCH:
                # 会话搜索
                find_control = self.__main_window.EditControl(Name='搜索', Depth=7)
            elif tag == ControlTag.CONVERSATION_ACTIVE_TITLE:
                # 当前激活的会话，基于消息控件定位
                message_list_control = self.cache_control(ControlTag.MESSAGE_LIST, with_check=with_check)
                base_control = select_parent_control(message_list_control, 4)
                find_control = select_control(base_control, 'pane.pane.pane:1.pane.pane.pane')
            elif tag == ControlTag.CONVERSATION_SEARCH_RESULT:
                # 搜索会话结果
                find_control = self.__main_window.ListControl(Name='搜索结果', Depth=6)
            elif tag == ControlTag.MESSAGE_LIST:
                # 消息列表
                # message_list_selector = 'pane:1.pane:1.pane:2.pane.pane.pane.pane.pane:1.pane.pane.list'
                # list_control = select_control(self.__main_window, message_list_selector)
                find_control = self.__main_window.ListControl(Name='消息', Depth=11)
            elif tag == ControlTag.MESSAGE_INPUT:
                # 消息输入框
                # selector = 'pane:1.pane:2.pane.pane.pane.pane.pane:1.pane:1.pane:1.pane.pane.edit'
                # selector += self.__active_selector_prefix
                # self.__message_input_control = select_control(self.__main_window, selector)
                find_control = self.__main_window.EditControl(Name='输入', Depth=13)
            elif tag == ControlTag.MESSAGE_SEND_FILE:
                # 发送文件按钮
                find_control = self.__main_window.ButtonControl(Name='发送文件', Depth=13)
        except LookupError as e:
            logger.error('can not find control: {}'.format(tag), e)
            if with_check:
                raise Exception('can not find control: {}'.format(tag))
            else:
                return None

        if not find_control:
            logging.error('can not find control: {}'.format(tag))
            if with_check:
                raise Exception('can not find control: {}'.format(tag))
            else:
                return None

        logger.info('Attach control success: {}'.format(tag))
        self.__cache_control[tag] = find_control
        return self.__cache_control[tag]

    def start(self):
        # 微信多开，需要将多个启动明确写入bat文件，同时打开多个，如果已经打开以后，后续都是该账号
        # subprocess.Popen(CONFIG.get('wechat_path'))  # 已经启动不再进行
        # os.system('start "{}"'.format(CONFIG.get('wechat_path')))
        self.__main_window = auto.WindowControl(searchDepth=1, Name='微信')
        if not self.__main_window:
            report_error('微信启动失败')
            return

    # 激活当前窗口
    def active(self):
        self.__main_window.SetActive()
        time.sleep(random.uniform(0.5, 1))

    # 控件点击，需要加入随机演示，会提前确保窗口 active
    def control_click(self, control: Control):
        if not control:
            logger.warning('The control is None, can not click.')
            return
        self.active()  # 确保激活当前窗口后点击
        logger.info('click control: {}'.format(control))
        control.Click()
        # time.sleep(random.uniform(0.5, 1))

    # 检查和发送登录二维码，登录二维码页面时返回Ture，否则返回False
    def check_and_send_login_qr_code(self) -> bool:
        qr_code_control = select_control(self.__main_window, 'pane:1.pane.pane:1.pane.pane.pane.pane')
        if not qr_code_control:
            logger.info('Can not find qr code control.')
            return False
        child_controls = qr_code_control.GetChildren()
        if len(child_controls) >= 2 and child_controls[0].Name == '扫码登录' and child_controls[1].Name == '二维码':
            logger.info('Attach login qr code page')
            self.__state = WechatAppState.QR_CODE
            # 截屏二维码并保存
            qr_code_rect = child_controls[1].BoundingRectangle
            qr_code_region = (qr_code_rect.left, qr_code_rect.top, qr_code_rect.width(), qr_code_rect.height())
            qr_code_save_path = os.path.join(CONFIG.get('screen_image_path'), 'qr-code-{}.png'.format(time.time_ns()))
            pyautogui.screenshot(qr_code_save_path, region=qr_code_region)
            return True
        return False

    # 检查和登录确认，登录确认页面时返回Ture，否则返回False
    def check_and_confirm_login(self) -> bool:
        # 登录确认控件
        confirm_control = select_control(self.__main_window, 'pane:1.pane.pane:1.pane.pane.pane.pane.pane:1.pane')
        if not confirm_control:
            logger.warning('Can not find login confirm control.')
            return False
        child_controls = confirm_control.GetChildren()
        # 检查是否登录确认
        if len(child_controls) >= 2 and child_controls[0].Name == '扫码完成' \
                and child_controls[1].Name == '需在手机上完成登录':
            logger.info('Attach login confirm page.')
            self.__state = WechatAppState.LOGIN_CONFIRM

            # 获取登录用户名称
            user_control = select_control(confirm_control.GetParentControl().GetParentControl(), 'pane.pane')
            self.__login_user_name = user_control.Name if user_control else None
            return True
        return False

    # 检查会话列表，有会话列表返回True，否则False
    def get_conversation_item_controls(self) -> List[Control]:
        conversation_list_control = self.cache_control(ControlTag.CONVERSATION_LIST)
        conversation_item_controls = conversation_list_control.GetChildren()
        logger.info('conversation list size: {}'.format(len(conversation_item_controls)))
        return conversation_item_controls

    # 获取当前打开的对话框，并同步当前激活会话，以实际消息框的标题为准，可能没有打开消息窗口，比如首次进入app
    def attach_active_conversation(self):
        # 激活会话标题控件
        conversation_title_control = self.cache_control(ControlTag.CONVERSATION_ACTIVE_TITLE, with_check=False)
        if not conversation_title_control or not conversation_title_control.GetFirstChildControl():
            logger.warning('Can not find active conversation title control.')
            return None
        self.__active_conversation = conversation_title_control.GetFirstChildControl().Name
        logger.info('attach active conversation: {}'.format(self.__active_conversation))
        return self.__active_conversation

    # 获取消息列表控件
    def get_message_item_controls(self, filter_time=True) -> List[Control]:
        message_list_control = self.cache_control(ControlTag.MESSAGE_LIST)
        message_item_controls = []
        for message_item_control in message_list_control.GetChildren():
            if filter_time and (message_item_control.GetFirstChildControl() is None
                                or message_item_control.GetFirstChildControl().GetFirstChildControl() is None):
                # logger.info('filter message: ' + message_item_control.Name)
                continue
            message_item_controls.append(message_item_control)
        logger.info('message list size: {}'.format(len(message_item_controls)))
        return message_item_controls

    # 获取当前激活对话的消息列表
    def get_conversation_messages(self, conversation: str = None):
        if conversation:
            self.search_switch_conversation(conversation)
        messages = []
        message_item_controls = self.cache_control(ControlTag.MESSAGE_LIST).GetChildren()
        base_time = datetime.now()  # 消息基准时间
        for message_item_control in message_item_controls:
            if message_item_control.GetFirstChildControl() is None:
                continue
            if message_item_control.GetFirstChildControl().GetFirstChildControl() is None:
                base_time = parse_time_str(message_item_control.Name)
                # logger.info('parse base time: {}'.format(base_time))
            elif message_item_control.Name == '以下为新消息':
                base_time = datetime.now()
            elif message_item_control.Name == '查看更多消息':
                continue
            else:
                sender = message_item_control.GetFirstChildControl().GetFirstChildControl().Name
                self_sender = False
                if not sender:
                    # 自己发送的消息
                    self_sender = True
                    sender = message_item_control.GetFirstChildControl().GetLastChildControl().Name
                message_content = message_item_control.Name
                message = {
                    'time': base_time.strftime("%Y-%m-%d %H:%M:%S"),
                    'sender': sender,
                    'self': self_sender,
                    'content': message_content
                }
                messages.append(message)
                logger.info('message: {}'.format(message))
        logger.info('message list size: {}'.format(len(messages)))
        return messages

    # 切换到指定对话
    def search_switch_conversation(self, conversation: str):
        logger.info('<--search_switch_conversation-->')
        # 检查当前激活会话窗口，如果匹配则不需要切换【对于相同名称对话不能处理】
        if self.attach_active_conversation() == conversation:
            # 会话窗口没有变化则不进行切换
            logger.info('current active conversation is match, not need to switch: {}'.format(conversation))
            return self.__active_conversation

        # 先检查当前会话列表中是否有匹配，避免搜索
        for conversation_control in self.get_conversation_item_controls():
            if conversation_control.Name == conversation:
                logger.info('conversation list item is match, direct switch: {}'.format(conversation))
                self.control_click(conversation_control)
                self.__active_conversation = conversation
                return

        # 搜索会话
        search_control = self.cache_control(ControlTag.CONVERSATION_SEARCH)
        self.control_click(search_control)
        pyperclip.copy(conversation)
        search_control.SendKeys('{Ctrl}a')  # 避免还有旧的搜索
        search_control.SendKeys('{Ctrl}v')

        # 检查搜索结果列表
        search_list_control = self.cache_control(ControlTag.CONVERSATION_SEARCH_RESULT)
        result_type = ''
        for search_item in search_list_control.GetChildren():
            if search_item.ControlTypeName == 'PaneControl':
                result_type = search_item.GetFirstChildControl().Name
                continue
            if conversation == search_item.Name:
                if result_type == '联系人':
                    # 私聊消息处理
                    logger.info('find match private conversation and switch: {}'.format(conversation))
                if result_type == '群聊':
                    # 群聊消息处理
                    logger.info('find match group conversation and switch: {}'.format(conversation))
                self.control_click(search_item)
                break

        # 检查是否切换成功
        if self.attach_active_conversation() != conversation:
            logger.error('search and switch conversation failed: {}'.format(conversation))
            return
        logger.info('search and switch conversation success: {}'.format(conversation))

    # 首次切换到新对话，需要记录历史消息，避免重复回答，可设置保留最近消息并处理，用于首次启动继续回复
    def record_history_message(self, conversation_name: str, keep_recent_count=1):
        if conversation_name in self.__history_message_map:
            # 非首次启动，已经有历史消息记录，不处理
            return
        # 初始化，避免新对话没有任何消息，最后一条消息留用，后续会判断是否是自己发的消息，如果是对面发的消息，则可以回复
        self.__history_message_map[conversation_name] = self.__history_message_map.get(conversation_name, [])
        history_message_controls = self.get_message_item_controls()
        history_message_controls = history_message_controls[:-1 * keep_recent_count] if keep_recent_count > 0 \
            else history_message_controls
        for message_control in history_message_controls:
            self.__history_message_map[conversation_name].append(message_control.Name)
        logger.info('conversation: {}, history message size: {}, record size: {}'
                    .format(conversation_name, len(history_message_controls),
                            len(self.__history_message_map.get(conversation_name))))
        # logger.info('history message: {}'.format(self.__history_message_map.get(conversation_name)))

    def get_history_message_map(self) -> dict:
        return self.__history_message_map

    # 发送文本内容消息
    def send_text_message(self, message) -> bool:
        logger.info('send message: {}'.format(message))
        input_control = self.cache_control(ControlTag.MESSAGE_INPUT)
        self.control_click(input_control)
        # 使用粘贴板输入更快，还能处理换行符的问题
        # input_control.SendKeys(message)  # 换行符输入有问题，不能使用这种方式
        pyperclip.copy(message)
        input_control.SendKeys('{Ctrl}v')
        # 使用快捷键Enter发送消息，而不是点击，查找元素消耗0.5秒左右
        # send_button_control = wechat_windows.ButtonControl(Name='sendBtn', Depth=14).Click()
        self.__main_window.SendKeys('{Enter}')

        # 检查是否发送成功，检查最近5条消息是否包含发送的消息，暂不处理离线发送失败
        time.sleep(0.5)
        for message_control in self.get_message_item_controls()[-5:]:
            if message_control.Name == message:
                return True
        return False

    # 发送图片消息
    def send_image_message(self, image_path):
        send_file_button_control = self.cache_control(ControlTag.MESSAGE_SEND_FILE)
        self.control_click(send_file_button_control)
        # 打开文件上传窗口
        file_select_window = self.__main_window.WindowControl(searchDepth=1, Name='打开')
        # 选中文件并发送
        file_name_edit_control = file_select_window.ComboBoxControl(RegexName='文件名').EditControl(Depth=1)
        # 直接复制粘贴文件绝对路径并发送即可
        pyperclip.copy(image_path)
        file_name_edit_control.SendKeys('{Ctrl}v')
        file_name_edit_control.SendKeys('{Enter}')
        self.__main_window.SendKeys('{Enter}')