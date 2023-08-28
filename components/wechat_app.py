"""
微信自动化操作封装，微信可以支持多开
方法一：微软应用商店再下载一个微信，和PC版共存。
方法二：鼠标右键微信图标，选择属性——目标——整行复制目标中的路径——新建记事本，输入start "" +刚刚复制的路径，
开几个微信就复制几行——保存记事本，改后缀名为.bat，启动的多个微信窗口重叠，拉开上面的窗口即可。
方法三：鼠标右键微信图标，然后快速敲回车。打开多少个，取决于你的手速。
方法四：左键点击微信，长按回车0.5秒。如果按的时间超过1秒，估计就要开了几十个微信。电脑马上会死机….慎用。
"""
import os.path
import random
import re
import threading
import time
from datetime import datetime
from enum import unique, Enum
from typing import List

import pyautogui
import uiautomation as auto
from uiautomation import Control

from base.control_util import select_control
from base.log import logger
from base.util import report_error, win32_clipboard_text, win32_clipboard_files, MessageSendException

WECHAT_LOCK = threading.Lock()
auto.SetGlobalSearchTimeout(5)


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
    MESSAGE_LIST = '消息列表'
    MESSAGE_INPUT = '消息输入框'
    MESSAGE_SEND_FILE = '发送文件按钮'
    CONVERSATION_LIST = '会话列表'
    CONVERSATION_SEARCH = '会话搜索'
    CONVERSATION_SEARCH_CLEAR = '清空搜索'
    CONVERSATION_ACTIVE_TITLE = '激活会话标题'
    CONVERSATION_SEARCH_RESULT = '会话搜索结果'
    NAVIGATION = '左侧导航窗口'


# 微信客户端封装
class WechatApp(object):

    def __init__(self, main_window=None):
        self.main_window = main_window
        # 微信中缓存控件map
        self.cached_control = {}
        # 微信当前状态
        self.state = 'INIT'
        # 当前登录的用户名称
        self.login_user_name = None
        # 当前激活聊天会话框名称
        self.active_conversation = None
        # 当前激活聊天会话框备注名
        self.active_conversation_remark = None
        self.history_message_map = {}
        # 截图保存路径
        self.screen_image_path = r'images'
        self.init_mian_window()
        self.init_login_user_name()

    @staticmethod
    def build_all_wechat_apps():
        # 处理同时打开多个微信的情况，找到指定微信名称的微信窗口
        wechat_apps = []
        root_control = auto.GetRootControl()
        for app_control in root_control.GetChildren():
            # 匹配微信主窗口
            if app_control.Name != '微信':
                continue
            # 匹配微信名称，从导航窗口下的第一个子元素可以取到
            wechat_apps.append(WechatApp(app_control))
        logger.info('检测到当前打开微信窗口个数： {}'.format(len(wechat_apps)))
        return wechat_apps

    def init_mian_window(self):
        # 如果main_window为空，默认选第一个
        if not self.main_window:
            self.main_window = auto.WindowControl(searchDepth=1, Name='微信')

    def init_login_user_name(self):
        # 设置微信名，需要已经登录，不跑出异常避免影响启动，比如没有登录也能启动
        nav_control = self.search_control(ControlTag.NAVIGATION, with_check=False)
        if nav_control:
            self.login_user_name = nav_control.GetFirstChildControl().Name

    def search_control(self, tag: ControlTag, use_cache=True, with_check=True) -> Control | None:
        """
        搜索微信中的空间元素，主要为了缓存一些控件，避免每次都便利控件树，便利空间数会花费几百毫秒以上
        下面的定位方式适用于最新版3.9.6，其他版本可能不适用
        :param tag: 控件标识
        :param use_cache: 是否使用缓存
        :param with_check: 是否坚持控件存在
        :return: 查询到的控件
        """
        # 使用缓存
        if use_cache and tag in self.cached_control and self.cached_control[tag]:
            logger.info('search control with cache: {}'.format(tag))
            return self.cached_control[tag]

        # 尝试查找控件
        find_control = None
        try:
            if tag == ControlTag.CONVERSATION_LIST:
                # 会话列表控件
                find_control = self.main_window.ListControl(Name='会话')
            elif tag == ControlTag.CONVERSATION_SEARCH:
                # 会话搜索控件
                find_control = self.main_window.EditControl(Name='搜索')
            elif tag == ControlTag.CONVERSATION_SEARCH_CLEAR:
                find_control = self.main_window.ButtonControl(Name='清空')
            elif tag == ControlTag.CONVERSATION_ACTIVE_TITLE:
                # 当前激活的会话，基于消息控件定位
                anchor_control = self.search_control(ControlTag.NAVIGATION, use_cache, with_check)
                find_control = select_control(anchor_control, 'p>pane:1>pane-6>pane:1>pane-2')
            elif tag == ControlTag.CONVERSATION_SEARCH_RESULT:
                # 搜索会话结果
                # find_control = self.main_window.ListControl(Name='搜索结果')
                find_control = self.main_window.ListControl(Name='@str:IDS_FAV_SEARCH_RESULT:3780')
            elif tag == ControlTag.MESSAGE_LIST:
                # 消息列表
                # message_list_selector = 'pane:1>pane:1>pane:2>pane>pane>pane>pane>pane:1>pane>pane>list'
                # list_control = select_control(self.__main_window, message_list_selector)
                find_control = self.main_window.ListControl(Name='消息')
            elif tag == ControlTag.MESSAGE_INPUT:
                # 消息输入框
                # selector = 'pane:1>pane:2>pane>pane>pane>pane>pane:1>pane:1>pane:1>pane>pane>edit'
                # selector += self.__active_selector_prefix
                # self.__message_input_control = select_control(self.__main_window, selector)
                # find_control = self.main_window.EditControl(Name='输入')
                # 通过表情按钮来定位
                anchor_control = self.main_window.ButtonControl(Name='表情')
                find_control = select_control(anchor_control, '.>.>pane>edit')
            elif tag == ControlTag.MESSAGE_SEND_FILE:
                # 发送文件按钮
                find_control = self.main_window.ButtonControl(Name='发送文件')
            elif tag == ControlTag.NAVIGATION:
                find_control = self.main_window.ToolBarControl(Name='导航', searchDepth=3)
        except LookupError as e:
            logger.error('can not find control: {}'.format(tag), e)
            if with_check:
                raise Exception('can not find control: {}'.format(tag))
            else:
                return None

        if not find_control or not find_control.Exists(1, 1):
            logger.error('can not find control: {}'.format(tag))
            if with_check:
                raise Exception('can not find control: {}'.format(tag))
            else:
                return None

        logger.info('Attach control success: {}'.format(tag))
        self.cached_control[tag] = find_control
        return self.cached_control[tag]

    # 激活微信窗口窗口
    def active(self):
        self.main_window.SetActive()

    # 控件点击，需要加入随机演示，会提前确保窗口 active
    def control_click(self, control: Control, right_click=False):
        if not control:
            logger.error('The control is None, can not click.')
            return
        # 确保激活当前窗口后点击，其中会检测当前窗口如果时候TopLevel不会进行操作
        self.active()
        logger.info('click control name: {} position: {}'.format(control.Name, control.BoundingRectangle))
        if right_click:
            control.RightClick()
        else:
            control.Click()
        time.sleep(random.uniform(0.5, 1))

    # 检查和发送登录二维码，登录二维码页面时返回Ture，否则返回False
    def check_and_send_login_qr_code(self) -> bool:
        qr_code_control = select_control(self.main_window, 'pane:1>pane>pane:1>pane>pane>pane>pane')
        if not qr_code_control:
            logger.info('Can not find qr code control.')
            return False
        child_controls = qr_code_control.GetChildren()
        if len(child_controls) >= 2 and child_controls[0].Name == '扫码登录' and child_controls[1].Name == '二维码':
            logger.info('Attach login qr code page')
            self.state = WechatAppState.QR_CODE
            # 截屏二维码并保存
            qr_code_rect = child_controls[1].BoundingRectangle
            qr_code_region = (qr_code_rect.left, qr_code_rect.top, qr_code_rect.width(), qr_code_rect.height())
            qr_code_save_path = os.path.join(self.screen_image_path, 'qr-code-{}.png'.format(time.time_ns()))
            pyautogui.screenshot(qr_code_save_path, region=qr_code_region)
            return True
        return False

    # 检查和登录确认，登录确认页面时返回Ture，否则返回False
    def check_and_confirm_login(self) -> bool:
        # 登录确认控件
        confirm_control = select_control(self.main_window, 'pane:1>pane>pane:1>pane>pane>pane>pane>pane:1>pane')
        if not confirm_control:
            logger.warning('Can not find login confirm control.')
            return False
        child_controls = confirm_control.GetChildren()
        # 检查是否登录确认
        if len(child_controls) >= 2 and child_controls[0].Name == '扫码完成' \
                and child_controls[1].Name == '需在手机上完成登录':
            logger.info('Attach login confirm page.')
            self.state = WechatAppState.LOGIN_CONFIRM

            # 获取登录用户名称
            user_control = select_control(confirm_control.GetParentControl().GetParentControl(), 'pane>pane')
            self.login_user_name = user_control.Name if user_control else None
            return True
        return False

    # 检查会话列表，有会话列表返回True，否则False
    def get_conversation_item_controls(self) -> List[Control]:
        conversation_list_control = self.search_control(ControlTag.CONVERSATION_LIST)
        conversation_item_controls = conversation_list_control.GetChildren()
        logger.info('conversation list size: {}'.format(len(conversation_item_controls)))
        return conversation_item_controls

    def roll_up_conversation_controls(self, roll_times):
        """
        滚动会话列表
        :return:
        :param roll_times: 滚动次数
        """
        conversation_list_control = self.search_control(ControlTag.CONVERSATION_LIST)
        for i in range(roll_times):
            time.sleep(0.5)
            conversation_list_control.WheelDown(wheelTimes=3, waitTime=0.1 * i)

    def attach_active_conversation(self):
        """
        获取当前打开的对话框，并同步当前激活会话，以实际消息框的标题为准，可能没有打开消息窗口，比如首次进入app
        :return: 当前激活的会话名称
        """
        # 激活会话标题控件
        conversation_title_control = self.search_control(ControlTag.CONVERSATION_ACTIVE_TITLE, with_check=False)
        conversation_remark_control = select_control(conversation_title_control, 'pane > text')
        if not conversation_remark_control:
            logger.warning('Can not find active conversation title control.')
            # 重置激活会话信息，避免缓存数据
            self.active_conversation = None
            self.active_conversation_remark = None
            return False
        self.active_conversation_remark = conversation_remark_control.Name
        # 替换群聊后面的人数，注意这儿取的是备注名称
        self.active_conversation_remark = re.sub(r' \(\d+\)', '', self.active_conversation_remark)
        source_conversation_control = select_control(conversation_title_control, 'text')
        if source_conversation_control:
            # 有原会话名称节点的case，取原会话名称
            self.active_conversation = source_conversation_control.Name
        else:
            self.active_conversation = self.active_conversation_remark
        # 新版本的微信可以直接从编辑框获取会话，可能没有打开会话（只能获取备注名）
        # self.active_conversation = self.search_control(ControlTag.MESSAGE_INPUT, with_check=False).Name
        logger.info('attach active conversation: {}, remark: {}'
                    .format(self.active_conversation, self.active_conversation_remark))
        return self.active_conversation

    # 查看当前打开的会话是否匹配
    def is_match_current_conversation(self, conversation):
        self.attach_active_conversation()
        return self.active_conversation == conversation or self.active_conversation_remark == conversation

    # 获取消息列表控件
    def get_message_item_controls(self, filter_time=True) -> List[Control]:
        message_list_control = self.search_control(ControlTag.MESSAGE_LIST)
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
        """
        获取某个对话中加载的消息列表
        :param conversation: 指定会话，如果未指定，则使用当前激活的会话，否则会切换到指定会话
        :return: 消息列表
        """
        if conversation:
            self.search_switch_conversation(conversation)
        messages = []
        message_item_controls = self.search_control(ControlTag.MESSAGE_LIST).GetChildren()
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
                    'content': message_content,
                    'control': message_item_control
                }
                messages.append(message)
                logger.info('message: {}'.format(message))
        logger.info('message list size: {}'.format(len(messages)))
        return messages

    def load_more_message(self, n=0.1):
        """
        当前激活的聊天款中加载更多消息
        :param n: 滚动参数
        :return:
        """
        n = 0.1 if n < 0.1 else 1 if n > 1 else n
        message_list_control = self.search_control(ControlTag.MESSAGE_LIST)
        message_list_control.WheelUp(wheelTimes=int(500 * n), waitTime=0.1)

    # 搜索快捷键，避免定位搜索框并移动鼠标
    def send_search_shortcut(self, anchor_control=None):
        anchor_control = anchor_control if anchor_control else self.main_window
        anchor_control.SendKeys('{Ctrl}f', waitTime=1)

    # 切换到指定对话
    def search_switch_conversation(self, conversation: str):
        logger.info('开始处理切换会话')
        target_conversation_control = None
        # 检查当前激活会话窗口，如果匹配则不需要切换【对于相同名称对话不能处理】
        if self.is_match_current_conversation(conversation):
            # 会话窗口没有变化则不进行切换
            logger.info('当前会话已经打开，无需进行切换: {}'.format(conversation))
            return True

        # 先检查当前会话列表中是否有匹配，避免搜索，对于有备注的会话，此处只能匹配备注名称
        for conversation_control in self.get_conversation_item_controls():
            if conversation_control.Name == conversation:
                self.control_click(conversation_control)
                # 需要检查是否切换成功，一般显示10个会话，可能因为屏幕原因未显示
                time.sleep(0.5)
                if self.is_match_current_conversation(conversation):
                    logger.info('会话列表中有需要切换的会话，直接点击切换无需搜索: {}'.format(conversation))
                    return True

        # 搜索会话，同时支持备注名称和原名称
        logger.info('搜索会话列表： {}'.format(conversation))
        self.main_window.SetFocus()
        time.sleep(0.2)
        self.send_search_shortcut()
        search_control = self.search_control(ControlTag.CONVERSATION_SEARCH)
        # self.control_click(search_control)  # 使用快捷键更快，不需要移动鼠标指针
        search_control.SendKeys('{Ctrl}a')  # 避免还有旧的搜索
        win32_clipboard_text(conversation)
        search_control.SendKeys('{Ctrl}v')

        # 检查搜索结果列表
        search_list_control = self.search_control(ControlTag.CONVERSATION_SEARCH_RESULT)
        result_type = ''
        for search_item in search_list_control.GetChildren():
            # result_type表示当前匹配的标签，比如 '联系人' '群聊'
            if search_item.ControlTypeName == 'PaneControl':
                result_type = search_item.GetFirstChildControl().Name
                continue

            # 匹配备注名称
            if conversation == search_item.Name:
                logger.info('搜索到会话： {}, 类型： {}'.format(conversation, result_type))
                self.control_click(search_item)
                break

            # 匹配原名称
            source_conversation_node = select_control(search_item, 'pane>pane>pane>text')
            if not source_conversation_node:
                continue
            # 提取匹配内容
            match_conversation = source_conversation_node.Name
            match_conversation = re.sub(r'^群聊名称: ', '', match_conversation)
            match_conversation = re.sub(r'<em>([^<]*)</em>', r'\1', match_conversation)
            if match_conversation == conversation:
                target_conversation_control = search_item
                logger.info('搜索到会话： {}, 类型： {}'.format(conversation, result_type))
                self.control_click(search_item)
                break

        # 检查是否切换成功
        time.sleep(0.5)
        if not self.is_match_current_conversation(conversation):
            # 可能点击按钮没有生效
            logger.warning('第一次切换会话失败：{}，尝试快捷键切换'.format(conversation))
            # 尝试第二次切换
            if target_conversation_control:
                target_conversation_control.SendKeys('{Enter}')
                # self.control_click(target_conversation_control)

        time.sleep(0.5)
        if not self.is_match_current_conversation(conversation):
            logger.error('切换会话失败: {}'.format(conversation))
            # 未搜索到会话，退出搜索，抛出异常
            # search_control.SendKeys('{esc}')  # 有时候会退出异常
            # 点击输入框可以退出搜索状态
            self.control_click(self.search_control(ControlTag.MESSAGE_INPUT, with_check=False))
            raise MessageSendException('未搜索到该私聊名称或者群聊名称：' + conversation)
        logger.info('切换会话成功: {}'.format(conversation))
        return True

    # 首次切换到新对话，需要记录历史消息，避免重复回答，可设置保留最近消息并处理，用于首次启动继续回复
    def record_history_message(self, conversation_name: str, keep_recent_count=1):
        if conversation_name in self.history_message_map:
            # 非首次启动，已经有历史消息记录，不处理
            return
        # 初始化，避免新对话没有任何消息，最后一条消息留用，后续会判断是否是自己发的消息，如果是对面发的消息，则可以回复
        self.history_message_map[conversation_name] = self.history_message_map.get(conversation_name, [])
        history_message_controls = self.get_message_item_controls()
        history_message_controls = history_message_controls[:-1 * keep_recent_count] if keep_recent_count > 0 \
            else history_message_controls
        for message_control in history_message_controls:
            self.history_message_map[conversation_name].append(message_control.Name)
        logger.info('conversation: {}, history message size: {}, record size: {}'
                    .format(conversation_name, len(history_message_controls),
                            len(self.history_message_map.get(conversation_name))))
        # logger.info('history message: {}'.format(self.__history_message_map.get(conversation_name)))

    def send_clipboard(self):
        """
        发送剪贴板内容到输入框，需要确保当前输入框是focus的
        :return:
        """
        input_control = self.search_control(ControlTag.MESSAGE_INPUT)
        input_control.SendKeys('{Ctrl}v')
        # 使用快捷键Enter发送消息，而不是点击，查找元素消耗0.5秒左右
        # send_button_control = wechat_windows.ButtonControl(Name='sendBtn', Depth=14).Click()
        input_control.SendKeys('{Enter}')

    # 发送文本内容消息
    def send_text_message(self, message, click_input_control=False, clear=False, check_send_success=False) -> bool:
        """
        向当前聊天窗口发送文本消息，支持换行
        :param message: 文本消息
        :param click_input_control: 是否点击一下消息输入框，默认False
        :param clear: 是否清空当前输入框内容，默认False
        :param check_send_success: 检测是否发送成功，通过聊天消息列表中看是否有刚才发送的消息来确认，并不十分准确，比如重复消息
        :return: 是否发送成功
        """
        logger.info('send message: {}'.format(message))
        input_control = self.search_control(ControlTag.MESSAGE_INPUT)
        if click_input_control:
            # 点击一下输入框
            self.control_click(input_control)
        if clear:
            # 清空输入框的内容
            input_control.SendKeys('{Ctrl}a', waitTime=0)
        # 使用粘贴板输入更快，还能处理换行符的问题
        # input_control.SendKeys(message)  # 换行符输入有问题，不能使用这种方式
        win32_clipboard_text(message)
        self.send_clipboard()

        # 检查是否发送成功，检查最近5条消息是否包含发送的消息，暂不处理离线发送失败
        if check_send_success:
            time.sleep(0.5)
            for message_control in self.get_message_item_controls()[-5:]:
                if message_control.Name == message:
                    return True
            raise MessageSendException('消息发送失败，请检查微信输入是否可用')
        return False

    def batch_send_message(self, conversations, message):
        """
        批量发送消息，通过转发来实现，先将消息发送到文件组手，再进行转发
        :param conversations: 需要转发到哪些会话（群）
        :param message: 消息
        :return:
        """
        temp_conversation = '文件传输助手'
        logger.info('批量发送消息。发送到： {}， 内容：{}'.format(conversations, message))
        self.search_switch_conversation(temp_conversation)
        self.send_text_message(message)
        self.batch_forward_message(temp_conversation, -1, conversations)

    def batch_forward_message(self, from_conversation, forward_message_index, to_conversations):
        """
        批量转发消息，可以实现批量群发功能
        :param from_conversation: 转发消息来源对话
        :param forward_message_index: 转发消息的下标，一般为-1，表示最后一条消息
        :param to_conversations: 需要转发到哪些对话
        :return:
        """
        # 切换到需要转发消息的会话，并定位消息，选择转发
        self.search_switch_conversation(from_conversation)
        logger.info('搜索并切换到会话: {}'.format(from_conversation))
        from_message_controls = self.get_message_item_controls()
        if abs(forward_message_index) >= len(from_message_controls):
            logger.error('未找到需要转发的消息，索引： {}'.format(forward_message_index))
            return False
        forward_message_control = from_message_controls[forward_message_index]
        logger.info('定位转发消息成功: {}'.format(forward_message_control.Name))
        # 右键转发消息，注意需要点击到消息体部分
        select_control(forward_message_control, 'pane>pane:1').RightClick()
        logger.info('右键要转发的消息，呼出转发菜单.')
        time.sleep(1)

        # 查找转发按钮
        forward_button_control = self.main_window.MenuItemControl(Name='转发...')
        if not forward_button_control.Exists(1, 1):
            report_error('未找到转发按钮，聊天消息框：{}'.format(from_conversation))
            return False

        # 点击转发按钮
        self.control_click(forward_button_control)
        logger.info('点击转发...按钮进行转发')

        # 点击多选按钮
        multi_select_button_control = self.main_window.ButtonControl(Name='多选')
        if not multi_select_button_control.Exists(1, 1):
            report_error('没有找到多选按钮')
            return False
        self.control_click(multi_select_button_control)
        logger.info('点检选中多选按钮，支持同时转发多个会话')
        time.sleep(0.5)

        # 搜索转发的群，此时弹出转发框，可以直接Ctrl+F搜索，其实不用搜索，自动focus到搜索
        # self.main_window.SendKeys('{Ctrl}f', waitTime=1)
        send_group_count = 0
        search_anchor_control = multi_select_button_control
        for to_conversation in to_conversations:
            self.send_search_shortcut(search_anchor_control)
            win32_clipboard_text(to_conversation)
            search_anchor_control.SendKeys('{Ctrl}a')  # 避免还有旧的搜索
            search_anchor_control.SendKeys('{Ctrl}v')
            logger.info('选择搜索结果中的会话: {}'.format(to_conversation))

            search_list_control = self.main_window.ListControl(Name='请勾选需要添加的联系人')
            if not search_list_control.Exists(1, 1):
                logger.info('未搜索到任何会话： {}'.format(to_conversation))
                continue
            logger.info('搜索会话数量: {}'.format(len(search_list_control.GetChildren())))
            for search_item_control in search_list_control.GetChildren():
                # 过滤“群聊”或者“联系人”
                if not search_item_control.Name or search_item_control.Name != to_conversation:
                    continue

                # 选中群聊
                logger.info('选中会话： {}'.format(search_item_control.Name))
                self.control_click(search_item_control)
                send_group_count += 1

        if send_group_count == 0:
            report_error('未搜索到需要发送的群聊。')
            return False

        # 点击发送按钮
        logger.info('点击分别发送按钮')
        send_button_control = self.main_window.ButtonControl(RegexName='分别发送')
        if not send_button_control.Exists(1, 1):
            report_error('未定位到发送按钮')
            return False
        self.control_click(send_button_control)
        return True

    # 发送图片消息
    def send_image_message(self, image_path):
        """
        向当前聊天窗口发送图片消息
        :param image_path: 图片绝对路径
        :return: None
        """
        send_file_button_control = self.search_control(ControlTag.MESSAGE_SEND_FILE)
        self.control_click(send_file_button_control)
        # 打开文件上传窗口
        file_select_window = self.main_window.WindowControl(searchDepth=1, Name='打开')
        # 选中文件并发送
        file_name_edit_control = file_select_window.ComboBoxControl(RegexName='文件名').EditControl(Depth=1)
        # 直接复制粘贴文件绝对路径并发送即可
        win32_clipboard_text(image_path)
        file_name_edit_control.SendKeys('{Ctrl}v')
        file_name_edit_control.SendKeys('{Enter}')
        self.main_window.SendKeys('{Enter}')

    # 发送文件消息
    def send_file_message(self, filepaths: list):
        """
        向当前聊天窗口发送文件
        :param filepaths: 要发送文件的绝对路径列表
        :return:
        """
        valid_paths = []
        for filepath in filepaths:
            if not os.path.exists(filepath):
                logger.warning('The file is not exist: {}'.format(filepath))
                continue
            valid_paths.append(os.path.abspath(filepath))
        logger.info('send file message: {}'.format(filepaths))
        win32_clipboard_files(valid_paths)
        self.send_clipboard()
        return True

    def batch_send_task(self, to_conversations: list, text: str):
        use_batch_send_min_count = 3
        if len(to_conversations) >= use_batch_send_min_count:
            # 超过 use_batch_send_min_count 个群使用批量转发
            self.batch_send_message(to_conversations, text)
        else:
            # 逐个发送
            for to_conversation in to_conversations:
                self.search_switch_conversation(to_conversation)
                self.send_text_message(text)


def test_forward_message():
    wechat_app = WechatApp()
    wechat_app.active()
    # wechat_app.batch_send_message('测试消息', ['文件传输助手', '悠悠'])
    wechat_app.batch_forward_message('文件传输助手', -1, ['文件传输助手', '悠悠'])


def test_send_text_message():
    wechat_app = WechatApp()
    wechat_app.active()
    wechat_app.search_control(ControlTag.MESSAGE_INPUT)
    # wechat_app.search_switch_conversation('逗再一起 乐逗离职群')
    # wechat_app.search_switch_conversation('小新闻')
    # wechat_app.search_switch_conversation('这是一条测试消息')
    # 长文本搜索
    wechat_app.search_switch_conversation('创金合信基金持仓查询项目沟通群')
    wechat_app.batch_send_task(['文件传输助手'], '1')
    # wechat_app.search_switch_conversation('持仓导出')


if __name__ == '__main__':
    test_send_text_message()
    # test_forward_message()
