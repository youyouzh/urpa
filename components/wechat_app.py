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
import time
from datetime import datetime
from enum import unique, Enum
from typing import List

import uiautomation as auto
from uiautomation import Control

from base.control_util import select_control, control_click, check_control_exist, find_top_window_controls, \
    active_window
from base.exception import ControlInvalidException, MessageSendException
from base.log import logger
from base.util import win32_clipboard_text, win32_clipboard_files, get_screenshot

auto.SetGlobalSearchTimeout(5)
BATCH_SEND_WITH_FORWARD_COUNT = 2  # 批量发送时使用转发的最小会话数量
FORWARD_MAX_CONVERSATION_COUNT = 9  # 微信转发消息最大会话数量
CHECK_SEND_SUCCESS_SIZE = 3  # 检查是否发送成功时获取的消息数量
TEMP_CONVERSATION = '文件传输助手'   # 临时中转会话


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


class SendResult(object):

    def __init__(self):
        self.to_conversation = ''  # 发送的会话
        self.is_success = False  # 是否发送成功
        self.error_message = ''  # 错误消息

    @staticmethod
    def success(to_conversation: str):
        send_result = SendResult()
        send_result.to_conversation = to_conversation
        send_result.is_success = True
        send_result.error_message = 'success'
        return send_result

    @staticmethod
    def fail(to_conversation: str, error_message: str):
        send_result = SendResult()
        send_result.to_conversation = to_conversation
        send_result.is_success = False
        send_result.error_message = error_message
        return send_result

    def to_dict(self) -> dict:
        return {
            'toConversation': self.to_conversation,
            'isSuccess': self.is_success,
            'errorMessage': self.error_message
        }

    def __str__(self):
        return f'SendResult: [to_conversation]: {self.to_conversation}, [error_message]: {self.error_message}'


class MessageInfo(object):

    def __init__(self, message_control: Control):
        self.message_control = message_control
        self.type: str = 'text'
        self.content: str = message_control.Name
        self.filepath = ''
        self.link_url = ''
        self.process_file()

    def process_file(self):
        if not self.message_control.Name == '[文件]':
            return
        filepath_control = select_control(self.message_control, 'pane>pane:1>pane-6>text')
        self.type = 'file'
        if filepath_control:
            self.filepath = filepath_control.Name

    def process_link(self):
        if not self.message_control.Name == '[链接]':
            return
        self.type = 'link'
        self.link_url = ''

    def __eq__(self, other):
        return self.content == other.content and self.filepath == self.filepath


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
    CONVERSATION_SEARCH_RESULT = '会话搜索结果'
    CONVERSATION_SEARCH_CLEAR = '清空搜索'
    MAIN_CONVERSATION_SEARCH = '主窗口搜索'
    MAIN_CONVERSATION_SEARCH_RESULT = '主窗口会话搜索结果'
    MAIN_CONVERSATION_SEARCH_CLEAR = '主窗口清空搜索'
    CONVERSATION_ACTIVE_TITLE = '激活会话标题'
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
        self._init_mian_window()
        self._init_login_user_name()
        self.check_skip_update()

    @staticmethod
    def build_all_wechat_apps():
        # 处理同时打开多个微信的情况，找到指定微信名称的微信窗口
        wechat_apps = []
        weixin_window_controls = find_top_window_controls('微信')
        for weixin_window_control in weixin_window_controls:
            wechat_app = WechatApp(weixin_window_control)
            wechat_apps.append(wechat_app)
        logger.info('检测到当前打开微信窗口个数： {}'.format(len(wechat_apps)))
        return wechat_apps

    def _init_mian_window(self):
        if self.main_window:
            # 已经绑定main_window时不处理
            return True

        weixin_window_controls = find_top_window_controls('微信')
        if not weixin_window_controls:
            logger.error('未检测到任何打开的微信窗口，绑定失败')
            return False

        self.main_window = weixin_window_controls[0]
        logger.info('初始化绑定微信窗口成功：{}'.format(self.main_window))
        # self.active()  # 初始化时不置顶，发送的时候置顶
        return True

    def _init_login_user_name(self):
        # 设置微信名，需要已经登录，不抛出异常避免影响启动，比如没有登录也能启动
        nav_control = self._search_control(ControlTag.NAVIGATION, with_check=False)
        if nav_control:
            # 已经登录的情况可以找到侧边栏
            self.login_user_name = nav_control.GetFirstChildControl().Name
            logger.info('init login user name: {}'.format(self.login_user_name))
        else:
            # 没有登录的话，点击触发登录确认按钮，并置顶窗口方便扫码
            self.active(force=True)   # 置顶，避免二维码被遮住
            self.check_and_login_after_logout()
            self.login_confirm()
            self.close_alter_window()

    def _search_control(self, tag: ControlTag, use_cache=True, with_check=True) -> Control | None:
        """
        搜索微信中的空间元素，主要为了缓存一些控件，避免每次都便利控件树，便利空间数会花费几百毫秒以上
        下面的定位方式适用于最新版3.9.6，其他版本可能不适用
        :param tag: 控件标识
        :param use_cache: 是否使用缓存
        :param with_check: 是否坚持控件存在
        :return: 查询到的控件
        """
        # 使用缓存，非空才使用
        if use_cache and tag in self.cached_control and self.cached_control[tag]:
            logger.info('search control with cache: {}'.format(tag))
            return self.cached_control[tag]

        if not self.main_window:
            logger.error('还未绑定主窗口')
            return None

        # 尝试查找控件
        try:
            find_control = self._search_control_by_tag(tag, use_cache, with_check)
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

    def _search_control_by_tag(self, tag: ControlTag, use_cache=True, with_check=True):
        find_control = None
        if tag == ControlTag.CONVERSATION_LIST:
            # 会话列表控件
            find_control = self.main_window.ListControl(Name='会话')
        elif tag == ControlTag.CONVERSATION_SEARCH:
            # 会话搜索控件
            find_control = self.main_window.EditControl(Name='搜索')
        elif tag == ControlTag.CONVERSATION_SEARCH_CLEAR:
            anchor_control = self._search_control(ControlTag.CONVERSATION_SEARCH, use_cache, with_check)
            if anchor_control:
                find_control = anchor_control.GetParentControl().GetLastChildControl()
        elif tag == ControlTag.CONVERSATION_ACTIVE_TITLE:
            # 当前激活的会话，基于消息控件定位
            anchor_control = self._search_control(ControlTag.NAVIGATION, use_cache, with_check)
            if anchor_control:
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
            if anchor_control:
                find_control = select_control(anchor_control, '.>.>pane>edit')
        elif tag == ControlTag.MESSAGE_SEND_FILE:
            # 发送文件按钮
            find_control = self.main_window.ButtonControl(Name='发送文件')
        elif tag == ControlTag.NAVIGATION:
            find_control = self.main_window.ToolBarControl(Name='导航', searchDepth=3)
        return find_control

    # 激活微信窗口窗口
    def active(self, force=False, wait_time=0.5):
        active_window(self.main_window, force, wait_time)

    # 处理在其他PC上登录被踢出的场景
    def check_and_login_after_logout(self):
        confirm_button_control = self.main_window.ButtonControl(Name='确定')
        if not confirm_button_control.Exists(1, 1):
            logger.info('未找到【确定】按钮')
            return False
        info_control = select_control(confirm_button_control, 'p>p>pane>pane>text')
        if not info_control:
            logger.info('未找到关于确定按钮的提示信息')
            return False
        logger.info('退出登录信息: {}'.format(info_control.Name))
        # 点击确定按钮
        control_click(confirm_button_control)
        return True

    # 关闭提示框，避免卡住
    def close_alter_window(self):
        # 取消按钮
        cancel_button = select_control(self.main_window, '>:1>:2>:5>button:2')
        if cancel_button and cancel_button.Name == '取消':
            logger.info('点击【取消】按钮关闭提示框')
            control_click(cancel_button)

        # 多选后不能转发提示框
        i_known_button = select_control(self.main_window, '>:1>>:2>:1')
        if i_known_button and i_known_button.Name == '我知道了':
            logger.info('点击【我知道了】按钮关闭微信提示框')
            # 点击【关闭多选】工具栏
            control_click(i_known_button)
            close_multi_button = select_control(self._search_control(ControlTag.MESSAGE_LIST), 'p>p>p>:1>:1>>:1')
            if close_multi_button and close_multi_button.Name == '关闭多选':
                logger.info('点击【关闭多选】按钮关闭多选工具栏')
                control_click(close_multi_button)

    # 处理已经登录过不需要扫描二维码的场景
    def login_confirm(self):
        login_control: Control = self.main_window.ButtonControl(Name='登录')
        if not login_control.Exists(1, 1):
            logger.info('未找到登录按钮')
            return False
        login_account = login_control.GetParentControl().GetFirstChildControl()
        logger.info('登录账号信息： {}'.format(login_account))
        # 点击登录按钮，等待手机确认登录
        control_click(login_control)
        self.login_user_name = login_account

    # 检查并跳过强制更新
    def check_skip_update(self):
        update_control: Control = None
        for control in self.main_window.GetChildren():
            if control.Name == '升级':
                update_control = control
                break
        if not update_control:
            logger.info('没有更新窗口不需要操作.')
            return False
        # 忽略更新按钮点击
        skip_update_button_control: Control = update_control.ButtonControl(Name='忽略本次更新')
        if not check_control_exist(skip_update_button_control):
            logger.info('没有【忽略本次更新】按钮，不需要操作.')
            return False
        logger.info('点击【忽略本次更新】按钮.')
        control_click(skip_update_button_control)
        return True

    # 登录时点击切换账号按钮
    def switch_account(self):
        switch_account_control: Control = self.main_window.ButtonControl(Name='切换账号')
        if not check_control_exist(switch_account_control):
            logger.warning('未找到【切换账号】按钮')
            return False
        control_click(switch_account_control)
        return True

    # 检查和发送登录二维码，登录二维码页面时返回Ture，否则返回False
    def check_and_send_login_qr_code(self) -> bool:
        qr_code_control: Control = self.main_window.ImageControl(Name='二维码')
        if not check_control_exist(qr_code_control):
            logger.warning('未找到二维码.')
            return False
        qr_code_rect = qr_code_control.BoundingRectangle
        qr_code_region = (qr_code_rect.left, qr_code_rect.top, qr_code_rect.width(), qr_code_rect.height())
        screen_path = get_screenshot(qr_code_region)
        return screen_path

    # 检查会话列表，有会话列表返回True，否则False
    def _get_conversation_item_controls(self) -> List[Control]:
        conversation_list_control = self._search_control(ControlTag.CONVERSATION_LIST)
        conversation_item_controls = conversation_list_control.GetChildren()
        logger.info('conversation list size: {}'.format(len(conversation_item_controls)))
        return conversation_item_controls

    def _roll_up_conversation_controls(self, roll_times):
        """
        滚动会话列表
        :return:
        :param roll_times: 滚动次数
        """
        conversation_list_control = self._search_control(ControlTag.CONVERSATION_LIST)
        for i in range(roll_times):
            time.sleep(0.5)
            conversation_list_control.WheelDown(wheelTimes=3, waitTime=0.1 * i)

    def _attach_active_conversation(self):
        """
        获取当前打开的对话框，并同步当前激活会话，以实际消息框的标题为准，可能没有打开消息窗口，比如首次进入app
        :return: 当前激活的会话名称
        """
        # 激活会话标题控件
        conversation_title_control = self._search_control(ControlTag.CONVERSATION_ACTIVE_TITLE, with_check=False)
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
        logger.info('attach active conversation: {}, remark: {}'
                    .format(self.active_conversation, self.active_conversation_remark))
        return self.active_conversation

    # 查看当前打开的会话是否匹配
    def _is_match_current_conversation(self, conversation):
        self._attach_active_conversation()
        return self.active_conversation == conversation or self.active_conversation_remark == conversation

    # 获取消息列表控件
    def _get_message_item_controls(self, filter_time=True) -> List[Control]:
        message_list_control = self._search_control(ControlTag.MESSAGE_LIST)
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
            self._search_switch_conversation(conversation)
        messages = []
        message_item_controls = self._search_control(ControlTag.MESSAGE_LIST).GetChildren()
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

    def _load_more_message(self, n=0.1):
        """
        当前激活的聊天款中加载更多消息
        :param n: 滚动参数
        :return:
        """
        n = 0.1 if n < 0.1 else 1 if n > 1 else n
        message_list_control = self._search_control(ControlTag.MESSAGE_LIST)
        message_list_control.WheelUp(wheelTimes=int(500 * n), waitTime=0.1)

    # 搜索快捷键，避免定位搜索框并移动鼠标
    def _send_search_shortcut(self, anchor_control=None):
        anchor_control = anchor_control if anchor_control else self.main_window
        anchor_control.SendKeys('{Ctrl}f', waitTime=0.5)

    # 切换到指定对话
    def _search_switch_conversation(self, conversation: str):
        logger.info('开始处理切换会话: {}'.format(conversation))
        self.active()  # 先置顶窗口
        # 检查当前激活会话窗口，如果匹配则不需要切换【对于相同名称对话不能处理】
        if self._is_match_current_conversation(conversation):
            # 会话窗口没有变化则不进行切换
            logger.info('当前会话已经打开，无需进行切换: {}'.format(conversation))
            return True

        # 先检查当前会话列表中是否有匹配，避免搜索，对于有备注的会话，此处只能匹配备注名称
        for conversation_control in self._get_conversation_item_controls():
            if conversation_control.Name == conversation:
                control_click(conversation_control)
                # 需要检查是否切换成功，一般显示10个会话，可能因为屏幕原因未显示
                if self._is_match_current_conversation(conversation):
                    logger.info('会话列表中有需要切换的会话，直接点击切换无需搜索: {}'.format(conversation))
                    return True

        # 搜索会话，同时支持备注名称和原名称
        logger.info('搜索会话列表： {}'.format(conversation))
        self._send_search_shortcut()
        search_control = self._search_control(ControlTag.CONVERSATION_SEARCH)
        # self.control_click(search_control)  # 使用快捷键更快，不需要移动鼠标指针
        search_control.SendKeys('{Ctrl}a')  # 避免还有旧的搜索
        win32_clipboard_text(conversation)
        search_control.SendKeys('{Ctrl}v')

        # 检查搜索结果列表
        search_list_control = self._search_control(ControlTag.CONVERSATION_SEARCH_RESULT)
        result_type = ''
        for search_item in search_list_control.GetChildren():
            # result_type表示当前匹配的标签，比如 '联系人', '群聊', '聊天记录'
            if search_item.ControlTypeName == 'PaneControl' and search_item.GetFirstChildControl():
                result_type = search_item.GetFirstChildControl().Name
                continue

            # 暂时只匹配联系人、群聊、公众号
            if result_type not in ['联系人', '群聊', '公众号']:
                continue

            # 匹配备注名称
            if conversation == search_item.Name:
                logger.info('搜索到会话： {}, 类型： {}'.format(conversation, result_type))
                control_click(search_item)
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
                control_click(search_item)

                # 检查是否切换成功，有时候点击切换会不生效，所以重试第二次
                if not self._is_match_current_conversation(conversation):
                    # 可能点击按钮没有生效
                    logger.warning('第一次切换会话失败：{}，尝试快捷键切换'.format(conversation))
                    # 尝试第二次切换
                    if target_conversation_control:
                        target_conversation_control.SendKeys('{Enter}')
                        time.sleep(0.5)
                        # self.control_click(target_conversation_control)

        if not self._is_match_current_conversation(conversation):
            logger.error('切换会话失败: {}'.format(conversation))
            # 未搜索到会话，退出搜索，抛出异常
            control_click(self._search_control(ControlTag.CONVERSATION_SEARCH_CLEAR, with_check=False))
            raise ControlInvalidException('未搜索到该私聊名称或者群聊名称：' + conversation)
        logger.info('切换会话成功: {}'.format(conversation))
        return True

    # 首次切换到新对话，需要记录历史消息，避免重复回答，可设置保留最近消息并处理，用于首次启动继续回复
    def record_history_message(self, conversation_name: str, keep_recent_count=1):
        if conversation_name in self.history_message_map:
            # 非首次启动，已经有历史消息记录，不处理
            return
        # 初始化，避免新对话没有任何消息，最后一条消息留用，后续会判断是否是自己发的消息，如果是对面发的消息，则可以回复
        self.history_message_map[conversation_name] = self.history_message_map.get(conversation_name, [])
        history_message_controls = self._get_message_item_controls()
        history_message_controls = history_message_controls[:-1 * keep_recent_count] if keep_recent_count > 0 \
            else history_message_controls
        for message_control in history_message_controls:
            self.history_message_map[conversation_name].append(message_control.Name)
        logger.info('conversation: {}, history message size: {}, record size: {}'
                    .format(conversation_name, len(history_message_controls),
                            len(self.history_message_map.get(conversation_name))))
        # logger.info('history message: {}'.format(self.__history_message_map.get(conversation_name)))

    def _send_clipboard_messages(self, click_input_control=True, clear=False):
        """
        发送剪贴板内容到输入框，需要确保当前输入框是focus的
        :param click_input_control: 是否点击一下消息输入框，默认False
        :param clear: 是否清空当前输入框内容，默认False
        :return:
        """
        input_control = self._search_control(ControlTag.MESSAGE_INPUT)
        if click_input_control:
            # 点击一下输入框，确保聚焦
            control_click(input_control)
        if clear:
            # 清空输入框的内容
            input_control.SendKeys('{Ctrl}a', waitTime=0)
        input_control.SendKeys('{Ctrl}v')
        # 使用快捷键Enter发送消息，而不是点击，查找元素消耗0.5秒左右
        # send_button_control = wechat_windows.ButtonControl(Name='sendBtn', Depth=14).Click()
        time.sleep(random.uniform(0.3, 0.5))
        input_control.SendKeys('{Enter}')

    def send_text_message(self, message, check_send_success=True) -> List[Control]:
        """
        向当前聊天窗口发送文本消息，支持换行
        :param message: 文本消息
        :param check_send_success: 检测是否发送成功，通过聊天消息列表中看是否有刚才发送的消息来确认，并不十分准确，比如重复消息
        :return: 发送的消息控件
        """
        logger.info('send text message: {}'.format(message))
        # 使用粘贴板输入更快，还能处理换行符的问题
        # input_control.SendKeys(message)  # 换行符输入有问题，不能使用这种方式
        win32_clipboard_text(message)
        self._send_clipboard_messages()

        # 获取最后5条消息，比较是否已经发送
        send_message_controls = []
        check_message_controls = self._get_message_item_controls()[-CHECK_SEND_SUCCESS_SIZE:]
        check_message_controls.reverse()
        for message_control in check_message_controls:
            # 发送文本消息，尾部的换行不会发送，比较时去掉尾部的换行符
            if MessageInfo(message_control).content.rstrip('\n') == message.rstrip('\n'):
                send_message_controls.append(message_control)
                break

        if check_send_success and not send_message_controls:
            # 检查有没有发送成功
            raise MessageSendException('消息发送失败，请重试')
        return send_message_controls

    # 发送图片消息
    def send_image_message(self, image_path):
        """
        向当前聊天窗口发送图片消息
        :param image_path: 图片绝对路径
        :return: None
        """
        send_file_button_control = self._search_control(ControlTag.MESSAGE_SEND_FILE)
        control_click(send_file_button_control)
        # 打开文件上传窗口
        file_select_window = self.main_window.WindowControl(searchDepth=1, Name='打开')
        check_control_exist(file_select_window, '未找到【打开】窗口')
        # 选中文件并发送
        file_name_edit_control = file_select_window.ComboBoxControl(RegexName='文件名').EditControl(Depth=1)
        check_control_exist(file_name_edit_control, '未找到【文件名】编辑窗口')
        # 直接复制粘贴文件绝对路径并发送即可
        win32_clipboard_text(image_path)
        file_name_edit_control.SendKeys('{Ctrl}v')
        file_name_edit_control.SendKeys('{Enter}')
        self.main_window.SendKeys('{Enter}')

    # 发送文件消息
    def send_file_message(self, filepaths: list, check_send_success=True) -> List[Control]:
        """
        向当前聊天窗口发送文件
        :param filepaths: 要发送文件的绝对路径列表
        :param check_send_success: 强制检查是否发送成功
        :return: 发送的消息控件
        """
        valid_paths = []
        for filepath in filepaths:
            if not os.path.exists(filepath):
                # 如果有文件不存在，直接异常，直接跳过反而不好处理
                logger.warning('The file is not exist: {}'.format(filepath))
                raise MessageSendException('发送文件已被删除：{}'.format(filepath))
            if os.path.getsize(filepath) == 0:
                logger.warning('The file is empty: {}'.format(filepath))
                raise MessageSendException('发送文件为空文件：{}'.format(filepath))
            valid_paths.append(os.path.abspath(filepath))
        if not valid_paths:
            logger.error('发送文件全部无效: {}'.format(filepaths))
            raise MessageSendException('发送文件为空: {}'.format(filepaths))
        logger.info('send file message: {}'.format(filepaths))
        win32_clipboard_files(valid_paths)
        self._send_clipboard_messages()

        # 获取最后5条消息，比较是否已经发送
        send_message_controls = []
        filenames = [os.path.basename(x) for x in filepaths]
        check_message_controls = self._get_message_item_controls()[-CHECK_SEND_SUCCESS_SIZE:]
        check_message_controls.reverse()
        for message_control in check_message_controls:
            message_info = MessageInfo(message_control)
            if message_info.filepath in filenames:
                filenames.remove(message_info.filepath)  # 避免转发时重复
                send_message_controls.append(message_control)

        if check_send_success and len(send_message_controls) != len(valid_paths):
            raise MessageSendException('消息发送失败，请重试')
        return send_message_controls

    def _batch_forward_message(self, from_conversation: str,
                               to_conversations: list,
                               forward_message_controls: list) -> List[str]:
        """
        批量转发消息，可以实现批量群发功能
        :param from_conversation: 转发消息来源对话
        :param forward_message_controls: 需要转发的消息控件列表
        :param to_conversations: 需要转发到哪些对话
        :return: 发送成功的群列表
        """
        # 切换到需要转发消息的会话，并定位消息，选择转发
        logger.info('批量转发消息，消息来源： {}, 转发到： {}, 转发消息数量： {}'
                    .format(from_conversation, to_conversations, len(forward_message_controls)))
        self._search_switch_conversation(from_conversation)
        logger.info('搜索并切换到会话: {}'.format(from_conversation))
        if not forward_message_controls:
            raise ControlInvalidException('未找到需要转发的消息，消息来源： {}'.format(from_conversation))
        if not to_conversations:
            raise ControlInvalidException('需要转发消息的会话列表为空，消息来源： {}'.format(from_conversation))

        # 单条需要转发消息直接点击换出转发按钮然后转发
        if len(forward_message_controls) == 1:
            self._click_one_forward_message(forward_message_controls[0])
        else:
            # 多条消息转发，则需要进入多选界面然后选择需要转发的消息进行转发
            self._select_multi_forward_message(forward_message_controls)

        # 查找转发的联系人选择窗口
        select_contact_window = find_top_window_controls(class_name='SelectContactWnd',
                                                         with_exception_message='未找到转发的联系人选择窗口',
                                                         root_control=self.main_window)[0]
        # 选中需要转发到哪些会话
        real_forward_conversations = self._select_forward_conversations(select_contact_window, to_conversations)
        if len(real_forward_conversations) == 0:
            raise ControlInvalidException('未搜索到需要发送的群聊: {}'.format(to_conversations))

        # 点击发送按钮
        logger.info('点击【分别发送】按钮')
        send_button_control: Control = select_contact_window.ButtonControl(RegexName='分别发送')
        control_click(send_button_control, with_exception_message='未定位到发送按钮，消息来源: {}'.format(from_conversation))
        return real_forward_conversations

    def _open_link_browser_by_link(self, link: str):
        # 将链接发送到临时会话【文件传输助手】
        self._search_switch_conversation(TEMP_CONVERSATION)
        link_message_controls = self.send_text_message(link, check_send_success=True)

        # 点击链接，注意需要点击到消息体部分
        select_control(link_message_controls[0], 'pane>pane:1').Click()
        time.sleep(1)

        # 检查浏览窗口是否打开
        return find_top_window_controls('微信', 'Chrome_WidgetWin', '未找到微信内置浏览器窗口')[0]

    def _open_link_browser_by_account(self, link: str):
        # 首先搜索并去到 创金科技研发部
        self._search_switch_conversation('创金科技研发部')
        # 通过消息输出款定位到跳转链接按钮
        message_list_control = self._search_control(ControlTag.MESSAGE_LIST)
        jump_link_button = select_control(message_list_control, 'p-3>pane:1>pane>button')
        control_click(jump_link_button, with_exception_message='未找到跳转链接按钮')
        time.sleep(1)

        # 检查浏览窗口是否打开
        link_browser_window = find_top_window_controls('微信', 'Chrome_WidgetWin', '未找到微信内置浏览器窗口')[0]

        # 查找链接输入按钮，并输入链接
        time.sleep(3)
        link_input_control = link_browser_window.EditControl(Name='请输入链接')
        # 复制粘贴链接，进入页面
        control_click(link_input_control, with_exception_message='未找到微信内置浏览器中的链接输入框')
        win32_clipboard_text(link)
        # link_browser_window.SendKeys('Enter')

        # 确定按钮
        control_click(select_control(link_input_control, 'p>button'))
        return link_browser_window

    def send_link_card_message(self, link: str, to_conversation: str) -> Control:
        logger.info('发送连接卡片. link: {}, to_conversation: {}'.format(link, to_conversation))
        link_browser_window = self._open_link_browser_by_link(link)
        time.sleep(2)

        # 点击转发按钮
        more_menu_control = link_browser_window.MenuItemControl(Name='更多')
        control_click(more_menu_control, with_exception_message='未找到微信内置浏览器中的【更多】菜单按钮')

        # 点击转发按钮
        forward_menu_control = link_browser_window.MenuItemControl(Name='转发给朋友')
        control_click(forward_menu_control, with_exception_message='未找到微信内置浏览器中的【转发给朋友】菜单按钮')

        # 获取转发窗口
        select_contact_window = find_top_window_controls(class_name='SelectContactWnd',
                                                         with_exception_message='未找到转发的联系人选择窗口')[0]
        self._select_forward_conversations(select_contact_window, [to_conversation])

        # 点击发送按钮
        logger.info('点击【分别发送】按钮')
        send_button_control = select_contact_window.ButtonControl(RegexName='分别发送')
        control_click(send_button_control, with_exception_message='未找到【分别发送】按钮')

        # 关闭窗口
        link_browser_window.SendKeys('{Ctrl}w')
        self._search_switch_conversation(to_conversation)

        check_message_control = self._get_message_item_controls()[-1]
        if check_message_control.Name != '[链接]':
            raise MessageSendException('转发链接失败： {}'.format(link))
        return check_message_control

    # 点击单条消息转发按钮
    def _click_one_forward_message(self, forward_message_control):
        # 右键转发消息，注意需要点击到消息体部分
        select_control(forward_message_control, 'pane>pane:1').RightClick()
        logger.info('右键要转发的消息，呼出转发菜单.')
        time.sleep(1)

        # 查找转发按钮
        # 如果是文件，需要上传完后才能进行转发，需要等一会儿
        logger.info('点击【转发...】按钮进行转发')
        forward_button_control = self.main_window.MenuItemControl(Name='转发...')
        if not check_control_exist(forward_button_control):
            error_message = '未找到消息【转发】按钮，可能消息不能转发，请稍后重试'
            logger.error(error_message)
            # 可能窗口卡住，需要关闭一下提示框
            self.close_alter_window()
            raise ControlInvalidException(error_message)

    # 选中需要转发的消息列表
    def _select_multi_forward_message(self, forward_message_controls):
        # 右键任意一条消息，点击多选按钮
        select_control(forward_message_controls[0], 'pane>pane:1').RightClick()
        logger.info('右键其中一条需要转发的消息')
        time.sleep(1)

        # 点击多选按钮
        logger.info('点击【多选】菜单按钮')
        multi_button_control = self.main_window.MenuItemControl(Name='多选')
        control_click(multi_button_control, with_exception_message='未找到消息【多选】按钮，请稍后重试')

        # 逐条消息点击选中
        logger.info('逐条选中所有需要转发的消息')
        forward_message_controls = forward_message_controls[1:]  # 多选按钮的那条消息默认是勾选的
        for forward_message_control in forward_message_controls:
            control_click(forward_message_control)

        # 点击逐条转发按钮进行消息转发
        logger.info('点击【逐条转发】菜单按钮')
        # 从消息列表来定位搜索比较快，【多选】工具栏
        forward_button_control = select_control(self._search_control(ControlTag.MESSAGE_LIST), 'p>p>p>:1>:1>>>>button')
        control_click(forward_button_control, with_exception_message='未找到消息【逐条转发】按钮，请稍后重试', check_name='逐条转发')
        return True

    # 选中需要转发到哪些会话
    def _select_forward_conversations(self, select_contact_window: Control, forward_conversations, append_text=''):
        # 点击多选按钮
        multi_select_button_control = select_contact_window.ButtonControl(Name='多选')
        if not check_control_exist(multi_select_button_control):
            # 此时可能触发不能转发的提示框，检查一下抛出异常
            self.close_alter_window()
            raise ControlInvalidException('转发消息到多个会话时未找到【多选】按钮，请稍后重试')
        control_click(multi_select_button_control)
        logger.info('点击选中多选按钮，支持同时转发多个会话')

        # 搜索转发的群，此时弹出转发框，可以直接Ctrl+F搜索，其实不用搜索，自动focus到搜索
        # self.main_window.SendKeys('{Ctrl}f', waitTime=1)
        real_forward_conversations = []   # 记录实际转发成功的会话列表
        search_anchor_control = multi_select_button_control
        for to_conversation in forward_conversations:
            self._send_search_shortcut(search_anchor_control)
            win32_clipboard_text(to_conversation)
            search_anchor_control.SendKeys('{Ctrl}a')  # 避免还有旧的搜索
            search_anchor_control.SendKeys('{Ctrl}v')
            logger.info('选择搜索结果中的会话: {}'.format(to_conversation))

            search_list_control = select_contact_window.ListControl(Name='请勾选需要添加的联系人')
            if not check_control_exist(search_list_control):
                logger.warning('未搜索到任何会话： {}'.format(to_conversation))
                continue
            logger.info('搜索会话数量: {}'.format(len(search_list_control.GetChildren())))
            for search_item_control in search_list_control.GetChildren():
                # 过滤“群聊”或者“联系人”
                if not search_item_control.Name or search_item_control.Name != to_conversation:
                    continue

                # 选中群聊
                logger.info('选中会话： {}'.format(search_item_control.Name))
                control_click(search_item_control)
                real_forward_conversations.append(to_conversation)
        # 留言处理，如果包含换行，则只能取第一行的内容
        if append_text:
            if '\n' in append_text:
                raise MessageSendException('转发留言不能包含换行: {}'.format(append_text))
            append_text_control = select_contact_window.EditControl(Name='给朋友留言')
            control_click(append_text_control)
            win32_clipboard_text(append_text)
        return real_forward_conversations

    def batch_send_message(self, to_conversations: list, text='', filepaths=None,
                           share_link='', check_pre_message='') -> List[SendResult]:
        """
        批量发送文本或者文件消息，通过转发来实现，先将消息发送到文件助手，再进行转发
        如果text和filepaths都不为空，则连续发送文件和文本，一般发送文件会有一个提示语
        :param to_conversations: 发送的会话列表
        :param text: 发送的文本消息内容，如果filepaths不为空，则表示附加的文本消息
        :param filepaths: 发送的文件消息内容
        :param share_link: 发送的卡片链接
        :param check_pre_message: 检查前置消息
        :return:
        """
        filepaths = filepaths if filepaths else []
        logger.info('批量发送消息。发送到: {}，文本: {}，文件: {}，链接：{}'.format(to_conversations, text, filepaths, share_link))
        if not text and not filepaths and not share_link:
            raise MessageSendException('发送内容为空，请检查参数')
        send_results = []
        self.active(force=True)
        if len(to_conversations) >= BATCH_SEND_WITH_FORWARD_COUNT:
            # 超过 use_batch_send_min_count 个群使用批量转发
            # 首先将消息发送到文件助手，此时发送用严格检查
            self._search_switch_conversation(TEMP_CONVERSATION)
            send_message_controls = []
            # 先发连接消息，因为会多发一条连接文本
            if share_link:
                send_message_controls.append(self.send_link_card_message(share_link, TEMP_CONVERSATION))
            # 先发文件再发文本消息
            if filepaths:
                # 发送文件
                send_message_controls.extend(self.send_file_message(filepaths))
                # 发送文件比较慢，需要等一下全部上传以后才能转发
                total_file_size = 0
                for filepath in filepaths:
                    total_file_size += os.path.getsize(filepath)
                # 基础3秒，500K多1秒
                wait_file_upload_seconds = 3 + total_file_size // 1024 // 500
                logger.info('批量转发文件，文件数：{}，总大小：{}，等待秒：{}'
                            .format(len(filepaths), total_file_size, wait_file_upload_seconds))
                time.sleep(wait_file_upload_seconds)
            if text:
                # 发送文本消息
                send_message_controls.extend(self.send_text_message(text))

            # ----> 微信批量转发一次只能转发9个群，所以要分批处理，转发成功后，会停留在 TEMP_CONVERSATION 会话
            page_size = FORWARD_MAX_CONVERSATION_COUNT
            total_pages = len(to_conversations) // page_size + (len(to_conversations) % page_size > 0)
            for page in range(0, total_pages):
                page_to_conversations = to_conversations[page * page_size:(page + 1) * page_size]
                # 通过文件助手对消息进行转发，前置保证消息发送成功了
                try:
                    send_success_conversations = self._batch_forward_message(TEMP_CONVERSATION,
                                                                             page_to_conversations,
                                                                             send_message_controls)
                    # 处理判断发送成功或失败的会话
                    for to_conversation in page_to_conversations:
                        if to_conversation in send_success_conversations:
                            send_results.append(SendResult.success(to_conversation))
                        else:
                            send_results.append(SendResult.fail(to_conversation, '未搜索到该群聊: ' + to_conversation))
                except ControlInvalidException as exception:
                    # 如果发送过程中出现异常，则整个批次都失败
                    logger.error('转发消息异常，page_to_conversations: {}, exception: {}'
                                 .format(page_to_conversations, exception.message), stack_info=True)
                    for to_conversation in page_to_conversations:
                        send_results.append(SendResult.fail(to_conversation, f'{exception.message}:{to_conversation}'))
        else:
            # 逐个发送，如果有一个发送失败，则全失败
            for to_conversation in to_conversations:
                try:
                    self._search_switch_conversation(to_conversation)
                    send_message_controls = []
                    if share_link:
                        send_message_controls.append(self.send_link_card_message(share_link, to_conversation))
                    if filepaths:
                        send_message_controls.extend(self.send_file_message(filepaths))
                    if text:
                        send_message_controls.extend(self.send_text_message(text))

                    # 严格判断是否发送成功
                    if send_message_controls:
                        send_results.append(SendResult.success(to_conversation))
                    else:
                        send_results.append(SendResult.fail(to_conversation, '未发送成功需重试'))
                except ControlInvalidException as exception:
                    logger.warning('conversation send exception. conversation: {}'.format(to_conversation), stack_info=True)
                    send_results.append(SendResult.fail(to_conversation, exception.message))

        send_fail_results = [x for x in send_results if not x.is_success]
        if send_fail_results and len(send_fail_results) == len(to_conversations):
            # 全部发送失败，则直接抛出异常结束
            raise MessageSendException('全部消息发送失败：{}...'
                                       .format('、'.join([x.error_message for x in send_fail_results[:2]])))
        return send_results


def test_send_message():
    wechat_app = WechatApp()
    wechat_app.active()
    wechat_app.close_alter_window()
    # wechat_app._get_link_browser_window_control()
    # wechat_app.send_link_card_message('https://mp.weixin.qq.com/s/ZLWFNxknX6fgQ60Qbd0C-A', TEMP_CONVERSATION)
    # wechat_app._search_switch_conversation('文件传输助手')
    # forward_message_controls = wechat_app._get_message_item_controls()
    # forward_message_controls = forward_message_controls[-2:]
    # wechat_app._select_multi_forward_message(forward_message_controls)


def test_batch_send():
    wechat_app = WechatApp()
    wechat_app.active()
    select_control(wechat_app.main_window, ':1>>:1>-4>text')
    to_conversations = ['文件传输助手', 'urpa测试', 'rpa消息测试群']
    filepaths = [r'build\测试-大文件.pptx', r'build\test-file-1-2.txt']
    # send_results = wechat_app.batch_send_text_or_files(to_conversations, '测试消息-1024', [])
    send_results = wechat_app.batch_send_message(to_conversations, '请查收文件', filepaths)
    logger.info('send_results: {}'.format(str(send_results)))


if __name__ == '__main__':
    test_send_message()
    # test_batch_send()
