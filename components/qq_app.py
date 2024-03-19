"""
QQ客户端托管代理
注意修改设置：
取消勾选: 主面包 -> 始终保持在其他窗口前端，
取消勾选: 状态 -> 自动回复，
取消勾选: 状态 -> 自动回复，
取消勾选: 登录 -> 总是打开登录提示，
取消勾选: 登录 -> 订阅“腾讯视频”，
取消勾选: 登录 -> 订阅“每日精选”，
取消勾选: 软件更新 -> 有更新时不要自动安装，
确定勾选：会话窗口 -> 使用多彩气泡
取消勾选：权限设置 -> 咨询提醒 -> 登录后显示“腾讯网迷你版”

首次启动或者登录QQ时绑定QQ主窗口，发送消息时有两种处理方式：
1. 首次发送时打开发送对话窗口，并绑定该窗口，后续都基于该对话窗口来发送消息，可以缓存元素查询结果
   同一台机多个QQ号，如果首次发送给相同昵称的消息，则会绑定错误
2. 每次发送消息时都打开发送对话窗口，发送完成后立即关闭，这时每次都要查找

适配PC QQ V9.4.6.27770版本，
"""
import re
import time
from typing import List

import uiautomation as auto
from uiautomation import Control

from base.control_util import select_control, control_click, check_control_exist, find_top_window_controls, \
    active_window, check_controls_exist
from base.exception import ControlInvalidException, MessageSendException
from base.log import logger
from base.util import win32_clipboard_text, win32_read_clipboard_text
from components.wechat_app import WechatApp, ControlTag, SendResult

auto.SetGlobalSearchTimeout(5)


class QQApp(WechatApp):

    def __init__(self, main_window=None, is_retain_conversation_window=True, is_search_in_conversation_window=False):
        self.qq_number = None
        self.conversation_window: Control = None   # 记录会话窗口
        self.is_retain_conversation_window: bool = is_retain_conversation_window        # 是否保持会话窗口存活而不关闭
        self.is_search_in_conversation_window: bool = is_search_in_conversation_window  # 是否在会话窗口中搜索
        self.open_conversations = set()     # 记录打开的会话列表
        super().__init__(main_window)

    @staticmethod
    def build_all_qq_apps(is_retain_conversation_window=True, is_search_in_conversation_window=False):
        # 处理同时打开多个微信的情况，找到指定微信名称的微信窗口
        qq_apps = []
        qq_window_controls = find_top_window_controls('QQ', 'TXGuiFoundation')
        for qq_window_control in qq_window_controls:
            qq_app = QQApp(qq_window_control, is_retain_conversation_window, is_search_in_conversation_window)
            qq_apps.append(qq_app)
        logger.info('检测到当前打开QQ窗口个数： {}'.format(len(qq_apps)))
        return qq_apps

    def _init_mian_window(self):
        if self.main_window:
            # 已经绑定main_window时不处理
            return True
        # 如果main_window为空，默认选第一个，首先检查有没有已打开的QQ窗口
        qq_window_controls = find_top_window_controls('QQ', 'TXGuiFoundation')
        if check_controls_exist(qq_window_controls):
            self.main_window = qq_window_controls[0]
            logger.info('初始化绑定QQ窗口：{}'.format(self.main_window))
            return True

        # 没有已打开QQ窗口则尝试点击一下QQ图标激活窗口
        self.click_toolbar_qq_icon()

        # 点击打开窗口后，再重新查找一次
        qq_window_controls = find_top_window_controls('QQ', 'TXGuiFoundation')
        if not check_controls_exist(qq_window_controls):
            logger.error('未检测到任何QQ窗口，绑定窗口失败')
            return False
        self.main_window = qq_window_controls[0]
        logger.info('初始化绑定QQ窗口成功：{}'.format(self.main_window))
        return True

    def click_toolbar_qq_icon(self):
        # 没有已打开QQ窗口则尝试点击一下QQ图标激活窗口
        qq_icon_controls = self.get_toolbar_qq_icon_controls()
        if not qq_icon_controls:
            logger.error('未检测到任何QQ客户端图标')
            return False
        logger.info('点击右下角的QQ图标，size: {}'.format(len(qq_icon_controls)))
        for qq_icon_control in qq_icon_controls:
            control_click(qq_icon_control)
            time.sleep(1)

    def _init_login_user_name(self, open_data_window=False):
        conversation_search_control = self._search_control(ControlTag.MAIN_CONVERSATION_SEARCH, with_check=False)
        if conversation_search_control:
            # 已经登录的情况可以找到会话搜索框
            login_user_name_control = select_control(conversation_search_control, 'p>p>pane>pane:2>pane')
            if not login_user_name_control:
                logger.error('未找到账号名称控件，无法设置当前窗口登录的QQ账号名')
                return False
            self.login_user_name = login_user_name_control.GetLegacyIAccessiblePattern().Description
            logger.info('初始化绑定QQ窗口和账号名: {}'.format(self.login_user_name))
            # 绑定QQ号需要打开资料卡
            if open_data_window:
                # 点击头像打开资料卡
                control_click(select_control(self.main_window, ':1>>:3>:2>>:1>-2>pane'))
                info_window = auto.WindowControl(searchDepth=1, Name='我的资料')
                if not check_control_exist(info_window):
                    logger.error('未找到资料卡窗口，初始化QQ号失败')
                    return False
                self.qq_number = info_window.EditControl(Name='帐号').GetLegacyIAccessiblePattern().Value
                # 关闭资料卡
                control_click(select_control(info_window, ':2>:1'))
        elif self.main_window:
            # 没有登录的话，点击触发登录确认按钮，并置顶窗口方便扫码
            self.active(force=True)   # 置顶，避免二维码被遮住
            self.ready_login_qr_code()

    def check_skip_update(self):
        alert_control = None
        for control in self.main_window.GetChildren():
            if control.Name == '提示':
                alert_control = control
                break
        if not alert_control:
            logger.info('没有任何提示框，不做处理')
            return False
        # 点击关闭按钮
        close_button = select_control(alert_control, ':2>button')
        logger.info('点击关闭提示框按钮')
        control_click(close_button)

    def get_toolbar_qq_icon_controls(self, filter_name='') -> List[Control]:
        # 获取右下角任务栏QQ图标列表
        toolbar_control = auto.PaneControl(searchDepth=1, Name='任务栏')
        if not check_control_exist(toolbar_control):
            logger.error('未找到【任务栏】控件')
            return []
        toolbar_control = toolbar_control.ToolBarControl(searchDepth=3, Name='用户提示通知区域')
        if not check_control_exist(toolbar_control):
            logger.error('未找到【用户提示通知区域】控件')
            return []
        qq_toolbar_name_regex = re.compile(r'QQ: (.+)\((\d+)?\)')

        qq_toolbar_controls = []
        for app_toolbar_control in toolbar_control.GetChildren():
            qq_toolbar_match_result = qq_toolbar_name_regex.match(app_toolbar_control.Name)
            if qq_toolbar_match_result and len(qq_toolbar_match_result.groups()) >= 2 \
                    and (not filter_name or self.login_user_name == filter_name):
                self.login_user_name = qq_toolbar_match_result.groups()[0]
                qq_toolbar_controls.append(app_toolbar_control)
        logger.info('检测到任务栏中有 {} 个QQ图标'.format(len(qq_toolbar_controls)))
        return qq_toolbar_controls

    def set_online_state(self):
        # QQ可能会掉线，手动设置上线
        qq_toolbar_controls = self.get_toolbar_qq_icon_controls(self.login_user_name)
        if not qq_toolbar_controls:
            raise ControlInvalidException('未找到匹配当前窗口的QQ图标：{}'.format(self.login_user_name))

        # 右键图标
        control_click(qq_toolbar_controls[0], right_click=True)
        # 查找弹窗的菜单窗口
        menu_control = auto.MenuControl(searchDepth=1, Name='TXMenuWindow')
        check_control_exist(menu_control, '未找到工具栏菜单: {}'.format(self.login_user_name))
        # 我在线上按钮
        online_button_control = menu_control.MenuItemControl(searchDepth=1, Name='我在线上')
        logger.info('点击【我在线上】按钮: {}'.format(self.login_user_name))
        control_click(online_button_control, with_exception_message='该账号未登录，没有【我在线上】按钮')

        # 检查是否需要重新登录
        re_login_windows = find_top_window_controls('重新登录', 'TXGuiFoundation')
        if not check_controls_exist(re_login_windows):
            logger.warning('检测到【重新登录】窗口，该账号需要重新登录: {}'.format(self.login_user_name))
            cancel_button = re_login_windows[0].ButtonControl(Name='取消')
            control_click(cancel_button)
            logger.info('点击【取消】关闭重新登录窗口')

    def exit_account(self, sub_menu='退出'):
        if not check_control_exist(self.main_window):
            logger.info('未检测到主窗口，不做处理')
            return False
        # 点击左下角菜单，选择退出账号
        menu_button = select_control(self._search_control(ControlTag.MAIN_CONVERSATION_SEARCH), 'p>p>:2>>>')
        if not check_control_exist(menu_button) or menu_button.GetLegacyIAccessiblePattern().Description != '主菜单':
            logger.warning('未找到主窗口的【主菜单】按钮')
            return False
        control_click(menu_button)

        # 主窗口菜单窗口
        menu_window = self.main_window.MenuControl(searchDepth=1, Name='TXMenuWindow')
        if not check_control_exist(menu_window):
            logger.info('未找到菜单窗口')
            return False

        # 退出账号按钮
        exit_account_button = menu_window.MenuItemControl(searchDepth=1, Name=sub_menu)
        control_click(exit_account_button, with_exception_message='未找到【{}】子菜单按钮'.format(sub_menu))
        logger.info('点击【{}】子菜单按钮退出当前账号: {}'.format(sub_menu, self.login_user_name))

        alert_window = find_top_window_controls('提示', 'TXGuiFoundation')
        if not check_controls_exist(alert_window):
            logger.info('未检测到【提示】窗口')
            return False
        alert_window = alert_window[0]
        confirm_button = alert_window.ButtonControl(Name='确定')
        control_click(confirm_button, '未找到【确定】按钮')
        logger.info('点击【确认】按钮关闭提示框')

    def switch_account(self):
        self.exit_account(sub_menu='切换帐号')

    @staticmethod
    def close_alert_window():
        alert_windows = find_top_window_controls('提示', 'TXGuiFoundation')
        if check_controls_exist(alert_windows):
            active_window(alert_windows[0], force=True)
            logger.info('点击【取消】按钮关闭提示框')
            control_click(alert_windows[0].ButtonControl(Name='取消'), '未找到提示框中的【取消】按钮')

        alert_windows = find_top_window_controls('下线通知', 'TXGuiFoundation')
        if check_controls_exist(alert_windows):
            active_window(alert_windows[0], force=True)
            logger.info('点击【确定】按钮关闭【下线通知】提示框')
            control_click(alert_windows[0].ButtonControl(Name='确定'), '未找到提示框中的【确定】按钮')

    def close_login_window(self):
        close_button = select_control(self.main_window, ':1>>:1>>button:2')
        if check_control_exist(close_button) and close_button.GetLegacyIAccessiblePattern().Description == '关闭':
            logger.info('点击【关闭】QQ登录窗口')
            control_click(close_button)

    def ready_login_qr_code(self):
        # 点击验证码登录
        login_button = self.main_window.ButtonControl(searchDepth=6, Name='登录')
        if check_control_exist(login_button):
            logger.info('检测到【登录】按钮，点击使用二维码登录')
            login_qr_button = login_button.GetPreviousSiblingControl()
            control_click(login_qr_button)

        # 点击刷新二维码
        return_button = self.main_window.ButtonControl(searchDepth=7, Name='返回')
        if check_control_exist(return_button):
            logger.info('检测到【返回】按钮，点击刷新二维码')
            qr_code_control = return_button.GetPreviousSiblingControl().GetPreviousSiblingControl().GetChildren()[0]
            # 移动鼠标悬浮等待二维码移动后在点击
            qr_code_control.MoveCursorToMyCenter()
            time.sleep(1)
            qr_code_control = return_button.GetPreviousSiblingControl().GetPreviousSiblingControl().GetChildren()[0]
            control_click(qr_code_control)

    def _search_control(self, tag: ControlTag, use_cache=True, with_check=True) -> Control | None:
        if not self.is_retain_conversation_window and tag in [ControlTag.MESSAGE_LIST, ControlTag.MESSAGE_INPUT]:
            # 非绑定模式下关于会话窗口的查找不能使用缓存
            use_cache = False
        return super()._search_control(tag, use_cache, with_check)

    def _search_control_by_tag(self, tag: ControlTag, use_cache=True, with_check=True):
        find_control = None
        if tag == ControlTag.MAIN_CONVERSATION_SEARCH:
            find_control = self.main_window.EditControl(searchDepth=6, Name='搜索：联系人、群聊、企业')
        elif tag == ControlTag.MAIN_CONVERSATION_SEARCH_RESULT:
            anchor_control = self._search_control(ControlTag.MAIN_CONVERSATION_SEARCH, use_cache, with_check)
            if anchor_control:
                find_control = select_control(anchor_control, 'p>p>:3>:1>>>:1')
        elif self.conversation_window:
            if tag == ControlTag.CONVERSATION_SEARCH:
                find_control = self.conversation_window.EditControl(searchDepth=3, Name='搜索：联系人、群聊、企业')
            elif tag == ControlTag.CONVERSATION_SEARCH_RESULT:
                find_control = select_control(self.conversation_window, ':5>>>>:1')
            elif tag == ControlTag.MESSAGE_LIST:
                # 无法获取具体的消息控件列表
                find_control = self.conversation_window.ListControl(Name='消息')
            elif tag == ControlTag.MESSAGE_INPUT:
                find_control = self.conversation_window.EditControl(Name='输入')
        return find_control

    def _bind_conversation_control(self, window_control):
        self.conversation_window = window_control
        self.open_conversations = set()  # 清空打开的会话列表
        # 清空会话窗口的搜索换成
        self.cached_control[ControlTag.MESSAGE_INPUT] = None
        self.cached_control[ControlTag.MESSAGE_LIST] = None
        self.cached_control[ControlTag.CONVERSATION_SEARCH] = None
        self.cached_control[ControlTag.CONVERSATION_SEARCH_RESULT] = None

    def _get_conversation_active_title(self, conversation_window=None):
        # 默认使用缓存的会话窗口
        conversation_window = conversation_window if conversation_window else self.conversation_window
        title_selector = ':1>>>>'
        conversation_title_control = select_control(conversation_window, title_selector)
        if not conversation_title_control or not conversation_title_control.Name:
            # 有时候切换比较慢，还没加载会导致title为空，此时尝试延时后二次获取
            logger.info('延迟3秒重复查找会话标题控件')
            time.sleep(3)
            conversation_title_control = select_control(conversation_window, title_selector)
            if not conversation_title_control:
                # 首次从主窗口搜索切换过来时标题控件加载比较慢，如果为了效率可以关闭会话窗口驻留，跳过窗口校验
                logger.info('延迟5秒再次重复查找会话标题控件')
                time.sleep(5)
                conversation_title_control = select_control(conversation_window, title_selector)
        conversation_title = conversation_title_control.Name if conversation_title_control else '^^^^^^没查到会话标题'
        return conversation_title

    def _search_switch_conversation(self, conversation: str):
        # 切换到指定会话窗口，有两种方式，一种是从QQ主窗口搜索切换，另一种是直接在对话框搜索切换
        logger.info('开始处理切换会话: {}'.format(conversation))

        # 搜索会话然后打开会话窗口，搜索结果中包含昵称、QQ号、备注
        match_names = self._search_switch_conversation_by_window(conversation)

        # 查询会话窗口列表，多个QQ号时可能有多个会话窗口，校验和绑定会话窗口
        for window_control in auto.GetRootControl().GetChildren():
            # 通过一般属性过滤非会话窗口
            if window_control.ClassName != 'TXGuiFoundation' or 'QQ' in window_control.Name:
                continue

            # 会话窗口的Name一般为备注、昵称，单个会话窗口时匹配Name即可,，如果是多个会话堆叠的话，叫做xxx等x个会话
            # 注意：新版本可能需要使用 _get_conversation_active_title 来获取激活会话窗口标题
            active_conversation_title_name = re.sub(r'等\d+个会话', '', window_control.Name)

            if active_conversation_title_name in match_names:
                if not self.conversation_window:
                    # 如果没有绑定过窗口（一般是首次打开会话窗口或者会话窗口非驻留模式），则需要绑定窗口
                    self._bind_conversation_control(window_control)
                    break

        if not self.conversation_window:
            raise ControlInvalidException('未找到匹配的会话窗口. match_names: {}'.format(match_names))

        self.open_conversations.add(conversation)
        return self.conversation_window

    def _search_switch_conversation_by_window(self, conversation: str) -> List[str]:
        logger.info('搜索和切换会话列表： {}'.format(conversation))
        # 打开会话框搜索并且已经打开多个会话框时才会在会话框中搜索
        search_in_conversation = len(self.open_conversations) > 2 if self.is_search_in_conversation_window else False
        active_window(self.conversation_window if search_in_conversation else self.main_window)
        search_control = self._search_control(ControlTag.CONVERSATION_SEARCH if search_in_conversation
                                              else ControlTag.MAIN_CONVERSATION_SEARCH)
        control_click(search_control)
        time.sleep(0.7)
        search_control.SendKeys('{Ctrl}a')  # 避免还有旧的搜索
        win32_clipboard_text(conversation)
        search_control.SendKeys('{Ctrl}v')

        # 检查搜索结果列表
        search_conversation_controls = []
        search_result_root_control = self._search_control(ControlTag.CONVERSATION_SEARCH_RESULT if search_in_conversation
                                                          else ControlTag.MAIN_CONVERSATION_SEARCH_RESULT)
        for pane_control in search_result_root_control.GetChildren():
            result_type_control = select_control(pane_control, '>')
            # 依次将每一项的搜索结果添加，比如好友、群聊
            if result_type_control and result_type_control.Name in ['好友', '群聊']:
                search_conversation_controls.extend(result_type_control.GetParentControl().GetChildren()[1:])

        if not search_conversation_controls:
            raise ControlInvalidException('未搜索到该会话-搜索结果为空：{}'.format(conversation))

        item_match_regex = re.compile(r'^([^(]+)(\((.+)\))? (\d+)$')
        for item_control in search_conversation_controls:
            # 匹配备注名、昵称、QQ号
            item_name = item_control.GetFirstChildControl().Name
            item_match_result = item_match_regex.match(item_name)
            if (not item_match_result or not item_match_result.groups()) and item_name != self.login_user_name:
                raise ControlInvalidException('未匹配的会话名称-未知匹配项： {}'.format(item_name))

            # 分别匹配备注、昵称、QQ号
            for match_item in item_match_result.groups():
                # 分别匹配备注、昵称、QQ号
                if not match_item:
                    continue
                if match_item == conversation:
                    logger.info('匹配并切换会话，item_name: {}, conversation: {}'.format(item_name, conversation))
                    item_control.DoubleClick()
                    time.sleep(0.8)
                    return list(item_match_result.groups())
        raise ControlInvalidException('未搜索到该会话-未找到匹配项：{}'.format(conversation))

    def _get_message_item_controls(self, filter_time=True) -> List[Control]:
        # QQ获取消息列表暂时有问题
        return []

    def check_last_message_match(self, message: str) -> bool:
        # 点击聊天消息区域
        message_list_control = self._search_control(ControlTag.MESSAGE_LIST)
        control_click(message_list_control)

        # Ctrl+A C复制所有消息
        message_list_control.SendKeys('{Ctrl}a')
        message_list_control.SendKeys('{Ctrl}c')

        # 读取剪贴板消息
        all_message_text = win32_read_clipboard_text()
        # 去掉Tim沟通提示
        all_message_text = all_message_text.replace('正在和企业用户沟通，为了提供更好服务，企业可能会保存与你的沟通内容。', '')
        # message_items = all_message_text.split('\r\n\r\n')[:-1]  # 拆分成每一条消息，最后一条消息为空
        all_message_text = all_message_text.replace('\r\n', '').replace('\n', '')
        # 去掉换行进行匹配
        message = message.replace('\r\n', '').replace('\n', '')
        return all_message_text.endswith(message)

    def close_conversation_window(self, force=False):
        if not check_control_exist(self.conversation_window):
            logger.error('会话窗口不存在。login_user_name: {}'.format(self.login_user_name))
            return False
        if force:
            self.conversation_window.SendKeys('{Alt}{F4}')
            self.open_conversations.clear()
        elif not self.is_retain_conversation_window:
            # 非绑定模式，关闭会话窗口
            self.conversation_window.SendKeys('{Ctrl}w')
            self.open_conversations.clear()
        # 清空搜索换成
        self.cached_control = {}

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
        if not text and not filepaths:
            raise MessageSendException('发送内容为空，请检查参数')
        send_results = []
        self.active(force=True)

        # 逐个发送，如果有一个发送失败，则全失败
        for to_conversation in to_conversations:
            try:
                self._search_switch_conversation(to_conversation)
                # 首先检查前置消息是否匹配
                if check_pre_message and not self.check_last_message_match(check_pre_message):
                    logger.error('前置消息不匹配. check_pre_message: {}'.format(check_pre_message))
                    send_results.append(SendResult.fail(to_conversation, '前置消息不匹配：' + check_pre_message))
                    self.close_conversation_window()
                    continue

                send_message_controls = []
                if filepaths:
                    send_message_controls.extend(self.send_file_message(filepaths, check_send_success=False))
                if text:
                    send_message_controls.extend(self.send_text_message(text, check_send_success=False))
                self.close_conversation_window()

                # 严格判断是否发送成功
                if text and not self.check_last_message_match(text):
                    # 可能离线状态就会发生失败
                    send_results.append(SendResult.fail(to_conversation, '校验发送消息失败请检查QQ状态'))
                else:
                    send_results.append(SendResult.success(to_conversation))
            except ControlInvalidException as exception:
                logger.warning('conversation send exception. conversation: {}'.format(to_conversation), stack_info=True)
                send_results.append(SendResult.fail(to_conversation, exception.message))

        send_fail_results = [x for x in send_results if not x.is_success]
        if send_fail_results and len(send_fail_results) == len(to_conversations):
            # 全部发送失败，则直接抛出异常结束
            raise MessageSendException('全部消息发送失败：{}...'
                                       .format('、'.join([x.error_message for x in send_fail_results[:2]])))
        return send_results


def simple_test():
    qq_app = QQApp()
    qq_app.active()
    # qq_app.click_toolbar_qq_icon()
    # qq_app.set_online_state()
    # qq_app.ready_login_qr_code()
    # qq_app.set_online_state()
    # qq_app.exit_account()
    # qq_app.switch_account()
    qq_app.close_login_window()
    qq_app.close_alert_window()


def benchmark_test():
    qq_apps = QQApp.build_all_qq_apps(is_retain_conversation_window=True)
    qq_app_y = None
    qq_app_y_number = '1947828953'
    qq_app_y_username = '咏春 叶问'
    qq_app_u = None
    qq_app_u_number = '2317373692'
    qq_app_u_username = 'uusama'
    for qq_app in qq_apps:
        if qq_app.login_user_name == qq_app_u_username:
            qq_app_u = qq_app
        if qq_app.login_user_name == qq_app_y_username:
            qq_app_y = qq_app

    qq_app_y.batch_send_message([qq_app_u_number], '测试QQ号私发消息不检查前置')
    qq_app_y.batch_send_message(['★絮&媛☆^'], '测试通过昵称发送给另外一个好友')
    qq_app_y.batch_send_message(['xxxx', '文件交流', qq_app_u_number], '测试检查前置，测试无好友，测试用昵称发送\n换行测试',
                                check_pre_message='测试QQ号私发消息不检查前置')
    qq_app_y.batch_send_message([qq_app_u_number], '测试前置消息包含换行',
                                check_pre_message='测试检查前置，测试无好友，测试用昵称发送\n换行测试')

    qq_app_u.batch_send_message([qq_app_y_number], '测试使用另外一个QQ号发送消息-无前置检查')
    qq_app_y.batch_send_message([qq_app_u_number], '测试接收消息检查并发送消息',
                                check_pre_message='测试使用另外一个QQ号发送消息-无前置检查')
    qq_app_u.batch_send_message([qq_app_y_number], '二次测试另外一个QQ号接收消息并发送',
                                check_pre_message='测试接收消息检查并发送消息')
    qq_app_u.batch_send_message(['307943267'], '测试另一个号发送群聊消息')
    qq_app_u.batch_send_message(['222222222'], '测试发送给非好友')


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    # simple_test()
    benchmark_test()
