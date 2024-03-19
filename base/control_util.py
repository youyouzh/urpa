import random
import re
import time
from typing import List

import uiautomation as auto
from uiautomation import Control

from base.exception import ControlInvalidException
from base.log import logger


def select_control_by_tree(root_control: Control, select_items: list) -> Control | None:
    """
    递归获取control，示例：pane.pane.pane:2.edit
    """
    # root为空则没找到
    if not root_control:
        return None

    # root不为空，且查询数量为0，则返回
    if len(select_items) == 0:
        return root_control

    select_item = select_items.pop(0)

    # 查找父节点
    if select_item == 'p' or select_item == '.':
        return select_control_by_tree(root_control.GetParentControl(), select_items)

    # 没有子节点返回空
    if not root_control.GetChildren():
        return None

    child_control_index = 0
    if ':' in select_item:
        # 子节点索引解析
        split_select_items = select_item.split(':')
        child_control_name = split_select_items[0]
        child_control_index = int(split_select_items[1])
    else:
        child_control_name = select_item

    # child_control_name 控件简化信息补全
    child_control_name = child_control_name.strip().lower() + 'control'

    same_control_index = 0
    for control in root_control.GetChildren():
        if control.ControlTypeName.lower() == child_control_name or child_control_name == 'control':
            # 和索引相同则命中，递归查找
            if child_control_index == same_control_index:
                # logger.info('attach control. sub_select_items: {}, control: {}'
                #             .format(child_select_items, control.ControlTypeName))
                return select_control_by_tree(control, select_items)

            # 相同的control类型计数 +1
            same_control_index += 1
    return None


def select_control(root_control: Control, selector: str) -> Control | None:
    """
    control查找封装，使用特定的树路径查找写法，递归查找，这样比直接search快很多，书写也简单，具体语法如下：
    > 表示取父节点或者子节点，>p 表示父节点，其他则表示子节点
    > 后面跟控件类型的小写，如pane表示找PaneControl子节点，text表示TextControl子节点，p表示父节点，留空表示匹配所有子节点
    - 后面跟数字，表示重复多少次下一个节点查找，pane-2表示查找当前节点下的PaneControl子节点下的PaneControl子节点重复2次
    : 后面跟数字，表示查找第几个字节点，索引从0开始，pane:1表示查找当前节点下的第2个PaneControl子节点
    示例： select_control(wechat_window, 'pane:1>pane>pane>edit')
    :param root_control 基准control
    :param selector 选择器
    """
    child_select_items = re.split(r' *> *', selector)
    # 处理多级简写
    expend_items = []
    for item in child_select_items:
        # 重复标识，p-4 处理为 [p, p, p, p]
        if '-' in item:
            item, times = item.split('-')
            for i in range(int(times)):
                expend_items.append(item)
        else:
            expend_items.append(item)
    return select_control_by_tree(root_control, expend_items)


def select_parent_control(root_control: Control, level: int) -> Control | None:
    """
    快速获取多层级父节点
    :param root_control 基准control
    :param level 父节点层级
    """
    if not root_control:
        return None
    parent_control = root_control
    for index in range(level):
        parent_control = parent_control.GetParentControl()
    return parent_control


# 控件点击，需要加入随机演示，会提前确保窗口 active
def control_click(control: Control, right_click=False, with_exception_message='', check_name=''):
    """
    点击某个空间
    :params control: 控件，可为空
    :params right_click: 鼠标右键点击
    :params exception_message: 控件为空时输出的异常信息，如果为空则不抛出异常
    :params check_name: 检查名称是否匹配，只适用于非空检查
    """
    if not control or not control.Exists(1, 1) or (check_name and control.Name != check_name):
        logger.error('The control is not exist, can not click.')
        if with_exception_message:
            raise ControlInvalidException(with_exception_message)
        return False
    # 确保激活当前窗口后点击，其中会检测当前窗口如果时候TopLevel不会进行操作
    # self.active(force=True)
    logger.info('click control name: {} position: {}'.format(control.Name, control.BoundingRectangle))
    if right_click:
        control.RightClick()
    else:
        control.Click()
    time.sleep(random.uniform(0.1, 0.3))
    return True


def check_controls_exist(controls: List[Control], with_exception_message=''):
    if not controls:
        # 空数组直接当成空控件处理
        return check_control_exist(None, with_exception_message)
    result = True
    for control in controls:
        # 检查每个控件是否存在，只要有一个控件不存在则返回False
        result = result and check_control_exist(control, with_exception_message)
    return result


def check_control_exist(control: Control | None, with_exception_message=''):
    """
    检查控件是否存在
    :params control: 控件，可为空
    :params with_exception_message: 异常消息，如果不为空则会抛出异常
    """
    if not control or not control.Exists(1, 1):
        if with_exception_message:
            raise ControlInvalidException(with_exception_message)
        return False
    return True


def find_top_window_controls(name='', class_name='', with_exception_message='',
                             root_control=None) -> List[Control]:
    """
    查找顶层窗口列表
    :params name: 匹配窗口控件名称，如果为空则不匹配
    :params class_name: 匹配窗口控件类名称，如果为空则不匹配
    :params root_control: 根控件，默认是桌面
    :return 符合条件的窗口控件列表
    """
    if not root_control:
        root_control = auto.GetRootControl()
    top_window_controls = []
    for app_control in root_control.GetChildren():
        # 匹配微信主窗口，不通版本可能是PaneControl，可能是WindowControl
        if (not name or app_control.Name == name) and (not class_name or class_name in app_control.ClassName):
            top_window_controls.append(app_control)
    logger.info('查找到当前顶层窗口数量为：{}，name: {}, class_name: {}'.format(len(top_window_controls), name, class_name))
    if with_exception_message and not top_window_controls:
        raise ControlInvalidException(with_exception_message)
    return top_window_controls


def active_window(window_control: Control, force=False, wait_time=0.5):
    """
    激活窗口
    :params window_control: 需要激活窗口控件
    :params force: 是否强制，如果不为True，不会抢占
    :params wait_time: 激活窗口后的等待时间
    """
    # 发送按键时通过self.main_window.SendKeys，不会有问题，需先SendKeys后才能进行点击操作
    if window_control and window_control.IsTopLevel():
        handle = window_control.NativeWindowHandle
        if auto.IsIconic(handle):
            ret = auto.ShowWindow(handle, auto.SW.Restore)
        elif not auto.IsWindowVisible(handle):
            ret = auto.ShowWindow(handle, auto.SW.Show)
        if force:
            # 一个trick，SetForegroundWindow 并不总是有效，给窗口发送一个按键事件，增加生效成功率
            # 参考： https://stackoverflow.com/questions/19136365/win32-setforegroundwindow-not-working-all-the-time
            window_control.SendKeys('{Alt}')
        ret = auto.SetForegroundWindow(handle)
        time.sleep(wait_time)
        return ret
    return False
