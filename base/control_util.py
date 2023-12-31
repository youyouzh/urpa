import re

from base.log import logger
from uiautomation import Control


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
    child_control_name = child_control_name.lower() + 'control'

    same_control_index = 0
    for control in root_control.GetChildren():
        if control.ControlTypeName.lower() == child_control_name:
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
    control查找封装，使用类似css的方式，递归查找，这样比直接search快很多，书写也简单，其中 .. 表示查找父级
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
def control_click(control: Control, active: bool = False):
    if not control:
        logger.warning('The control is None, can not click.')
        return
    # 确保激活当前窗口后点击
    if active:
        control.Show()
    logger.info('click control: {}'.format(control))
    control.Click()
    # time.sleep(random.uniform(0.5, 1))
