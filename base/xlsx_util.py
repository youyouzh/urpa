import os
import openpyxl

from typing import List, Dict
from base.log import logger


def read_sheet_to_dict(sheet, key_map: dict = None) -> List[Dict]:
    # 获取第一行数据，即列名
    keys = [cell.value for cell in sheet[1]]
    # 如果给定 key 映射，则进行转换
    if key_map:
        keys = [key_map.get(key, key) for key in keys]

    # 获取除第一行外的所有行数据
    rows = [[cell.value for cell in row] for row in sheet.iter_rows(min_row=2)]
    # 过滤全空的数据行
    rows = list(filter(lambda row: [cell for cell in row if cell], rows))
    # 将列名和值组合成字典
    return [dict(zip(keys, row)) for row in rows]


def read_excel_to_dict(xlsx_path: str, key_map: dict = None) -> List[Dict]:
    if not os.path.isfile(xlsx_path):
        logger.error('The xlsx file is not exist: {}'.format(xlsx_path))
        return []
    # 加载工作簿
    workbook = openpyxl.load_workbook(xlsx_path)
    # 获取了活动工作表, sheet = workbook['Sheet2']
    # 获取第一个工作表
    sheet = workbook.active
    return read_sheet_to_dict(sheet, key_map)


if __name__ == '__main__':
    result = read_excel_to_dict(r'E:\需求相关资料\需求-GM邮件发送\GM清单 的副本.xlsx')
    logger.info('result: {}'.format(result))
