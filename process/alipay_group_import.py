"""
阿里支付宝粉丝群导入
"""
import copy
import json
import os.path
import re
import sys
import time

from base.log import logger
from base.xlsx_util import read_excel_to_dict
from urllib import parse
import requests

PUBLIC_QUERY_PARAM = ''


# 获取当前文件夹
def get_current_dir():
    if getattr(sys, 'frozen', False):  # 是否Bundle Resource
        # exe运行时获取当前文件夹
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        # 不能使用 __file__
        return os.path.dirname(os.path.abspath(sys.argv[0]))


def load_header(header_path: str):
    if not os.path.isfile(header_path):
        logger.error('请求头配置文件不存在: {}'.format(header_path))
        return []
    header = {}
    with open(header_path, 'r', encoding='utf-8') as fin:
        lines = fin.readlines()
        for i in range(0, len(lines), 2):
            header[lines[i].replace(':\n', '')] = lines[i+1].replace('\n', '')
    logger.info('header: {}'.format(header))
    return header


def load_header_from_curl_bash(header_path: str):
    if not os.path.isfile(header_path):
        logger.error('请求头配置文件不存在: {}'.format(header_path))
        return []
    header = {}
    with open(header_path, 'r', encoding='utf-8') as fin:
        lines = fin.readlines()
        # 提取ctoken
        ctoken_regex = re.compile('ctoken=([^&]+)&')
        ctoken_search_result = ctoken_regex.search(lines[0])
        if not ctoken_search_result:
            logger.error('请检查curl_bash.txt文件格式，无法解析ctoken信息')
            return []
        global PUBLIC_QUERY_PARAM
        PUBLIC_QUERY_PARAM = 'ctoken={}&_input_charset=utf-8'.format(ctoken_search_result.groups()[0])

        # 提取请求头
        for line in lines:
            if '-H ' not in line:
                continue
            line = line.replace('  -H \'', '').replace('\' \\', '')
            header_kv = line.split(':')
            if len(header_kv) < 2:
                logger.info('not format header: {}'.format(line))
                continue
            header[header_kv[0].strip()] = header_kv[1].strip()
    logger.info('header: {}'.format(header))
    return header


HEADERS = load_header_from_curl_bash('curl_bash.txt')
GROUP_QUERY_API = 'https://v.alipay.com/api/group/pagingGroupConfig?' + PUBLIC_QUERY_PARAM \
                  + '&_ksTS={timestamp}&page={page}&size={pageSize}&publicId=2018061160393077&groupId=&groupTypeId='
GROUP_MEMBER_MODIFY_SAVE_API = 'https://v.alipay.com/api/group/updateGroupAdmin?' + PUBLIC_QUERY_PARAM


def query_group_info(page: int = 1, page_size: int = 10):
    params = {
        'timestamp': str(int(time.time() * 1000)),
        'page': page,
        'pageSize': page_size
    }
    query_url = GROUP_QUERY_API.format(**params)
    response = requests.get(query_url, headers=HEADERS)
    if response.status_code != 200:
        logger.error('请求群信息失败，范围HTTP状态码: {}'.format(response))
        return []
    if '<!DOCTYPE HTML>' in response.text:
        logger.warning('登录信息过期，请重新登录并更新curl_bash.txt文件')
        return []
    data = json.loads(response.text)
    data = data['groupConfigVOList']
    logger.info('request group count: {}'.format(len(data)))
    return data


def query_all_group_info() -> dict:
    page_size = 10
    max_page = 50
    cache_file_path = r'group-info.json'
    if os.path.isfile(cache_file_path):
        logger.info('load all group info from cache file: {}'.format(cache_file_path))
        with open(cache_file_path, 'r', encoding='utf-8') as fp:
            return json.load(fp)
    group_infos = []

    for page in range(max_page):
        data = query_group_info(page, page_size)
        if not data:
            logger.info('request data is empty. break. page: {}'.format(page))
            break
        group_infos.extend(data)
        logger.info('request group info finish. page: {}, page_size: {}'.format(page, page_size))

    if not group_infos:
        logger.warning('爬取群信息失败')
        return {}

    # 转成map映射，方便后续操作
    group_info_map = {}
    for group_info in group_infos:
        group_info_map[group_info['groupId']] = group_info
    with open(cache_file_path, 'w', encoding='utf-8') as file_handler:
        json.dump(group_info_map, file_handler, ensure_ascii=False, indent=4)
    return group_info_map


def save_group_admin(id, group_id, add_user_ids: list, remove_user_ids: list):
    params = {
        "id": id,
        "groupId": group_id,
        "addUserIdList": json.dumps(add_user_ids),
        "removeUserIdList": json.dumps(remove_user_ids)
    }
    logger.info('模拟请求添加群管理员，参数： {}'.format(params))
    headers = copy.deepcopy(HEADERS)
    headers['Content-Type'] = 'application/x-www-form-urlencoded; charset=utf-8'
    data = parse.urlencode(params, encoding='utf-8')
    response = requests.post(GROUP_MEMBER_MODIFY_SAVE_API, data=data, headers=headers)
    if response.status_code != 200:
        logger.error('response code is not 200: {}'.format(response))
        return False
    if '<!DOCTYPE HTML>' in response.text:
        logger.warning('登录信息过期，请更新登录cookie')
        return False

    result = json.loads(response.text)
    if 'code' not in result or result['code'] != 0:
        logger.error('request save return code is not success: {}'.format(response))
        return False
    return True


def import_uid_to_group(id, group_id, user_ids: list):
    # TODO: 检测群人数是否已满
    # 先设置为管理员实现拉人入群
    logger.info('处理拉群，群id: {}, 显示群id: {}, uid列表: {}'.format(id, group_id, user_ids))
    result = save_group_admin(id, group_id, user_ids, [])
    # 再移除管理员
    if result:
        result = save_group_admin(id, group_id, [], user_ids)
    logger.info('拉群结果： {}，群id: {}, 显示群id: {}, uid列表: {}'.format(result, id, group_id, user_ids))
    return result


def read_import_uid_from_xlsx(xlsx_path: str) -> dict:
    if not os.path.isfile(xlsx_path):
        logger.error('uid导入表格文件不存在： {}'.format(xlsx_path))
        return {}
    key_map = {
        '昵称': 'nickname',
        'UID账号': 'uid',
        '加入粉丝群名称': 'group_name',
        '加入粉丝群ID': 'group_show_id'
    }
    uid_records = read_excel_to_dict(xlsx_path, key_map)
    logger.info('读取导入uid表格成功，uid记录条数： {}'.format(len(uid_records)))
    # 对加入同一个群的用户进行合并
    uid_record_map = {}
    for uid_record in uid_records:
        group_show_id = uid_record['group_show_id']
        if group_show_id not in uid_record_map:
            # 初始化
            uid_record_map[group_show_id] = []
        uid_record_map[group_show_id].append(uid_record)
    logger.info('转换成按照群id来处理，群数量： {}'.format(len(uid_record_map.keys())))
    return uid_record_map


def run_import_uid():
    current_dir = get_current_dir()
    # 处理产品xlsx文件路径
    xlsx_path = ''
    for file in os.listdir(current_dir):
        if 'xlsx' in file and '~$' not in file:
            xlsx_path = os.path.join(current_dir, file)
            break
    if not xlsx_path:
        logger.error('未在当前文件夹下找到群uid表格xlsx文件')
        return []
    logger.info('识别到群uid表格文件： {}'.format(xlsx_path))
    uid_record_map = read_import_uid_from_xlsx(xlsx_path)
    group_info_map = query_all_group_info()
    for group_show_id in uid_record_map.keys():
        uid_records = uid_record_map.get(group_show_id)
        if not uid_records:
            continue

        if group_show_id not in group_info_map:
            logger.error('未找到该群id： {}'.format(group_show_id))
            continue

        group_info = group_info_map[group_show_id]
        add_user_ids = [uid_record.get('uid') for uid_record in uid_records]
        logger.info('------> 处理拉群，群名称： {}, id: {}, 拉入用户id列表： {}'
                    .format(uid_records[0]['group_name'], group_info['groupId'], add_user_ids))
        import_uid_to_group(group_info['id'], group_info['groupId'], add_user_ids)
        # 更新结果并缓存，已经拉取过的不做处理


# 打包： pyinstaller -F alipay_group_import.py -p ../ --exclude-module pywin32 --exclude-module gevent --exclude-module flask --exclude-module APScheduler --exclude-module uiautomation
# 打包： pyinstaller alipay_group_import.spec
if __name__ == '__main__':
    run_import_uid()
    input("按任意键结束...")
