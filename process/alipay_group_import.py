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
        ctoken_regex = re.compile(r'ctoken=([\w-]+)')
        ctoken_search_result = ctoken_regex.search(lines[0])
        if not ctoken_search_result:
            logger.error('请检查curl_bash.txt文件格式，无法解析ctoken信息')
            return []
        global PUBLIC_QUERY_PARAM
        PUBLIC_QUERY_PARAM = 'ctoken={}&_input_charset=utf-8'.format(ctoken_search_result.groups()[0])

        # 提取请求头
        match_header_regex = re.compile(r'^ +-H \'([\w-]+): ([^\'^]+)\'')
        for line in lines:
            if '-H' not in line:
                continue
            header_match_result = match_header_regex.search(line)
            if not header_match_result or not header_match_result.groups():
                logger.info('not format header: {}'.format(line))
                continue
            header_key = header_match_result.groups()[0]
            header_value = header_match_result.groups()[1]
            header[header_key.strip()] = header_value.strip()
    logger.info('header: {}'.format(header))
    return header


HEADERS = load_header_from_curl_bash('curl_bash.txt')
GROUP_QUERY_API = 'https://v.alipay.com/api/group/pagingGroupConfig?' + PUBLIC_QUERY_PARAM \
                  + '&_ksTS={timestamp}&page={page}&size={pageSize}&publicId=2018061160393077&groupId=&groupTypeId='
GROUP_MEMBER_MODIFY_SAVE_API = 'https://v.alipay.com/api/group/updateGroupAdmin?' + PUBLIC_QUERY_PARAM
GROUP_MEMBER_QUERY_API = 'http://cp.zhisheng.com/chatbc/group_member_list.json?' + PUBLIC_QUERY_PARAM
GROUP_ORG_LIST_API = 'http://cp.zhisheng.com/chatbc/org_list.json?' + PUBLIC_QUERY_PARAM
ORG_GROUP_LIST_API = 'http://cp.zhisheng.com/chatbc/group_list.json?' + PUBLIC_QUERY_PARAM


def request_group_info_by_page(page: int = 1, page_size: int = 10):
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


def request_all_group_info() -> dict:
    page_size = 10
    max_page = 50
    cache_file_path = r'cache\group-info-map.json'
    if os.path.isfile(cache_file_path):
        logger.info('load all group info from cache file: {}'.format(cache_file_path))
        with open(cache_file_path, 'r', encoding='utf-8') as fp:
            return json.load(fp)
    group_infos = []

    for page in range(max_page):
        data = request_group_info_by_page(page, page_size)
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


# 通用群管理后台请求接口
def request_group(params: dict, url: str):
    data = {
        'param': json.dumps(params),
        'mngSource': ''
    }
    data = parse.urlencode(data, encoding='utf-8')
    response = requests.post(url, data=data, headers=HEADERS)
    if response.status_code != 200:
        logger.error('request group failed，response: {}, url: {}'.format(response, url))
        return False
    result = json.loads(response.text)
    if 'data' not in result:
        logger.warning('request group data empty，response: {}, url: {}'.format(response.text, url))
        return False
    return result['data']


# 模拟请求查询群组的所有成员
def request_group_all_members(org_id, group_id):
    max_page = 50
    page_size = 100

    group_members = []
    for page_index in range(1, max_page):
        params = {
            "tenantId": "6",
            "orgId": str(org_id),
            "groupId": str(group_id),
            "pageNo": page_index,
            "pageSize": page_size
        }
        data = request_group(params, GROUP_MEMBER_QUERY_API)
        if not data or 'list' not in data or not data['list']:
            logger.warning('request group member by page end，goto next group. group_id: {}, page_index: {}'
                           .format(group_id, page_index))
            break
        group_members.extend(data['list'])
    logger.info('request all group member finish. group_id: {}, member size: {}'
                .format(group_id, len(group_members)))
    return group_members


# 模拟请求查询所有群的所有成员信息
def request_all_group_all_members():
    cache_file_path = r'cache\group-member-map.json'
    if os.path.isfile(cache_file_path):
        logger.info('从缓存中加载群成员信息: {}'.format(cache_file_path))
        with open(cache_file_path, 'r', encoding='utf-8') as fp:
            return json.load(fp)

    group_orgs = request_all_group_orgs_with_groups()
    group_member_map = {}
    for group_org in group_orgs:
        org_id = group_org['orgId']
        groups = group_org['groups']
        logger.info('begin request group members for org: {}'.format(org_id))
        for group in groups:
            group_id = group['groupId']
            group_members = request_group_all_members(org_id, group_id)
            if not group_members:
                logger.error('request group members is empty。group_id: {}'.format(group_id))
                continue
            logger.info('request group members success，group_id: {}， member size: {}'
                        .format(group_id, len(group_members)))
            group_member_map[group_id] = group_members
        logger.info('end request group members for org: {}'.format(org_id))

        with open(cache_file_path, 'w', encoding='utf-8') as file_handler:
            json.dump(group_member_map, file_handler, ensure_ascii=False, indent=4)
    return group_member_map


# 提取加入重复群的用户列表
def extract_repeat_group_user():
    group_member_map = request_all_group_all_members()
    group_info_map = request_all_group_info()
    user_group_map = {}
    # 检查每个用户加入的群列表
    for group_id, group_members in group_member_map.items():
        for group_member in group_members:
            user_id = group_member['userId']
            if user_id not in user_group_map:
                user_group_map[user_id] = []
            user_group_map[user_id].append(group_member)

    user_groups = []
    group_user_ids = []
    for user_id, group_members in user_group_map.items():
        if len(group_members) <= 1:
            continue
        group_member = group_members[0]
        user_group = {
            'userId': user_id,
            'nickName': group_member['nickName'],
            'loginId': group_member['loginId'],
            'groups': [],
            'groupNames': [],
            'groupCount': len(group_members)
        }
        group_user_ids.append(user_id)
        for group_member in group_members:
            group_id = group_member['groupId']
            if group_id not in group_info_map:
                logger.warning('group info is not exist. groupId: {}'.format(group_id))
                continue
            user_group['groups'].append({
                'groupId': group_id,
                'role': group_member['role'],
                'groupName': group_info_map[group_id]['groupName'],
                'groupSize': group_info_map[group_id]['groupSize']
            })
            user_group['groupNames'].append(group_info_map[group_id]['groupName'])
        user_groups.append(user_group)

    # 根据加群数量倒序排序
    user_groups = sorted(user_groups, key=lambda k: k['groupCount'])
    user_groups.reverse()

    # 导出所有重复加群用户的json
    with open(r'cache\user-groups.json', 'w', encoding='utf-8') as file_handler:
        json.dump(user_groups, file_handler, ensure_ascii=False, indent=4)

    # 导出重复加群用户的csv表格
    with open(r'cache\user-groups.csv', 'w', encoding='utf-8') as file_handler:
        # file_handler.writelines('用户ID,用户昵称,手机号,加入群')
        for user_group in user_groups:
            line = f"{user_group['userId']},{user_group['nickName']}," \
                   f"{user_group['loginId']},{','.join(user_group['groupNames'])}\n"
            file_handler.write(line)

    # 导出所有已入群的用户id列表json
    with open(r'cache\group_user_ids.json', 'w', encoding='utf-8') as file_handler:
        json.dump(group_user_ids, file_handler, ensure_ascii=False, indent=4)


def request_all_group_orgs():
    cache_file_path = r'cache\group-org-list.json'
    if os.path.isfile(cache_file_path):
        logger.info('load all group orgs from cache file: {}'.format(cache_file_path))
        with open(cache_file_path, 'r', encoding='utf-8') as fp:
            return json.load(fp)

    logger.info('begin request all group orgs.')
    params = {"tenantId": 6, "parentOrgId": "2020122914103700065623", "pageNo": 1, "pageSize": 100}
    data = request_group(params, GROUP_ORG_LIST_API)
    if not data or 'list' not in data or not data['list']:
        logger.error('request org list failed.')
        return False
    group_orgs = data['list']
    logger.info('end request all group orgs. size: {}'.format(len(group_orgs)))
    with open(cache_file_path, 'w', encoding='utf-8') as file_handler:
        json.dump(group_orgs, file_handler, ensure_ascii=False, indent=4)
    return group_orgs


# 获取所有群分组列表
def request_all_group_orgs_with_groups():
    cache_file_path = r'cache\group-org-list-with-groups.json'
    if os.path.isfile(cache_file_path):
        logger.info('load all group orgs with groups from cache: {}'.format(cache_file_path))
        with open(cache_file_path, 'r', encoding='utf-8') as fp:
            return json.load(fp)

    # 根据每个分组获取群列表
    max_page = 50
    group_orgs = request_all_group_orgs()
    for group_org in group_orgs:
        # 循环获取群列表
        org_groups = []
        org_id = group_org['orgId']
        logger.info('begin request org groups. orgId: {}'.format(group_org['orgId']))
        for page_index in range(1, max_page):
            params = {"orgId": org_id, "tenantId": "6", "pageNo": page_index}
            data = request_group(params, ORG_GROUP_LIST_API)
            if not data or 'list' not in data or not data['list']:
                logger.warning('request org list empty.')
                break
            logger.info('request org groups success. orgId: {}, page_index: {}, size: {}'
                        .format(org_id, page_index, len(data['list'])))
            org_groups.extend(data['list'])
        group_org['groups'] = org_groups
        logger.info('request all org groups success. org_id: {}, total size: {}'.format(org_id, len(org_groups)))

        with open(cache_file_path, 'w', encoding='utf-8') as file_handler:
            json.dump(group_orgs, file_handler, ensure_ascii=False, indent=4)
    return group_orgs


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

    # 已入群的用户id列表
    group_user_ids = []
    with open(r'cache\group_user_ids.json', 'r', encoding='utf-8') as file_handler:
        group_user_ids = json.load(file_handler)
    group_info_map = request_all_group_info()
    for group_show_id in uid_record_map.keys():
        uid_records = uid_record_map.get(group_show_id)
        if not uid_records:
            continue

        if group_show_id not in group_info_map:
            logger.error('未找到该群id： {}'.format(group_show_id))
            continue

        add_user_ids = []
        # 过滤已入群用户
        for uid_record in uid_records:
            uid = uid_record.get('uid')
            if uid in group_user_ids:
                logger.info('该用户已经加过群. uid: {}, group_show_id: {}'.format(uid, group_show_id))
                continue
            add_user_ids.append(uid)

        if not add_user_ids:
            logger.info('没有需要拉群的用户. group_show_id: {}'.format(group_show_id))
            continue

        group_info = group_info_map[group_show_id]
        logger.info('------> 处理拉群，群名称： {}, id: {}, 拉入用户id列表： {}'
                    .format(uid_records[0]['group_name'], group_info['groupId'], add_user_ids))
        import_uid_to_group(group_info['id'], group_info['groupId'], add_user_ids)
        # 更新结果并缓存，已经拉取过的不做处理


# 打包： pyinstaller -F alipay_group_import.py -p ../ --exclude-module pywin32 --exclude-module gevent
# --exclude-module flask --exclude-module APScheduler --exclude-module uiautomation
# 打包： pyinstaller alipay_group_import.spec
if __name__ == '__main__':
    # request_group_members_by_page('0194040001720220801200640503', 1, 10)
    # request_all_group_orgs()
    # request_all_group_orgs_with_groups()
    # request_all_group_all_members()
    # extract_repeat_group_user()
    run_import_uid()
    input("按任意键结束...")
