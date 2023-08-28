"""
银行渠道抓取
"""
import json
import os.path
import time

import requests
from base.log import logger


ALL_CHANNEL_API = 'https://digimkt.bosera.com/boseromee/aim-share/uc-system/mobile' \
                  '/api/v1/org/orgCategory/findAllDisplay?timeStemp={}'
CATEGORY_CHANNEL_API = 'https://digimkt.bosera.com/boseromee/aim-share/uc-system/mobile' \
                       '/api/v1/org/queryOrgNode' \
                       '?timeStemp={timestamp}&page={page}&pageSize={page_size}&id={id}' \
                       '&name=&orgCategoryId={category_id}'


def build_category_query_channel(page: int, page_size: int, id, category_id):
    data = {
        'page': str(page),
        'page_size': str(page_size),
        'id': str(id),
        'category_id': str(category_id),
        'timestamp': str(int(time.time() * 1000))
    }
    return CATEGORY_CHANNEL_API.format(**data)


def request_data(url: str) -> dict:
    response = requests.get(url)
    if response.status_code != 200:
        logger.info('request status not 200: {}'.format(response.status_code))
        return {}
    data = json.loads(response.text)
    if 'code' not in data or 'data' not in data or data['code'] != 0:
        logger.info('response json data error. data: {}, url: {}'.format(data, url))
        return {}
    return data['data']


def crawler_top_channel():
    top_chanel_query_url = ALL_CHANNEL_API.format(str(int(time.time() * 1000)))
    save_filepath = 'channel-top.json'
    if os.path.isfile(save_filepath):
        with open(save_filepath, 'r', encoding='utf-8') as fp:
            return json.load(fp)
    with open(save_filepath, 'w', encoding='utf-8') as fp:
        channels = request_data(top_chanel_query_url)
        json.dump(channels, fp, ensure_ascii=False, indent=4)
        return channels


# 爬取2级渠道
def crawler_level_2_tree(channels):
    save_filepath = 'channel-level-2.json'
    if os.path.isfile(save_filepath):
        with open(save_filepath, 'r', encoding='utf-8') as fp:
            logger.info('load content from cache: {}'.format(save_filepath))
            return json.load(fp)

    for channel in channels:
        # 爬取2级列表，第一次请求查询1页
        channel_id = channel['id']
        data = request_data(build_category_query_channel(0, 10, '', channel_id))
        logger.info('begin crawler channel 2 level data: {}'.format(channel_id))
        if 'content' not in data or not data['content']:
            logger.info('content is not exist for channel: {}'.format(channel_id))
            continue

        # 总页数
        total_size = data['totalElements']
        logger.info('query channel: {} total size: {}'.format(channel_id, total_size))
        data = request_data(build_category_query_channel(0, total_size, '', channel_id))
        if 'content' not in data or not data['content']:
            logger.info('content is not exist for channel: {}'.format(channel_id))
            continue
        channel['children'] = data['content']
        channel['size'] = total_size

    with open(save_filepath, 'w', encoding='utf-8') as save_handler:
        json.dump(channels, save_handler, ensure_ascii=False, indent=4)
    save_handler.close()
    return channels


# 爬取3级别以后的渠道
def crawler_level_3_tree(channels):
    save_filepath = 'channel-level-3.json'
    if os.path.isfile(save_filepath):
        with open(save_filepath, 'r', encoding='utf-8') as fp:
            logger.info('load content from cache: {}'.format(save_filepath))
            return json.load(fp)

    for channel in channels:
        sub_channels = channel['children']
        empty_count = 0
        for sub_channel in sub_channels:
            logger.info('crawler 2 level channel: {}'.format(sub_channel['id']))
            result = crawler_sub_channel(sub_channel)
            if not result:
                empty_count += 1
            if empty_count >= 4:
                logger.info('break child channel: {} crawler with empty count: {}'
                            .format(sub_channel['id'], empty_count))
                break


# 递归爬取子渠道
def crawler_sub_channel(root_channel):
    # 试探性的请求第一页得到总数据量
    channel_id = root_channel['id']
    logger.info('begin query channel: {}'.format(channel_id))
    data = request_data(build_category_query_channel(0, 10, channel_id, ''))
    logger.info('begin crawler channel 2 level data: {}'.format(channel_id))
    if 'content' not in data or not data['content']:
        logger.info('content is empty for channel: {}'.format(channel_id))
        return []

    # 一次性查询所有数据
    total_size = data['totalElements']
    logger.info('query channel: {} child total size: {}'.format(channel_id, total_size))
    data = request_data(build_category_query_channel(0, total_size, channel_id, ''))
    if 'content' not in data or not data['content']:
        logger.info('content is not exist for channel: {}'.format(channel_id))
        return []
    root_channel['children'] = data['content']
    root_channel['size'] = total_size
    with open('channel-cache.json', 'w', encoding='utf-8') as save_handler:
        json.dump(channels, save_handler, ensure_ascii=False, indent=4)

    # 递归循环爬取子节点
    empty_count = 0
    for child_channel in root_channel['children']:
        result = crawler_sub_channel(child_channel)

        # 统计查询结果，如果很多空的，则不继续遍历子节点
        if not result:
            empty_count += 1
        if empty_count >= 4:
            logger.info('break child channel: {} crawler with empty count: {}'
                        .format(child_channel['id'], empty_count))
            break
    return data['content']


if __name__ == '__main__':
    channels = crawler_top_channel()
    channels = crawler_level_2_tree(channels)
    crawler_level_3_tree(channels)
