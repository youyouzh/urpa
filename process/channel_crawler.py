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
    save_filepath = r'cache\channel-top.json'
    if os.path.isfile(save_filepath):
        with open(save_filepath, 'r', encoding='utf-8') as fp:
            return json.load(fp)
    with open(save_filepath, 'w', encoding='utf-8') as fp:
        channels = request_data(top_chanel_query_url)
        json.dump(channels, fp, ensure_ascii=False, indent=4)
        return channels


# 爬取2级渠道
def crawler_level_2_tree(channels):
    save_filepath = r'cache\channel-level-2.json'
    if os.path.isfile(save_filepath):
        with open(save_filepath, 'r', encoding='utf-8') as fp:
            logger.info('load content from cache: {}'.format(save_filepath))
            return json.load(fp)

    for channel in channels:
        # 爬取2级列表，第一次请求查询1页
        channel_id = channel['id']
        # 页数设置1000比较大的，一次把所有数据加载出来
        logger.info('begin crawler channel 2 level data: {}'.format(channel_id))
        data = request_data(build_category_query_channel(0, 1000, '', channel_id))
        if 'content' not in data or not data['content']:
            logger.info('content is not exist for channel: {}'.format(channel_id))
            continue

        # 总页数
        total_size = data['totalElements']
        logger.info('query channel: {} total size: {}'.format(channel_id, total_size))
        channel['children'] = data['content']
        channel['size'] = total_size

    with open(save_filepath, 'w', encoding='utf-8') as save_handler:
        json.dump(channels, save_handler, ensure_ascii=False, indent=4)
    save_handler.close()
    return channels


# 爬取3级别以后的渠道
def crawler_level_3_tree(save_filepath: str = None):
    if save_filepath and os.path.isfile(save_filepath):
        with open(save_filepath, 'r', encoding='utf-8') as fp:
            logger.info('load content from cache: {}'.format(save_filepath))
            channels = json.load(fp)
    else:
        channels = crawler_top_channel()
        channels = crawler_level_2_tree(channels)

    for channel in channels:
        sub_channels = channel['children']
        for sub_channel in sub_channels:
            sub_channel_cache_path = r'cache\channel-cache-{}.json'.format(sub_channel['id'])
            # 检查是否已经爬取过，从文件中加载并做校验
            if os.path.isfile(sub_channel_cache_path):
                with open(sub_channel_cache_path, 'r', encoding='utf-8') as fp:
                    cache_sub_channel = json.load(fp)
                    total_size = cache_sub_channel.get('size', 0)
                    cache_children = cache_sub_channel.get('children', [])
                    sub_channel['children'] = cache_sub_channel['children']   # 使用文件中的节点替换channels全局中的节点
                    if total_size == len(cache_children) and sub_channel['id'] == cache_sub_channel['id']:
                        # 需要保证爬取的数量和实际数量相同
                        logger.info('load 3 level channel from cache: {}'.format(sub_channel_cache_path))
                        continue
                    logger.warning('The size is not equal for channel: {}. total size: {}, real size: {}'
                                   .format(sub_channel['id'], total_size, len(cache_children)))
            else:
                logger.warning('The channel is not exist: {}'.format(sub_channel['id']))

            logger.info('crawler 3 level channel: {}'.format(sub_channel['id']))
            # 递归爬取子节点
            crawler_sub_channel(channels, sub_channel)
            with open(sub_channel_cache_path, 'w', encoding='utf-8') as save_handler:
                json.dump(sub_channel, save_handler, ensure_ascii=False, indent=4)

        # 后期文件比较大，保存文件会非常慢，可以注释下面代码，爬取完之后，从文件加载合并
        # with open(save_filepath, 'w', encoding='utf-8') as save_handler:
        #     json.dump(channels, save_handler, ensure_ascii=False, indent=4)


# 去重处理
def merge_repeat_channel(channel):
    cache_children_channels = channel.get('children')
    total_size = channel.get('size', 0)
    if 0 < total_size <= len(cache_children_channels):
        # 如果实际数量大于总数量，则可能有重复数据进行去重
        no_repeat_children_channels = []
        no_repeat_ids = []  # 根据channel_id去重
        for channel in cache_children_channels:
            if channel['id'] not in no_repeat_ids:
                no_repeat_ids.append(channel['id'])
                no_repeat_children_channels.append(channel)
        logger.info('process repeat channel: {}, total size: {}'.format(channel['id'], total_size))
        channel['children'] = no_repeat_children_channels
    return channel


# 递归爬取子渠道
def crawler_sub_channel(channels, root_channel):
    channel_id = root_channel['id']

    # 已经有children节点说明已经爬取，则不用再请求
    if 'children' in root_channel:
        logger.info('The children is exist. do not need crawler. channel_id: {}'.format(channel_id))
        cache_children_channels = root_channel.get('children')
        total_size = root_channel.get('size', 0)

        if not cache_children_channels:
            # 没有子节点则直接返回
            return []
        # merge_repeat_channel(root_channel)   # 处理之前因为BUG产生的数据

        # 如果总节点数和实际节点数相等，则直接爬取子节点
        if total_size == len(root_channel['children']):
            # 递归循环爬取子节点
            for child_channel in root_channel['children']:
                crawler_sub_channel(channels, child_channel)
            return root_channel['children']

    logger.info('begin query channel: {}'.format(channel_id))
    logger.info('begin crawler channel sub level data: {}'.format(channel_id))
    # 将page_size设置比较大，尝试一次性获取所有数据，这个接口有BUG，可能只获取到20条数据，尤其是在totalSize比较大的情况下
    data = request_data(build_category_query_channel(0, 100, channel_id, ''))
    if 'content' not in data or not data['content']:
        logger.info('content is empty for channel: {}'.format(channel_id))
        root_channel['children'] = []
        root_channel['size'] = 0
        return []

    children_channels = data['content']
    logger.info('query channel: {} child total size: {}, real size: {}'
                .format(channel_id, data['totalElements'], len(children_channels)))

    # 检查返回的数据量是否和总数据量相同，有时候一页可能没有返回所有数据，需要分页查询
    if data['numberOfElements'] != data['totalElements']:
        logger.warning('The size is not equal, retry crawler: {}'.format(channel_id))
        real_page_size = len(children_channels)
        total_page_count = (data['totalElements'] // real_page_size) + 1
        children_channels = []
        # 分页查询，直到返回结果为空
        for page_index in range(total_page_count):
            logger.info('page crawl channel. id: {}, page index: {}'.format(channel_id, page_index))
            data = request_data(build_category_query_channel(page_index, real_page_size, channel_id, ''))
            if 'content' not in data or not data['content']:
                break
            logger.info('page crawl channel. id: {}, page index: {}, real size: {}'
                        .format(channel_id, page_index, len(data['content'])))
            children_channels.extend(data['content'])

    root_channel['children'] = children_channels
    root_channel['size'] = data['totalElements']

    # 递归循环爬取子节点
    for child_channel in root_channel['children']:
        crawler_sub_channel(channels, child_channel)
    return data['content']


if __name__ == '__main__':
    cache_dir = 'cache'
    if not os.path.isdir(cache_dir):
        logger.info('create cache dir: {}'.format(cache_dir))
        os.makedirs(cache_dir)
    crawler_level_3_tree(r'cache\channel-cache.json')
