"""
整个服务的集成自动化测试，覆盖主流分支节点
"""
import datetime
import json

import requests


CHAT_MESSAGE_SEND_API = 'http://10.191.0.114/api/bot/chat/send'
WECHAT_FROM_SUBJECT = 'Ray'
NOW_TIME_STR = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
MESSAGE_CASES = [
    {
        "channel": "WECHAT",
        "fromSubject": WECHAT_FROM_SUBJECT,
        "toConversations": ["文件传输助手"],
        "messageType": "TEXT",
        "messageData": {
            "content": "自动化测试-文本消息发送: " + NOW_TIME_STR
        }
    },
    {
        "channel": "WECHAT",
        "fromSubject": WECHAT_FROM_SUBJECT,
        "toConversations": ["文件传输助手"],
        "messageType": "FILE",
        "messageData": {
            "filePaths": [r"test-upload-file.txt"],
            "content": "自动化测试-文件附带消息发送: " + NOW_TIME_STR
        }
    },
    {
        "channel": "WECHAT",
        "fromSubject": WECHAT_FROM_SUBJECT,
        "toConversations": ["文件传输助手", "rpa消息测试群", "rpa消息测试群-不存在"],
        "messageType": "FILE",
        "messageData": {
            "filePaths": [r"test-upload-file.txt"],
            "content": "自动化测试-文件附带消息发送: " + NOW_TIME_STR
        }
    },
    {
        "channel": "WECOM_GROUP_BOT",
        "fromSubject": "5e86bea4-a978-4c6a-97dd-35544e52485c",
        "messageType": "TEXT",
        "messageData": {
            "content": "自动化测试-企微群机器人消息发送：" + NOW_TIME_STR,
            "mentionedList": ["zhaohai"]
        },
    },
    {
        "channel": "WECOM_APP_MESSAGE",
        "fromSubject": "RPA助手",
        "messageType": "TEXT",
        "messageData": {
            "content": "自动化测试-企微应用消息推送：" + NOW_TIME_STR,
        },
    }
]


def request_agent(message):
    response = requests.post(CHAT_MESSAGE_SEND_API, json=message)
    assert response.status_code == 200
    return json.loads(response.text)


def message_send_test():
    for message in MESSAGE_CASES:
        result = request_agent(message)
        assert result['code'] == 0


# 测试正常单个微信文本消息发送
def message_send_wechat_text_test():
    result = request_agent(MESSAGE_CASES[0])
    assert result['code'] == 0


# 测试单个微信文件发送
def message_send_wechat_file_test():
    result = request_agent(MESSAGE_CASES[1])
    assert result['code'] == 0


# 测试微信多个群的微信文件和文本发送
def message_send_wechat_file_test_of_multi():
    result = request_agent(MESSAGE_CASES[2])
    assert result['code'] == 0


# 测试企微群机器人推送
def message_send_wecom_group_bot_test():
    result = request_agent(MESSAGE_CASES[3])
    assert result['code'] == 0


# 测试企微应用消息推送
def message_send_wecom_app_push_message_test():
    result = request_agent(MESSAGE_CASES[4])
    assert result['code'] == 0


def run_all_test():
    message_send_wechat_text_test()
    message_send_wechat_file_test()
    message_send_wechat_file_test_of_multi()
    message_send_wecom_group_bot_test()
    message_send_wecom_app_push_message_test()


if __name__ == '__main__':
    run_all_test()
