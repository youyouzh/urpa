"""
全局配置项
"""
import json
import os

CONFIG = {
    # http服务端口
    "http_server_port": 8032,

    # 上传文件保存路径
    "upload_path": "upload",

    # 企微应用id，这个测试id是自建的企业，需要测试可以找维护人员添加
    "wecom_corp_id": "wwb36926885246afd4",

    # 企微应用key
    "wecom_corp_key": "aHUghPqj3Fd-m4sVuc7Qv9zqo8uhlEl4mgXed5Na-5U",

    # 企微消息推送应用id
    "wecom_agent_id": "3010041",
}


def load_config(default_config: dict, config_json_path: str = 'config.json') -> dict:
    if os.path.isfile(config_json_path):
        with open(config_json_path, encoding='utf-8') as f:
            file_config = json.load(f)
            default_config.update(file_config)
    return default_config


load_config(CONFIG)
