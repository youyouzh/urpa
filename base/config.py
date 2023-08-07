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

    # 企微应用id
    "wecom_corp_id": "wx7a24dd5cd3ba80b9",

    # 企微应用key
    "wecom_corp_key": "bGHzcBY5GyFmbUJUeoPCMF9IbHYb4L-k-ZnutrHllsY",

    # 企微消息推送应用id
    "wecom_agent_id": "1000034",
}

config_json_path = 'config.json'
if os.path.isfile(config_json_path):
    with open(config_json_path, encoding='utf-8') as f:
        file_config = json.load(f)
        CONFIG.update(file_config)
