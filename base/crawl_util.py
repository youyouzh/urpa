import json
import re

from base.log import logger


def load_header_from_curl_bash(curl_bash_content: str):
    """
    从浏览器网络中抓包导出请求导出header，通过右键请求【复制 -> Curl Bash】导出
    :param curl_bash_content: 复制的curl bash请求内容
    :return:
    """
    # 提取请求头
    header = {}
    match_header_regex = re.compile(r' +-H \'([\w-]+): ([^\'^]+)\'')
    header_match_items = match_header_regex.findall(curl_bash_content)
    if not header_match_items:
        logger.info('not format header.')
        return {}

    for header_match_item in header_match_items:
        header_key = header_match_item[0]
        header_value = header_match_item[1]
        header[header_key.strip()] = header_value.strip()
    logger.info('extract header success: {}'.format(header))
    print(json.dumps(header, ensure_ascii=True, indent=2))
    return header


if __name__ == '__main__':
    header_content = """
curl 'https://www.chinawealth.com.cn/LcSolrSearch.go' \
  -H 'Accept: application/json, text/javascript, */*; q=0.01' \
  -H 'Accept-Language: zh-CN,zh;q=0.9' \
  -H 'Connection: keep-alive' \
  -H 'Content-Type: application/x-www-form-urlencoded; charset=UTF-8' \
  -H $'Cookie: BIGipServerPool_SuperFusion_LiCai_fe_8080=\u0021YVTnq940oMesT72XKFNNtTIPej9bYIBszyz2rS4nA/rMdaSeoFEIZCHRaNxCl9sjhiVZmFFiHYyomrw=; BIGipServerPool_SuperFusion_LiCai_Nginx_8080=\u002134YLRj5xeC8PrvKLR7FEljAmky2oSOExukhkQkiNzs44/bAfQZBTmJTF6UoErnPIMKOsrbus1N2fJOI=; JSESSIONID=CC75163965798C4B8BE0E984C5CFE66B; _pk_ses.12.8bc7=*; _pk_id.12.8bc7=21417eaa1785f5d6.1709013510.5.1709083630.1709083462.' \
  -H 'Origin: https://www.chinawealth.com.cn' \
  -H 'Referer: https://www.chinawealth.com.cn/zzlc/jsp/lccp.jsp' \
  -H 'Sec-Fetch-Dest: empty' \
  -H 'Sec-Fetch-Mode: cors' \
  -H 'Sec-Fetch-Site: same-origin' \
  -H 'User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36' \
  -H 'X-Requested-With: XMLHttpRequest' \
  -H 'sec-ch-ua: "Chromium";v="122", "Not(A:Brand";v="24", "Google Chrome";v="122"' \
  -H 'sec-ch-ua-mobile: ?0' \
  -H 'sec-ch-ua-platform: "Windows"' \
  --data-raw 'cpjglb=01&cpyzms=&cptzxz=&cpfxdj=&cpqx=&mjbz=&cpzt=02&mjfsdm=01%2CNA&cptssx=&cpdjbm=&cpmc=&cpfxjg=&yjbjjzStart=&yjbjjzEnd=&areacode=&pagenum=1&orderby=&code='
    """
    load_header_from_curl_bash(header_content)
