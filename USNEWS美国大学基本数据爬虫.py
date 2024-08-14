#encoding=utf-8
import requests
import time
import random
import logging
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import certifi

# 设置日志记录
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

user_agent_list = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0"
]

# 设置重试策略
retry_strategy = Retry(
    total=10,  # 重试次数
    backoff_factor=3,  # 重试间隔时间，指数退避策略
    status_forcelist=[429, 500, 502, 503, 504],  # 需要重试的状态码
    allowed_methods=['HEAD', "GET", "OPTIONS"]  # 需要重试的方法
)

adapter = HTTPAdapter(max_retries=retry_strategy)
http = requests.Session()
http.mount("https://", adapter)
http.mount("http://", adapter)

def random_delay(min_seconds=1, max_seconds=60):
    delay = random.uniform(min_seconds, max_seconds)
    logging.info(f"Sleeping for {delay:.2f} seconds")
    time.sleep(delay)

def test_proxy_connection(proxy, test_url='https://www.usnews.com'):
    try:
        response = requests.get(
            test_url, 
            params={'param': '1'}, 
            headers={'Connection':'close'}, 
            proxies=proxy, 
            timeout=10
        )
        response.raise_for_status()
        logging.info(f"Proxy {proxy} is working.")
        return True
    except requests.exceptions.RequestException as e:
        logging.error(f"Proxy {proxy} failed: {e}")
        return False

def fetch_data(start_page, end_page):
    for page in range(start_page, end_page):
        url = f'https://www.usnews.com/best-colleges/api/search?format=json&schoolType=national-universities&_sort=rank&_sortDirection=asc&page={page}'
        headers = {'User-Agent': random.choice(user_agent_list)}

        for attempt in range(5):  # 最多尝试5次
            try:
                #proxy = random.choice(proxies_list)  # 随机选择一个代理服务器（本地代理）
                #logging.info(f"Using proxy: {proxy}")
                logging.info(f"Using User-Agent: {headers['User-Agent']}")

                #if not test_proxy_connection(proxy):  # 测试代理连接
                 #   logging.error(f"Proxy {proxy} is not working, skipping...")
                 #   continue
                response = requests.get(
                    url='https://www.usnews.com/best-colleges/api/search?format=json&schoolType=national-universities&_sort=rank&_sortDirection=asc&page='+str(page),
                    #headers=headers,
                    params={'param': '1'}, 
                    headers=headers,
                    #proxies=proxy,
                    timeout=180,  # 增加超时时间到180秒
                    #verify=certifi.where()
                )
                response.raise_for_status()  # 检查请求是否成功
                logging.info(f"Page {page} response status: {response.status_code}")

                if response.status_code == 200:
                    try:
                        re_text = response.json()
                        li = re_text.get("items")
                        with open('collegeInfo_USA.txt', 'a', encoding='utf-8') as fp:
                            for i in li:
                                strConcat = f"{i['city']};{i['country_name']};{i['id']};{i['name']};{i['ranks'][0]['value']};{i['url']}"
                                print(strConcat)
                                fp.writelines(strConcat + '\n')
                    except ValueError as e:
                        logging.error(f"JSON解析失败: {e}")
                        logging.error(f"响应内容: {response.text[:500]}")
                    random_delay()  # 每次请求之间增加更长且随机的延迟
                break  # 如果请求成功则跳出重试循环
            except requests.exceptions.RequestException as e:
                logging.error(f"请求失败: {e}")
                random_delay(1, 60)  # 在重试之前增加更长且随机的间隔时间
            except requests.exceptions.ReadTimeout as e:  # 捕获读取超时异常
                logging.error(f"读取超时: {e}")
                random_delay(1, 60)  # 增加更长的延迟时间
            except ConnectionResetError as e:  # 捕获连接重置异常
                logging.error(f"连接重置错误: {e}")
                random_delay(1, 60)  # 增加更长的延迟时间
            except Exception as e:  # 捕获所有其他异常
                logging.error(f"其他错误: {e}")
                random_delay(1, 60)  # 增加更长的延迟时间

# 分段抓取过程，每次抓取一页并保存一次
for i in range(1, 44):
    fetch_data(i, i+1)
