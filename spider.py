# %%
'''国家人口健康科学数据中心'''
import requests
import pandas as pd
from io import StringIO
from progressbar import ProgressBar, SimpleProgress, Percentage, Bar
import time
# %%
port_page_dict = {'gc': 1064, # 国产
        'jk': 301 # 进口
        }
df = []
for port, page in port_page_dict.items():
    pages = list(range(page))
    pbar = ProgressBar(
        maxval=len(pages),
        widgets=['已完成 ', SimpleProgress(), ' 页 ', Percentage(), ' ', Bar()],
    ).start()
    # 启动后设置刷新间隔，每完成一页刷新一次
    pbar.update_interval = 1
    pbar.next_update = 1
    for i in pbar(pages):
        i = f'_{i}' if i > 0 else ''
        url = f'https://pharm.ncmi.cn/zysjk/yp{port}yp/index{i}.html'
        txt = requests.get(url)

        txt.raise_for_status()
        txt.encoding='utf-8'
        df_i = pd.read_html(StringIO(txt.text), extract_links='body')[0]
        df_cols = df_i.columns.tolist()
        df_i['link'] = df_i['详情'].map(lambda x: x[1])
        for i in df_cols:
            df_i[i] = df_i[i].map(lambda x: x[0])
        df_i = df_i.rename(columns={'批准文号': '批准文号/注册证号',
                                    '注册证号': '批准文号/注册证号',
                                    '产品名称(中文)': '产品名称'},
                           )
        df_i['port'] = port
        
        df.append(df_i)
        if len(df) % 5 == 0:
            time.sleep(3)
# %%
df = pd.concat(df)
df
# %%
df['detail_link']=df.apply(lambda x: f'https://pharm.ncmi.cn/zysjk/yp{x["port"]}yp/{"/".join(x["link"].rsplit("/")[1:])}', axis = 1)
df
# %%
df = df.reset_index(drop = True)
df.to_excel('国家人口健康科学数据中心药品库.xlsx',
            index = False)
# %%
df = pd.read_excel('国家人口健康科学数据中心药品库.xlsx')
df
# %%
txt = requests.get(df['detail_link'][0])
txt
# %%
txt.content
# %%
txt.raise_for_status()
txt.encoding='utf-8'
# %%
pd.read_html(StringIO(txt.text))
# %%


# %%
import base64
# Base64 解码示例：替换下面的 Base64 字符串后，单独运行本单元格即可
base64_to_decode = 'aWQ9MDAwMDdlMTMxMDA1NGMxZmI4ODY4MTU2NTY2OGM4MmUmaXRlbUlkPWZmODA4MDgxODNjYWQ3NTAwMTg0MDg4MWY4NDgxNzlm'
base64_to_decode = 'aWQ9NjI5NTNmZTM5NzBmNTQyNzBmZWNiMmQzZDZhNzY3OGYmaXRlbUlkPWZmODA4MDgxODNjYWQ3NTAwMTg0MDg4NjY1NzExODAw'
base64_decoded = base64.b64decode(base64_to_decode, validate=True).decode('utf-8')
print('Base64 解码结果：', base64_decoded)
# %%
urlparam_decoded = dict([a.split('=') for a in base64_decoded.split('&')])
urlparam_decoded
# %%
# MD5 哈希示例：修改下面的字符串后，单独运行本单元格即可
import hashlib

approve_code = '国药准字H11020913'
id = hashlib.md5(approve_code.encode('utf-8')).hexdigest()
print('原始字符串：', approve_code)
print('MD5 哈希值：', id)

# Base64 编码示例：修改字符串后，单独运行本单元格即可
urlparam = f'id={id}&itemId=ff80808183cad75001840881f848179f'
urlparam_base64 = base64.b64encode(urlparam.encode('utf-8')).decode('ascii')
print('原始字符串：', urlparam)
print('Base64 编码：', urlparam_base64)
# %%
url = f"https://datasearch.nmpa.gov.cn/datasearch/search-info.html?nmpa={urlparam_base64}"
print(url)
# %%
txt
# %%
url = "https://datasearch.nmpa.gov.cn/datasearch/search-info.html?nmpa=aWQ9NDA5YTM5ZjQwMDUxMWE5NzVjYzc3ZmNmNzEzYzhmNDMmaXRlbUlkPWZmODA4MDgxODNjYWQ3NTAwMTg0MDg4MWY4NDgxNzlm"
user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36"
# %%
import requests

# 对比页面入口的两种请求方式；接口地址和请求体需以浏览器实际请求为准
headers = {"User-Agent": user_agent}
responses = {}
for method in ("GET", "POST"):
    try:
        response = requests.request(method, url, headers=headers, timeout=(10, 20))
    except requests.RequestException as exc:
        print(f"{method} 请求失败，异常类型：{type(exc).__name__}")
        continue
    responses[method] = response
    print(f"请求方式：{method}，最终请求方式：{response.request.method}")
    print(f"状态码：{response.status_code}，响应类型：{response.headers.get('Content-Type', '未提供')}")
    print(f"最终地址：{response.url}")
    print(f"响应字节数：{len(response.content)}")
    print("响应内容前 500 个字符：", response.text[:500])
    print("请结合响应内容判断是否返回药品数据，状态码成功不代表提取成功。")
# %%
responses
# %%
