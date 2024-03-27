import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from urllib.parse import urlparse, parse_qs

# 读取xlsx文件
df = pd.read_excel('your_file.xlsx')

# 设置Chrome浏览器选项
chrome_options = Options()
chrome_options.add_argument('--headless')  # 无界面模式，可选
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')

# 初始化Chrome浏览器驱动
driver = webdriver.Chrome(options=chrome_options)

# 定义函数以提取链接跳转后的URL中的utm_campaign
def extract_utm_campaign(url):
    try:
        driver.get(url)
        final_url = driver.current_url
        parsed_url = urlparse(final_url)
        query_params = parse_qs(parsed_url.query)
        utm_campaign = query_params.get('utm_campaign', [''])[0]
        return utm_campaign
    except:
        return ''

# 遍历Video Description列，提取链接并放置在对应行的URL列（G列）
df['URL'] = df['Video Description'].str.findall(r'(bit\.ly/\S+)').apply(lambda x: ', '.join([f'https://{link}' if not link.startswith('https://') else link for link in x]))

# 提取短链跳转后的URL中的utm_campaign，并放置在对应行的utm_campaign列（H列）
df['Campaigns'] = df['URL'].str.split(', ').apply(lambda x: ', '.join([extract_utm_campaign(url) for url in x]))

# 重新排序列的顺序，并添加URL和Campaigns列到原有数据的右侧
output_df = pd.DataFrame(df, columns=list(df.columns) + ['URL', 'Campaigns'])

# 保存结果到新的xlsx文件
output_df.to_excel('updated_tenorshare_KR_data.xlsx', index=False)

print("数据处理完成。")

# 关闭浏览器驱动
driver.quit()
