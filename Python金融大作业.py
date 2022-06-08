# 姓名： 陈其桦
# 学号： L191000709
# 公司： 000014 沙河股份


# 第一题
import tushare as ts
import pandas as pd
from scipy.stats import pearsonr
import matplotlib.pyplot as plt
import xlwings as xw
pd.set_option('display.max_columns', None)
stock_code = '000014'
stock_name = '沙河股份'
start_date = '2022-05-05'  
end_date = '2022-06-02'
stock_data = ts.get_hist_data(stock_code, start=start_date, end=end_date) 

# 获取股价涨跌幅和成交量数据
stock_table = stock_data[['p_change', 'volume']]
stock_table.columns = ['股价涨跌幅','成交量']
print(stock_table)

# 采用昨日成交量作为初始量计算涨跌幅
stock_table['昨日成交量'] = stock_table['成交量'].shift(-1)
stock_table['成交量涨跌幅（昨日算法）(%)'] = (stock_table['成交量']-stock_table['昨日成交量'])/stock_table['昨日成交量']*100
print(stock_table)


# 采用多日成交量均值计算涨跌幅（采用10日）
stock_table['成交量10日均值'] = stock_table['成交量'].sort_index().rolling(10, min_periods=1).mean()
stock_table['成交量涨跌幅（前10日均值）(%)'] = (stock_table['成交量']-stock_table['成交量10日均值'])/stock_table['成交量10日均值']*100
print(stock_table)


# 昨日成交量涨跌辐(%)和股价涨幅的相关性分析
corr1 = pearsonr(abs(stock_table['股价涨跌幅'][:-1]), abs(stock_table['成交量涨跌幅（昨日算法）(%)'][:-1]))
print(f'通过昨日成交量计算的相关系数r值为{corr1[0]}，显著性水平P值为{corr1[1]}，因此为{"显著" if corr1[1] < 0.05 else "不显著"}')

# 多日成交量涨跌幅2(%)和股价涨幅的相关性分析
corr2 = pearsonr(abs(stock_table['股价涨跌幅']), abs(stock_table['成交量涨跌幅（前10日均值）(%)']))
print(f'通过多日成交量计算的相关系数r值为{corr2[0]}，显著性水平P值为{corr2[1]}，因此为{"显著" if corr2[1] < 0.05 else "不显著"}')

# 判断哪一个算法更优
if corr1[1] < corr2[1]:
    print(f"使用昨日成交量计算更佳，相关系数r值为{corr1[0]}")
else:
    print(f"使用多日成交量计算更佳，相关系数r值为{corr2[0]}")

# 提取数据列
target_columns = ['股价涨跌幅', '成交量涨跌幅（昨日算法）(%)', '成交量涨跌幅（前10日均值）(%)']
final_table = stock_table[target_columns]
final_table = final_table[::-1]

fig = plt.figure(figsize=(10, 5))
plt.rcParams['font.family'] = ['SimHei']  #黑体
plt.rcParams['axes.unicode_minus'] = False

# 绘制第一个折线图：股价涨跌幅(%)
plt.plot(final_table.index, final_table['股价涨跌幅'].apply(lambda x: abs(x)), label='股价涨跌幅(%)', color='red') 
plt.legend(loc='upper left')

# 绘制第二个折线图：成交量涨跌幅(%)
plt.twinx()
plt.plot(final_table.index, final_table['成交量涨跌幅（前10日均值）(%)'].apply(lambda x: abs(x)), label='成交量涨跌幅（前10日均值）(%)', linestyle='--')
plt.legend(loc='upper right')

# 设置图片标题，自动调整x坐标轴刻度的角度并展示图片
plt.title(stock_name)  
plt.gcf().autofmt_xdate(rotation=60)  
plt.show()

# 使用Excel操作库
app = xw.App(visible=False)

# 创建新Excel工作簿
wb = app.books.add()

# 创建新工作表
sht = wb.sheets.add(stock_name)
sht.range('A1').value = final_table
xw.Range('A1:D1').columns.autofit()
sht.pictures.add(fig, name='图片1', update=True, left=450)
wb.save(r'D:\第一题Excel文档.xlsx')

wb.close()  
app.quit()



# 第二题
import requests
import re
import itertools

def parse_sogou_news(start_page, end_page, company, save_path):
    # 定义浏览器设置
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36'}
    
    # 定义各个字段的规律
    title_patt = '<h3 class="vr-title.*?>.*?<a.*?>(.*?)</a>'
    link_patt = '<h3 class="vr-title.*?>.*?<a.*?href="(.*?)".*?>'
    date_patt = '<span class="text-lightgray.*?>(.*?)-.*?</span>'
    source_patt = '<div class="citeurl.*?>.*?<span>(.*?)(?= |\.{2}).*?</span>'
    
    # 网站前缀
    website_prefix = "https://www.sogou.com"
    
    # 计数器
    idx = 0
    with open(save_path, "w") as file:
       
        # 爬取多页信息
        for page in range(start_page, end_page):
            
            # 换的是query和page变量
            website = f"https://www.sogou.com/sogou?query={company}&pid=sogou-wsse-9fc36fa768a74fa9&duppid=1&cid=&interation=1728053249&s_from=result_up&sut=5663&sst0=1654605625239&lkt=0%2C0%2C0&sugsuv=00183F30939EC58A629F470B97338784&sugtime=1654605625239&page={page}&ie=utf8&w=01029901&dr=1"
            result = requests.get(website, headers=headers).text
            titles = re.findall(title_patt, result, re.S)
            links = re.findall(link_patt, result, re.S)
            sources = re.findall(source_patt, result, re.S)
            dates = re.findall(date_patt, result, re.S)
            
            # 标题需要过滤掉<em>,<!--> 等信息
            for i in range(len(titles)):
                titles[i] = re.sub("<.*?>", "", titles[i].strip())
                
            # 写入文档
            for title, link, source, date in itertools.zip_longest(titles, links, sources, dates, fillvalue = "None"):

                file.writelines(f"{idx+1}. {title}\n来源于：{source}\n发布于：{date}\n{website_prefix}{link}\r\n")
                idx += 1

# 爬取前三页数据
parse_sogou_news(0, 3, stock_name, r"D:\第二题Txt文档.txt")



# 第三题
from selenium import webdriver
import time

# 模拟Chrome浏览器
browser = webdriver.Chrome()
browser.maximize_window()
browser.get("https://so.eastmoney.com/TieZi/s?keyword=")

# 使用输入框输入查询公司并点击搜索按钮
browser.find_element_by_xpath('//*[@id="search_key"]').clear()
browser.find_element_by_xpath('//*[@id="search_key"]').send_keys(stock_name)
browser.find_element_by_xpath('//*[@id="app"]/div[1]/div[1]/div[1]/form/input[2]').click()
time.sleep(3)

# 获取网页源代码
data = browser.page_source
browser.quit()

# 定义字段的爬取规律
title_patt = '<div class="article_title">.*?<a.*?>(.*?)</a>'
link_patt = '<div class="article_title">.*?<a href="(.*?)".*?>'
source_patt = '<div class="article_title">.*?<span class="articel_ba">(.*?)</span>'
date_patt = '<div class="article_content">.*?<label>(.*?)</label>'

titles = re.findall(title_patt, data, re.S)
links = re.findall(link_patt, data, re.S)
sources = re.findall(source_patt, data, re.S)
dates = re.findall(date_patt, data, re.S)

# 标题需要过滤掉<em>,<!--> 等信息
for i in range(len(titles)):
    titles[i] = re.sub("<.*?>", "", titles[i])

# 写入文件
idx = 0
with open(r"D:\第三题Txt文档.txt", "w") as file:
    for title, link, source, date in zip(titles, links, sources, dates):
        file.writelines(f"{idx+1}. {title}\n来源于：{source}\n发布于：{date}\n{link}\r\n")
        idx += 1



# 第四题
import tushare as ts
import numpy as np
import pandas as pd
import scipy.stats as scs
import statsmodels.api as sm
import matplotlib.pyplot as plt

start_date = "2021-01-04"
end_date = "2021-12-31"
stock_codes = ["000014","399001"]

# 获取收盘价
a_stock_data = ts.get_hist_data(stock_codes[0], start=start_date, end=end_date)['close']
index_data = ts.get_hist_data(stock_codes[1], start=start_date, end=end_date)['close']

# 将数据整合清洗
data = pd.concat([a_stock_data, index_data], axis=1, join='inner')
data.columns = ["沙河股份收盘价","深证成指收盘价"]
data = data.dropna()[::-1]
print(data)

# 对数据进行规范化，采用的是正态标准化，计算z-score，公式 = (每个样本 - 均值) / 标准差
standard_data = scs.zscore(data)
print(standard_data)

# 绘制规范化数据线图
fig = plt.figure(figsize=(10, 5))
plt.plot(standard_data["沙河股份收盘价"], 'go--', linewidth=2, label="沙河股份标准化收盘价")
plt.plot(standard_data["深证成指收盘价"], 'b-', linewidth=2, label="深证成指标准化收盘价")
plt.legend()
plt.xticks(standard_data.index[::10])
plt.title("标准化收盘价")
fig.autofmt_xdate(rotation=80)
plt.show()

# 绘制股票与指数的散点图
fig = plt.figure(figsize=(10, 5))
plt.scatter(standard_data["沙河股份收盘价"], standard_data["深证成指收盘价"], c="blue")
plt.title("沙河股份收盘价对深证成指收盘价散点图")
plt.show()

# 计算对数收益率
returns = np.log(data / data.shift(1))
returns.columns = ["沙河股份对数收益率","深证成指对数收益率"]
print(returns)

# 使用QQ图对对数收益率进行正态性检验
sm.qqplot(returns["沙河股份对数收益率"], line='s')
plt.title("沙河股份对数收益率QQ图")
plt.xlabel('theoretical quantiles')
plt.ylabel('sample quantiles')

sm.qqplot(returns["深证成指对数收益率"], line='s')
plt.title("深证成指对数收益率QQ图")
plt.xlabel('theoretical quantiles')
plt.ylabel('sample quantiles')

# 使用正态性双侧检验，零假设：数据服从正态分布
def normality_tests(arr):
    print(f'{"-"*40}')
    print('Skew of data set  %14.3f' % scs.skew(arr))
    print('Skew test p-value %14.3f' % scs.skewtest(arr)[1])
    print('Kurt of data set  %14.3f' % scs.kurtosis(arr))
    print('Kurt test p-value %14.3f' % scs.kurtosistest(arr)[1])
    print('Norm test p-value %14.3f' % scs.normaltest(arr)[1])
    if scs.normaltest(arr)[1] < 0.05:
        print("The data is not normally distributed")
    else:
        print("The data is normally distributed")
    print(f'{"-"*40}')

for col in returns:
    normality_tests(returns[col][1:])