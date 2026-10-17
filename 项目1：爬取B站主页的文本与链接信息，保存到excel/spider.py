import requests
import xlwt
from bs4 import BeautifulSoup
import lxml
import urllib


# 设置请求头，防止被封禁
# 隐藏伪装headers(内部为bilibili主页的User-Agent)
headers = {
    'User-Agent': 
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
    ' AppleWebKit/537.36 (KHTML, like Gecko)'
    ' Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0'
}


# 引入主页链接地址
url = 'https://www.bilibili.com/'
# 定义主函数
def main():
    # 发送GET请求获取网页内容
    res = requests.get(url, headers=headers)
    # 使用BeautifulSoup解析网页内容
    soup = BeautifulSoup(res.text, 'lxml')
    # 返回解析后的soup对象
    return soup


# 定义获取数据的函数
def get_data(soup):
    # 获取所有的a标签
    a_tags = soup.find_all('a')
    # 创建一个空列表来存储数据
    data = []
    # 遍历所有的a标签，获取href和文本内容
    for a in a_tags:
        # 获取href属性和文本内容
        href = a.get('href')
        # 获取文本内容并去除首尾空格,strip=True表示去除首尾空格
        text = a.get_text(strip=True)
        # 如果href和文本内容都不为空，则将其添加到数据列表中
        if href and text:
            data.append((text, href))
    return data


# 定义写入Excel的函数
def write_to_excel(data):
    # 创建一个新的Excel工作簿   
    workbook = xlwt.Workbook()
    # 创建一个新的工作表,bilibili Links为工作表名称
    sheet = workbook.add_sheet('Bilibili Links')
    # 写入表头,0表示第一行，0表示第一列
    sheet.write(0, 0, 'Text')
    # 写入表头,0表示第一行，1表示第二列
    sheet.write(0, 1, 'URL')
    # 写入数据,index表示行号，text表示文本内容，href表示链接地址,enumerate(data, start=1)表示从1开始计数,0表示第一行，1表示第二列
    for index, (text, href) in enumerate(data, start=1):
        # 写入数据到Excel文件中,0表示第一列，1表示第二列
        sheet.write(index, 0, text)
        # 写入数据到Excel文件中,0表示第一列，1表示第二列
        sheet.write(index, 1, href)
    # 保存Excel文件
    workbook.save('bilibili_links.xls') 

    
# 调用主函数获取soup对象
soup = main()
# 调用获取数据的函数，传入soup对象
data = get_data(soup)
# 打印获取到的数据
print(data)
# 打印获取到的链接总数
print(f"Total links found: {len(data)}")
# 调用写入Excel的函数，传入获取到的数据
write_to_excel(data)


"""
以上是我的爬虫代码
主要爬取bilibili主页的所有链接
将其写入Excel文件中。
是我人生中的第一个爬虫项目
虽然很简单，但是我学到了很多东西
比如

如何使用requests库发送请求
如何使用BeautifulSoup解析网页内容
如何使用xlwt库写入Excel文件等等。

开启了我的爬虫之路，也为我打开了新的大门。
"""