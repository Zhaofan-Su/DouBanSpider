import requests
from bs4 import BeautifulSoup
from openpyxl import workbook
from openpyxl import load_workbook
import os
import re

types = ['编程', '互联网', 'web', '交互设计', '经济', '算法']


def get_data_by_type(t):
    # 全局工作表对象
    global ws
    url = "https://book.douban.com/tag" + "/" + t
    # 得到网页
    data = requests.get(url)
    # 网页数据转换
    text = BeautifulSoup(data.text, 'lxml')
    # 书籍名称

    names = []
    # 简介
    descs = []
    # 价格
    prices = []
    # 详情
    details = []
    # 图片地址，使用豆瓣公开图库
    images = []

    # 先找存图书图片的img标签
    imgs = text.select("#subject_list > ul > li > div.pic > a > img")
    # 找到标签里的src的值，存起来
    for img in imgs:
        images.append(img.get('src'))

    # 找书名
    titles = text.select("#subject_list > ul > li > div.info > h2 > a")
    for title in titles:
        names.append(title.get_text().split()[0])

    # 找简介,包含作者，（译者），出版社，出版时间;找价格
    pubs = text.select("#subject_list > ul > li > div.info > div.pub")
    for pub in pubs:
        pub_text = pub.get_text()
        descs.append(pub_text[:pub_text.rindex("/") - 1])
        price = (pub_text[pub_text.rindex("/") + 1:])
        if price.split()[0] == "USD":
            prices.append(float(price.split()[1]) * 6)
        elif price.split()[0] == "CNY":
            prices.append(float(price.split()[1]))
        elif price.split()[0] == "$":
            prices.append(float(price.split()[1]) * 6)
        elif (price.split()[0])[-1] == "元":
            prices.append(float((price.split()[0])[0:-2]))
        elif price.split()[0].isdigit():
            prices.append(float(price.split()[0]))
        else:
            prices.append(45.00)

    # 找详情
    decos = text.select("#subject_list > ul > li > div.info > p")
    for deco in decos:
        details.append(deco.get_text())

    for i in range(0, len(names)):
        ws.append([names[i], descs[i], prices[i], details[i], images[i]])


if __name__ == '__main__':
    for t in types:
        # 创建Excel表并写入数据
        wb = workbook.Workbook()
        # 获取当前正在操作的表对象
        ws = wb.active
        ws.append(['names', 'descs', 'prices', 'details', 'images'])
        get_data_by_type(t)
        wb.save(t + '.xlsx')
