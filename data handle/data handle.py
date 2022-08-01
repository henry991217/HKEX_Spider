import code
import csv
import os
import re
import urllib
from time import sleep

import xlwt
from lxml import etree
import requests
from bs4 import BeautifulSoup
from multiprocessing.dummy import Pool
import urllib3
from urllib3.util import url

urllib3.disable_warnings()

# 披露易数据搜集半自动爬虫#


filename = input("请输入文档创建日期：") + ".csv"
workdir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(workdir, filename)

# 创建workbook和sheet对象
workboot = xlwt.Workbook(encoding='utf-8')
worksheet = workboot.add_sheet('test')  # 设置工作表的名字
worksheet.col(0).width = 256 * 20  # 设置第一列列宽, 256为衡量单位，10表示10个字符宽度
worksheet.col(1).width = 256 * 10
worksheet.col(2).width = 256 * 10


def main():
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36'}
    url = "https://www3.hkexnews.hk/sdw/search/searchsdw_c.aspx"
    testusl="https://www3.hkexnews.hk/sdw/search/js/search_form.js?v=1640057418154"
    data = {
        '__EVENTARGUMENT': '',
        '__EVENTTARGET': 'btnSearch',
        '__VIEWSTATE': '/wEPDwUKMTY0ODYwNTA0OWRkM79k2SfZ+Vk Dy88JRhbk+XZIdM0=',
        '__VIEWSTATEGENERATOR': '3B50BBBD',
        'today': '',  # 今日日期
        'sortBy': 'shareholding',
        'sortDirection': 'desc',
        'alertMsg': '',
        'txtShareholdingDate': '',  # 查询的日期
        'txtStockCode': '01865',  # 查询的股票代码
        'txtStockName': '卓航控股',  # 股票名称
        'txtParticipantID': '',
        'txtParticipantName': ''
    }
    try:
        # another method
        # response = requests.get(url, headers=headers, data=data, verify=False)#缺少接口，无法获取相应数据
        # response = session.get(url, headers=headers)

        session = requests.Session()  # 通过会话请求
        cookies = requests.utils.dict_from_cookiejar(session.cookies)
        response = session.get(testusl, headers=headers,cookies=cookies,data=data)
        text = response.text
        print(text)  # 验证爬取的信息
        html = etree.XML(text)  # 解析字符串为html对象，自动补全html，body
        div = html.xpath("//div[@id='pnlResultNormal']|//table[@class='table table-scroll table-sort table-mobile-list']/tbody//tr/text()")  # xpath语法搜寻属性id为txtsharehodingdate的html元素,[]搜寻指定值
        sleep(5)
        data1 = []
        # print(data1)
        print("----" * 2)
        for tb in div:
            participant_id = tb.xpath("./td[1]/div/text()")[1]
            print("-----" + participant_id)
            participant_name = tb.xpath("./td[2]/div/text()")[1]
            participant_address = tb.xpath("./td[3]/div/text()")[1]
            right = tb.xpath("./td[4]/div/text()")[1]
            percent = tb.xpath("./td[5]/div/text()")[1]
            datadic = {"参与者编号": participant_id, "中央系统参与者名称": participant_name,
                       "地址": participant_address, "持股量": right, "占比": percent}
            data.append(datadic)
            csvhead = ["参与者编号", "中央系统参与者名称", "地址", "持股量", "占比"]
            os.chdir("D:\\披露易爬取文件\\data.xlsx")
            os.getcwd()
            with open(filename, 'wb', newline='') as fp:
                write = csv.DictWriter(fp, csvhead)
                write.writeheader()
                write.writerows(data1)

    except Exception as e:
        with open("log.txt", "a") as f:
            f.write(str() + "\n")
            print(e)
    print("爬取成功！")


# valuetime = html.xpath("//input[@id='txtShareholdingDate']/@value")
# print(valuetime[1] + "有数据")


if __name__ == "__main__":
    main()
