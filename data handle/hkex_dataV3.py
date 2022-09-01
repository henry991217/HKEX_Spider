import datetime
import os
import tkinter
from lxml import etree
import requests
import requests_html
import xlwt
from tkinter import simpledialog
from tkinter import *

requests.packages.urllib3.disable_warnings()

participant_name=[]
shareholding=[]
sharehoding_percent=[]#差额
# hkex披露易1529定向爬虫
class HKex_Search:
    URL = "https://www3.hkexnews.hk/sdw/search/searchsdw_c.aspx"
    path='D:/披露易每日定向数据/'
    def __init__(self):

        self.session = requests_html.HTMLSession()  # 创建的request对象


    def get_stockdata(self, searchDate, stockcode):

        ''':argument:对象参数包括"20XX/XX/XX“和股票代号'''


        header = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36',
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9'
        }
        postdata = {
            '__EVENTTARGET': 'btnSearch',
            '__EVENTARGUMENT': '',
            '__VIEWSTATE': "/wEPDwUKMTY0ODYwNTA0OWRkM79k2SfZ+VkDy88JRhbk+XZIdM0=",  # 表单数据提交
            "__VIEWSTATEGENERATOR": "3B50BBBD",
            "sortBy": "shareholding",
            "sortDirection": "desc",
            "txtShareholdingDate": searchDate,
            "txtStockCode": stockcode,
        }


        response = requests.post(url=self.URL, data=postdata, headers=header)
        html0 = response.content.decode("utf-8")  # 确认页面爬取内容
        # print(html0)
        parse = etree.HTML(response.text)  # 转换为页面树
        # data_participant_name = parse.xpath('//div[@id="pnlResultNormal"]//table//td[@class="col-participant-name"]//div[@class="mobile-list-body"]/text()')#券商名
        #data_sharehoding=parse.xpath('//div[@id="pnlResultNormal"]//table//td[@class="col-shareholding text-right"]//div[@class="mobile-list-body"]/text()')#持股量
        #data_sharehoding_percent=parse.xpath('//div[@id="pnlResultNormal"]//table//td[@class="col-shareholding-percent text-right"]//div[@class="mobile-list-body"]/text()')#占已发行股份/权证/单位百分比
        #print(data_participant_name,data_sharehoding,data_sharehoding_percent)#验证输出

        for i in range(len(parse.xpath('//div[@id="pnlResultNormal"]//table//td[@class="col-participant-name"]//div[@class="mobile-list-body"]/text()'))):
            participant_name.append(parse.xpath('//div[@id="pnlResultNormal"]//table//td[@class="col-participant-name"]//div[@class="mobile-list-body"]/text()')[i])
            shareholding.append(parse.xpath('//div[@id="pnlResultNormal"]//table//td[@class="col-shareholding text-right"]//div[@class="mobile-list-body"]/text()')[i])
            sharehoding_percent.append(parse.xpath('//div[@id="pnlResultNormal"]//table//td[@class="col-shareholding-percent text-right"]//div[@class="mobile-list-body"]/text()')[i]) #将每个元素放入列表
        # print(participant_name)
        # print(shareholding)
        # print(sharehoding_percent)#验证xpath元素爬取



    def getYesterday(self):
        today = datetime.date.today()
        oneday = datetime.timedelta(days=1)
        yesterday = today - oneday
        return yesterday

    def file_save(self,participant_name_final,shareholding_final,sharehoding_percent_final,input_stock_code,input_date_data):#在D盘新建文件夹并将信息保存进excel

        input_date_data=str.replace(input_date_data,'/','-')#替换‘/’否侧会导致路径解析错误
        if os.path.exists(self.path)==True:#判断路径存在
            if os.path.exists(self.path+'{}'.format(input_date_data)+'_'+input_stock_code+'.xls'):#判断文件存在
                os.system(self.path+'{}'.format(input_date_data)+'_'+input_stock_code+'.xls')#自动打开存在路径下已有的excel
                return 0
            if os.path.exists(self.path+'{}'.format(input_date_data)+'_'+input_stock_code+'.xls')==False:#误删文件后重新爬取下载
                workbook=xlwt.Workbook(encoding='utf-8')
                data_sheet=workbook.add_sheet("{}".format(input_date_data))
                data_sheet.write(0,0,label='券商名')
                data_sheet.write(0,1,label='持股量')
                data_sheet.write(0,2,label='持股差额')

                for i in range(len(participant_name_final)):
                    data_sheet.write(i+1,0,label=participant_name_final[i])
                    data_sheet.write(i+1,1,label=shareholding_final[i])
                    data_sheet.write(i+1,2,label=sharehoding_percent_final[i])
                workbook.save(self.path+'{}'.format(input_date_data)+'_'+input_stock_code+'.xls')
                os.system(self.path+'{}'.format(input_date_data)+'_'+input_stock_code+'.xls')#打开文件
        elif os.path.exists(self.path)==False:
            os.makedirs(self.path)
            workbook=xlwt.Workbook(encoding='utf-8')
            data_sheet=workbook.add_sheet("{}".format(input_date_data))
            data_sheet.write(0,0,label='券商名')
            data_sheet.write(0,1,label='持股量')
            data_sheet.write(0,2,label='流通比')

            for i in range(len(participant_name_final)):
                data_sheet.write(i+1,0,label=participant_name_final[i])
                data_sheet.write(i+1,1,label=shareholding_final[i])
                data_sheet.write(i+1,2,label=sharehoding_percent_final[i])
            workbook.save(self.path+'{}'.format(input_date_data)+'_'+input_stock_code+'.xls')
            os.system(self.path+'{}'.format(input_date_data)+'_'+input_stock_code+'.xls')#生成并打开excel

    def input_stockCode(self):#输入获取的股票代码

        while True:
            box=simpledialog.askstring(title='披露易数据爬取',prompt='请输入要爬取数据的股票代码：(格式：0XXXX)',initialvalue='0')
            if (str(box)!=None and len(box)==5 and str(box).isdigit()==True):#保证用户输入内容非空长度5且都为数字
                stockcode=str(box)
                return stockcode
            elif (len(box)<5 or len(box)>5 or str(box).isdigit()!=True):
                warning=simpledialog.messagebox.showerror(title='严重警报',message='股票代码格式有误,唔好乱鬼咁输啊！')


    def main_window(self):#主窗口

            window=Tk()
            window.geometry("750x350+380+180")#主界面窗口显示
            window.title('披露易数据爬虫')
            button1=tkinter.Button(window,text='获取当天披露易持股信息',command=dataget.getData())#获取当天披露易持股信息
            button2=tkinter.Button(window,text='获取日期持股差额',bg='SkyBlue',command=dataget.getBalance())
            button_exit=tkinter.Button(window,text='退出程序',command=window.quit())
            button1.grid(row=0,column=0)
            button2.grid(row=0,column=2)
            button_exit.grid(row=2,column=6)


    def getData(self):
        input_stock_code=dataget.input_stockCode()
        input_date=dataget.input_Date()
        dataget.get_stockdata(input_date,input_stock_code)
        dataget.file_save(participant_name_final=participant_name,shareholding_final=shareholding,sharehoding_percent_final=sharehoding_percent,input_stock_code=input_stock_code,input_date_data=input_date)

    def input_Date(self):
        while True:
            box=simpledialog.askstring(title='披露易数据爬取',prompt='请输入要爬取数据的日期：(格式：XXXX/XX/XX)',initialvalue='2022/')
            if (str(box)!=None and len(box)==10):#保证用户输入内容非空长度为10且无字母
                data_date=str(box)
                return data_date
            elif (len(box)<10 or len(box)>10 or str(box).isalpha()!=True):
                warning=simpledialog.messagebox.showerror(title='严重警报',message='日期格式有误,唔好乱鬼咁输啊！')

    def getBalance(self):
        pass



if __name__ == "__main__":
    dataget = HKex_Search()#创建对象
    dataget.main_window()#进入程序主窗口

'''1.返回当天披露易数据
   2.返回数据差额     '''
