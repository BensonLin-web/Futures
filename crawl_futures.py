from selenium import webdriver
from selenium.webdriver.support.ui import Select
from time import sleep as sl
import pymongo
from datetime import datetime as dt
import re
import csv
import numpy as np 
import matplotlib.pyplot as mp
import matplotlib.dates as md
import seaborn as sns



class FuturesCrawl(object):
    def __init__(self):
        self.url = "https://www.taifex.com.tw/cht/3/futDailyMarketReport"
        self.conn = pymongo.MongoClient("localhost",27017)
        self.db = self.conn.futures 
        self.collection = self.db.txDecJanFeb
    
    #獲取頁面
    def getPage(self):
        # 設置無介面瀏覽
        opt = webdriver.ChromeOptions()
        opt.set_headless()
        self.driver = webdriver.Chrome(options = opt)
        self.driver.get(self.url)
        sl(1)
        #修改查詢日期
        queryDate = self.driver.find_element_by_name("queryDate")
        queryDate.clear()
        queryDate.send_keys(r"2019/12/02")
        self.driver.find_element_by_name("button").click()
        sl(1)
        #選取交易時段
        MarketCode = Select(self.driver.find_element_by_id("MarketCode"))
        MarketCode.select_by_value("0")
        self.driver.find_element_by_name("button").click()
        sl(1)
        self.parsePage()

    #解析頁面
    def parsePage(self):
        #新增csv文件，寫入爬取過程的狀態
        for i in range(31):
            if i == 0:
                with open("txStatus.csv","a+",newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["date","status"])
            try:
            	#獲取內容節點對象列表
                r_list = self.driver.find_elements_by_xpath('//div[@id="printhere"]/table/tbody/tr[2]/td/table[2]/tbody/tr')
                #獲取日期的節點對象
                date = self.driver.find_element_by_xpath('//div[@id="printhere"]/table/tbody/tr[2]/td/h3')
                dateString = re.findall(r'\d+[/]\d+[/]\d+',date.text)[0]
                print(dateString,"crawler walking now")
                x_list = []
                for r in r_list:
                	x_list.append(r.text.split())
                #定義標題列表
                title_list = x_list[0]
                #合併相關標題
                title_list[1] = "".join(title_list[1:4])
                del title_list[2:4]
                title_list[5] = "".join(title_list[5:7])
                del title_list[6]
                #標題欄新增日期
                title_list.append("日期")

                del x_list[-1]
                x_list = [i + [dateString] for i in x_list]
                
                #存進Mongo數據庫
                self.saveToMongo(x_list,title_list)
                print(dateString,"insert to Mongo success!")
                with open("txStatus.csv","a+",newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow([dt.now(),"sucess"])
                #跳轉到後一日
                button = self.driver.find_element_by_id("button4")
                button.click()
                sl(2)
            
            #例假日沒有期貨資訊，用except接收異常
            except Exception as e:
                with open("txStatus.csv","a+",newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow([dt.now(),"error:" + str(e)])
            
                button = self.driver.find_element_by_id("button4")
                button.click()
                sl(2)
    
    def saveToMongo(self,x_list,title_list):
    	for i in x_list[1:]:
    		D = {}
    		for title,data in zip(title_list,i):
    			D[title] = data
    		self.collection.insert_one(D)

    def loadTx(self):
        #從數據庫讀取資料
        #datas為一個cursor對象
        datas = self.collection.find({"到期月份(週別)":"202006"},{"_id":0})
        #最高價
        highest_prices = np.zeros(datas.count(),dtype="int")
        #最低價
        lowest_prices = np.zeros(datas.count(),dtype="int")
        #開盤價
        opening_prices = np.zeros(datas.count(),dtype="int")
        #收盤價
        closing_prices = np.zeros(datas.count(),dtype="int")
        #日期
        dates = np.zeros(datas.count(),dtype="U10")
        i = 0
        for data in datas:
            highest_prices[i] = data["最高價"]
            lowest_prices[i] = data["最低價"]
            opening_prices[i] = data["開盤價"]
            closing_prices[i] = data["最後成交價"]
            #將'2019/12/02' -> '2019-12-02'
            dates[i] = "-".join(data["日期"].split("/"))
            i += 1
        #將日期轉為datetime類型
        dates = dates.astype(np.datetime64)
        # 折線圖
        self.plotTx(dates,highest_prices,lowest_prices)
        # K線圖，移動平均線
        self.kTx(dates,highest_prices,lowest_prices,opening_prices,closing_prices)
        # K線圖，布林通道
        self.ebbTx(dates,highest_prices,lowest_prices,opening_prices,closing_prices)
  
    def plotTx(slef,dates,highest_prices,lowest_prices):
        #設置字體，圖形樣式
        #whitegrid設置白色網格線，但X軸Y軸刻度會消失
        # sns.set_style("whitegrid")
        mp.rcParams['font.sans-serif'] = ['Microsoft YaHei'] 
        mp.rcParams['font.family']='sans-serif'
        #設置正常顯示文字(如有負號可顯示)
        mp.rcParams['axes.unicode_minus'] = False

        #設置圖形窗口
        mp.figure("Futures",facecolor="lightgray")
        mp.title("2019/12~2020/02 到期月份 202006 TX 最高最低價折線圖",fontsize=30)
        mp.xlabel("日期",fontsize=20)
        mp.ylabel("價格",fontsize=20)
        ax = mp.gca()
        
        # 以下為一個月數據的X軸的座標刻度，需切換到數據表self.db.txDec
        # #設置每一天為主刻度
        # ax.xaxis.set_major_locator(md.DayLocator())
        # #設置水平座標主刻度的標籤格式
        # ax.xaxis.set_major_formatter(md.DateFormatter('%m/%d'))
        
        # 以下為三個月數據的X軸的座標刻度，需切換到數據表self.db.txDecJanFeb
        # 設置每個星期一為主刻度
        ax.xaxis.set_major_locator(md.WeekdayLocator(byweekday=md.MO))
        #設置每一天為次刻度
        ax.xaxis.set_minor_locator(md.DayLocator())
        #設置水平座標主刻度的標籤格式
        ax.xaxis.set_major_formatter(md.DateFormatter('%d %b %Y'))
        mp.tick_params(labelsize=10)
        mp.grid(linestyle=":")
        
        #畫折線圖
        mp.plot(dates,highest_prices,linestyle="--",label="最高價",marker="o")
        mp.plot(dates,lowest_prices,linestyle="--",label="最低價",marker="o")
        #為每個點添加數據標籤
        for d,h,l in zip(dates,highest_prices,lowest_prices):
            mp.text(d,h,h,ha='center', va='bottom', fontsize=10)
            mp.text(d,l,l,ha='center', va='bottom', fontsize=10)
        mp.legend(loc="upper left")
        #自動調整水平座標軸的日期標籤
        mp.gcf().autofmt_xdate()
        mp.show()

    def kTx(self,dates,highest_prices,lowest_prices,opening_prices,closing_prices):
        #設置字體，圖形樣式
        #whitegrid設置白色網格線，但X軸Y軸刻度會消失
        # sns.set_style("whitegrid")
        mp.rcParams['font.sans-serif'] = ['Microsoft YaHei'] 
        mp.rcParams['font.family']='sans-serif'
        #設置正常顯示文字(如有負號可顯示)
        mp.rcParams['axes.unicode_minus'] = False
        
        mp.figure('Futures',facecolor='lightgray')
        mp.title('2019/12~2020/02 到期月份 202006 TX K線圖',fontsize=30)
        mp.xlabel('日期',fontsize=20)
        mp.ylabel('價格',fontsize=20)
        mp.ylim(11200,12100)
        ax = mp.gca()
        
        # 以下為一個月數據的X軸的座標刻度，需切換到數據表self.db.txDec
        # #設置每一天為主刻度
        # ax.xaxis.set_major_locator(md.DayLocator())
        # #設置水平座標主刻度的標籤格式
        # ax.xaxis.set_major_formatter(md.DateFormatter('%m/%d'))
        
        # 以下為三個月數據的X軸的座標刻度，需切換到數據表self.db.txDecJanFeb
        # 設置每個星期一為主刻度
        ax.xaxis.set_major_locator(md.WeekdayLocator(byweekday=md.MO))
        #設置每一天為次刻度
        ax.xaxis.set_minor_locator(md.DayLocator())
        #設置水平座標主刻度的標籤格式
        ax.xaxis.set_major_formatter(md.DateFormatter('%d %b %Y'))
        mp.tick_params(labelsize=10)
        mp.grid(linestyle=':')

        #np.ones(5):產生5個1的數組
        #使用摺積計算5天移動平均線的值
        ma5 = np.convolve(closing_prices,np.ones(5) / 5,'valid')
        #使用摺積計算10天移動平均線的值
        ma10 = np.convolve(closing_prices,np.ones(10) / 10,'valid')

        #陽線
        rise = closing_prices - opening_prices >= 1
        #陰線
        fall = opening_prices - closing_prices >= 1
        #填充色
        fc = np.zeros(dates.size,dtype='3f4')
        fc[rise],fc[fall] = (1,0,0),(0,0.5,0)
        #邊緣色
        ec = np.zeros(dates.size,dtype='3f4')
        ec[rise],ec[fall] = (1,0,0),(0,0.5,0)
        
        #K線上、下影線
        mp.bar(dates,highest_prices - lowest_prices,0,lowest_prices,color=fc,edgecolor=ec)
        #K線實體
        mp.bar(dates,closing_prices - opening_prices,0.3,opening_prices,color=fc,edgecolor=ec)
        #收盤價
        mp.plot(dates,closing_prices,c='lightgray',label='Closing Price')
        #5日移動平均線
        mp.plot(dates[4:],ma5,c='limegreen',alpha=0.5,linewidth=6,label='MA-5')
        #10日移動平均線
        mp.plot(dates[9:],ma10,c='dodgerblue',label='MA-10')
        #為收盤價添加數據標籤
        for d,c in zip(dates,closing_prices):
            mp.text(d,c,c,ha='center', va='bottom', fontsize=10)
        mp.legend(loc="upper left")
        mp.gcf().autofmt_xdate()
        mp.show()

    def ebbTx(self,dates,highest_prices,lowest_prices,opening_prices,closing_prices):
        #設置字體，圖形樣式
        #whitegrid設置白色網格線，但X軸Y軸刻度會消失
        # sns.set_style("whitegrid")
        mp.rcParams['font.sans-serif'] = ['Microsoft YaHei'] 
        mp.rcParams['font.family']='sans-serif'
        #設置正常顯示文字(如有負號可顯示)
        mp.rcParams['axes.unicode_minus'] = False
        
        mp.figure('Futures',facecolor='lightgray')
        mp.title('2019/12~2020/02 到期月份 202006 TX 布林通道',fontsize=30)
        mp.xlabel('日期',fontsize=20)
        mp.ylabel('價格',fontsize=20)

        #一個月數據的Y軸的值，需切換到數據表self.db.txDec
        # mp.ylim(11200,12200)
        
        #三個月數據的Y軸的值，需切換到數據表self.db.txDecJanFeb
        mp.ylim(10600,12400)
        ax = mp.gca()
        
        # 以下為一個月數據的X軸的座標刻度，需切換到數據表self.db.txDec
        # # #設置每一天為主刻度
        # ax.xaxis.set_major_locator(md.DayLocator())
        # #設置水平座標主刻度的標籤格式
        # ax.xaxis.set_major_formatter(md.DateFormatter('%m/%d'))
        
        # 以下為三個月數據的X軸的座標刻度，需切換到數據表self.db.txDecJanFeb
        # 設置每個星期一為主刻度
        ax.xaxis.set_major_locator(md.WeekdayLocator(byweekday=md.MO))
        #設置每一天為次刻度
        ax.xaxis.set_minor_locator(md.DayLocator())
        #設置水平座標主刻度的標籤格式
        ax.xaxis.set_major_formatter(md.DateFormatter('%d %b %Y'))
        mp.tick_params(labelsize=10)
        mp.grid(linestyle=':')
        
        #np.linspace():通過定義間隔創建數值序列(包含-1和0)
        #np.exp(x):求e**x的值函數，e為自然常數2.71828
        weights = np.exp(np.linspace(-1,0,5))
        weights /= weights.sum()
        #使用摺積算出5天的指數移動平均的值
        #因為需符合遞增加權(時間越近，加權越重)，且摺積過程會自動將摺積核取反
        #所以需先將摺積核取反weights[::-1]
        medios = np.convolve(closing_prices,weights[::-1],'valid')
        #算出每個5天的標準差
        stds = np.zeros(medios.size)
        for i in range(stds.size):
            stds[i] = closing_prices[i:i+5].std()
        stds *= 2
        #壓力線 = 移動平均線 - 2倍標準差
        lowers = medios - stds
        #支撐線 = 移動平均線 + 2倍標準差
        uppers = medios + stds

        #陽線
        rise = closing_prices - opening_prices >= 1
        #陰線
        fall = opening_prices - closing_prices >= 1
        #填充色
        fc = np.zeros(dates.size,dtype='3f4')
        fc[rise],fc[fall] = (1,0,0),(0,0.5,0)
        #邊緣色
        ec = np.zeros(dates.size,dtype='3f4')
        ec[rise],ec[fall] = (1,0,0),(0,0.5,0)
        
        #K線上、下影線
        mp.bar(dates,highest_prices - lowest_prices,0,lowest_prices,color=fc,edgecolor=ec)
        #K線實體
        mp.bar(dates,closing_prices - opening_prices,0.3,opening_prices,color=fc,edgecolor=ec)
        #收盤價
        mp.plot(dates,closing_prices,c='lightgray',label='Closing Price')
        mp.plot(dates[4:],medios,c='dodgerblue',label='EMA-5')
        mp.plot(dates[4:],lowers,c='limegreen',label='支撐線')
        mp.plot(dates[4:],uppers,c='orangered',label='壓力線')
        #為收盤價添加數據標籤
        for d,c in zip(dates,closing_prices):
            mp.text(d,c,c,ha='center', va='bottom', fontsize=10)
        mp.legend(loc="upper left")
        mp.gcf().autofmt_xdate()
        mp.show()


    def workOn(self):
        # self.getPage()
        self.loadTx()
        # self.driver.quit()


if __name__ == "__main__":
    spider = FuturesCrawl()
    spider.workOn() 