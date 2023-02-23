from pywinauto.application import Application
from time import sleep
from pywinauto.keyboard import send_keys
import win32gui
import win32con
import win32api
import easyquotation
import urllib
import pickle
import json
import schedule
from datetime import datetime


###操作规范:
###设置好DO_LIST里的内容(3个)->手动撤单->调好设置(主要设置)->并观察程序是否正在运行(鼠标有选中,控制台是否在打印)

#full值每天要手动修改
TONG_XIN=['515880','100',False]
TONG_XIN_FULL=False
DO_LIST=[TONG_XIN]




#获取5日均线报价
def getEma5(market,code):
    #注意etf是sz还是sh
#http://money.finance.sina.com.cn/quotes_service/api/json_v2.php/CN_MarketData.getKLineData?symbol=[市场][股票代码]&scale=[周期]&ma=no&datalen=[长度]
   baseUrl='http://money.finance.sina.com.cn/quotes_service/api/json_v2.php/CN_MarketData.getKLineData?symbol='
   baseUrl1='&scale=15&ma=no&datalen=5'
   dataUrl=baseUrl+market+code+baseUrl1
   allDataJson= urllib.request.urlopen(dataUrl).read().decode('utf-8')
   allData=json.loads(allDataJson)
   sum=0.0
   print(allData)
   for item in allData:
        sum+=float(item['close'])
   ave=sum/5.0
   log("ema5:"+str(ave))
   return ave


def getLast(market,code):
    #注意etf是sz还是sh
#http://money.finance.sina.com.cn/quotes_service/api/json_v2.php/CN_MarketData.getKLineData?symbol=[市场][股票代码]&scale=[周期]&ma=no&datalen=[长度]
   baseUrl='http://money.finance.sina.com.cn/quotes_service/api/json_v2.php/CN_MarketData.getKLineData?symbol='
   baseUrl1='&scale=15&ma=no&datalen=1'
   dataUrl=baseUrl+market+code+baseUrl1
   allDataJson= urllib.request.urlopen(dataUrl).read().decode('utf-8')
   allData=json.loads(allDataJson)
   print(allData)
   for item in allData:
       log("nowprice:"+item['close'])
       return float(item['close'])


#获取实时报价
def getNow(code):

    quotation=easyquotation.use('sina')
    data=quotation.real(code)
    ans=data[code]['now']
    log("nowprice:"+ans)
    return ans





class Trader:
    @staticmethod
    def buy(code,amount):
        log("start___buy:")    
        log("code=="+code)
        log("amount=="+amount)
        #读取交易界面句柄，获得窗口
        hwnd = win32gui.FindWindow(None, "网上股票交易系统5.0")
        app = Application().connect(handle=hwnd)
        win = app.top_window()
     
        win.set_focus()
        #确定切换到买入界面
        send_keys('{F2}')
        send_keys('{F1}')
        win.set_focus()
    
        #进行买入操作
        send_keys(code)
        send_keys('{ENTER}')
        sleep(0.3)
        send_keys('{ENTER}')
        send_keys(amount)
        send_keys('{ENTER 2}')
        log("end_buy")
 
 
       
    @staticmethod
    def sale(code,amount):
          #读取交易界面句柄，获得窗口
        log("start_sale")
        hwnd = win32gui.FindWindow(None, "网上股票交易系统5.0")
        app = Application().connect(handle=hwnd)
        win =app.top_window()
        win.set_focus()
        #确定切换到买入界面
        send_keys('{F2}')
        win.set_focus()
        #进行买入操作
        send_keys(code)
        send_keys('{ENTER}')
        sleep(0.3)
        send_keys('{ENTER}')
        send_keys(amount)
        send_keys('{ENTER 2}')
    

    @staticmethod
    def refresh():
        log("refresh")
        hwnd = win32gui.FindWindow(None, "网上股票交易系统5.0")
        app = Application().connect(handle=hwnd)
        win =app.top_window()
  
        win.set_focus()
        send_keys('{F2}')
        send_keys('{F1}')
        #获取刷新界面句柄
        refreshPage1 = win32gui.FindWindowEx(hwnd,None,'ToolbarWindow32',None)
        refreshPage2= win32gui.FindWindowEx(refreshPage1,None,'#32770',None)
        #模仿鼠标点击刷新按钮，这里刷新按钮的位置不是严格的抓抓中读取的位置
        position = win32api.MAKELONG(500,30) #x,y为点击点相对于该窗口的坐标
        win32api.SendMessage(refreshPage2,win32con.WM_LBUTTONDOWN,win32con.MK_LBUTTON,position)#向窗口发送模拟鼠标点击
        win32api.SendMessage(refreshPage2,win32con.WM_LBUTTONUP,win32con.MK_LBUTTON,position)
        #鼠标定位到代码输入框
        win.set_focus()
        send_keys('{F2}')
        send_keys('{F1}')
       
    
        #清空代码框


#开始计算并进行交易    
def job():
    log("do___the___job")
    for item in DO_LIST:
        code=item[0]
        amount=item[1]
        full=item[2]
        ema5=getEma5("sh",code)
        nowPrice=getLast("sh",code)
        if nowPrice>=ema5:
            if full:
                log("buy--doNothing")
                return
            else:
                Trader.buy(code,amount)
                item[2]=True
        else:
            if full:
                Trader.sale(code,amount)
                item[2]=False
            else:
                log("sale---doNothing")
                return


def self_test() :
    schedule.every(10).seconds.do(Trader.refresh)
    schedule.every(20).seconds.do(job)
    for item in DO_LIST:
        code=item[0]
        amount=item[1]
        full=item[2]
        Trader.buy(code,amount)
        Trader.sale(code,amount)
    


def log(message):
    print(str(datetime.now().time())+"---"+message)



if __name__ == '__main__':
    log("start___the___process")
 
    #添加定时刷新业务
    minRange=[':04',':09',':14',':19',':24',':29',':34',':39',':44',':49',":54",':59']
    for min in minRange:
         schedule.every().hour.at(min).do(Trader.refresh)
    
    #self_test()

    #添加定时交易业务
    schedule.every().day.at("09:44:40").do(job)
    schedule.every().day.at("09:59:40").do(job)
    schedule.every().day.at("10:14:40").do(job)
    schedule.every().day.at("10:29:40").do(job)
    schedule.every().day.at("10:44:40").do(job)
    schedule.every().day.at("10:59:40").do(job)
    schedule.every().day.at("11:14:40").do(job)
    schedule.every().day.at("11:29:40").do(job)
    schedule.every().day.at("13:14:40").do(job)
    schedule.every().day.at("13:29:40").do(job)
    schedule.every().day.at("13:44:40").do(job)
    schedule.every().day.at("13:59:40").do(job)
    schedule.every().day.at("14:14:40").do(job)
    schedule.every().day.at("14:29:40").do(job)
    schedule.every().day.at("14:44:40").do(job)
    schedule.every().day.at("14:56:40").do(job)

    log("start___the___process")
    log("DO_LIST----")
    print(DO_LIST)
    while True :
        schedule.run_pending()
        

  
 
   


