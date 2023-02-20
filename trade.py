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



def start():
    try:
        hwnd = win32gui.FindWindow(None, "网上股票交易系统5.0")
        app = Application().connect(handle=hwnd)
        win = app.win = app.top_window()
        win.set_focus()
        send_keys('600512')
        send_keys('{ENTER}')
        send_keys('{ENTER}')
        send_keys('300')
        send_keys('{ENTER}')
        send_keys('{ENTER}')
        hwnd1 = 791354
        print(hwnd)
        print(hwnd1)
        win32api.SendMessage(hwnd, win32con.WM_SETTEXT, None, '519120')
        print("hello")
    except Exception as e:
         print(e)

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
       return float(item['close'])


#获取实时报价
def getNow(code):

    quotation=easyquotation.use('sina')
    data=quotation.real(code)
    ans=data[code]['now']
    return ans





class Trader:
    @staticmethod
    def buy(code,amount):
        #读取交易界面句柄，获得窗口
        hwnd = win32gui.FindWindow(None, "网上股票交易系统5.0")
        app = Application().connect(handle=hwnd)
        win = app.win = app.top_window()
        win.set_focus()
        #确定切换到买入界面
        send_keys('{F1}')
        win.set_focus()
        #进行买入操作
        send_keys(code)
        send_keys('{ENTER 2}')
        send_keys(amount)
        send_keys('{ENTER 2}')
       
   



    @staticmethod
    def sale(code,amount):
          #读取交易界面句柄，获得窗口
        hwnd = win32gui.FindWindow(None, "网上股票交易系统5.0")
        app = Application().connect(handle=hwnd)
        win = app.win = app.top_window()
        win.set_focus()
        #确定切换到买入界面
        send_keys('{F2}')
        win.set_focus()
        #进行买入操作
        send_keys(code)
        send_keys('{ENTER 2}')
        send_keys(amount)
        send_keys('{ENTER 2}')
    @staticmethod
    def refresh():
        #读取交易界面句柄，获得窗口
        hwnd = win32gui.FindWindow(None, "网上股票交易系统5.0")
        app = Application().connect(handle=hwnd)
        win = app.win = app.top_window()
        #获取刷新界面句柄
        refreshPage1 = win32gui.FindWindowEx(hwnd,None,'ToolbarWindow32',None)
        refreshPage2= win32gui.FindWindowEx(refreshPage1,None,'#32770',None)
        #模仿鼠标点击刷新按钮，这里刷新按钮的位置不是严格的抓抓中读取的位置
        position = win32api.MAKELONG(500,30) #x,y为点击点相对于该窗口的坐标
        win32api.SendMessage(refreshPage2,win32con.WM_LBUTTONDOWN,win32con.MK_LBUTTON,position)#向窗口发送模拟鼠标点击
        win32api.SendMessage(refreshPage2,win32con.WM_LBUTTONUP,win32con.MK_LBUTTON,position)
        #鼠标定位到代码输入框
        win.set_focus()
        #清空代码框
        send_keys('{BS 6}')
      
#开始计算并进行交易    
def job(full,code):
    ema5=getEma5()
    nowPrice=getLast()
    if nowPrice>=ema5 :
        if full:
            return
        else:
         Trader.sale()
    else:
        if full:
            Trader.sale
        else:
            return





if __name__ == '__main__':
   ans= getNow('512800')
   print(ans)
   


