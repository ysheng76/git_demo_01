from datetime import datetime
import schedule
import time

TONG_XIN_ETF='519820'
TONG_XIN_AMOUNT='200'
full=False
now = datetime.now().time().hour
print(now)
i=32
i=i+1

def job():
    print("do")
schedule.every().day.at("19:43").do(job)

while True :
   now = datetime.now().time().minute
   if now==46 :
    print("time")

    

      
   
    



