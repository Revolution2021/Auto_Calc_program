import win32com.client
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pyautogui
import os

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
dirname=os.getcwd()
file7=os.path.join(dirname,'samplex.txt')

#f = open('samplex.txt', 'r')
f = open(file7, 'r',encoding='utf-8')

data = f.read()
print(data)
pre_month=datetime.strftime(datetime.today()-relativedelta(months=1),"%m")

mail.to = 'userx1@gmail.com;usery@hotmail.com'
mail.cc = 'userz@gmail.com'
mail.subject = 'IMSI提供サービス (GBD担当) 従量変動単価案件の請求 2021年'+pre_month+'月分'
mail.bodyFormat = 1
mail.body = data

f.close()


#mail.display(True)
mail.Send()

#pyautogui.press('f9')