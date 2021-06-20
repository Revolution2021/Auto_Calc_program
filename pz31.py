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

mail.to = 'userx@gmail.com;usery@hotmail.com'  #Input real email destination
mail.cc = 'userz@gmail.com'
mail.subject = 'Billing Charge' 
mail.bodyFormat = 1
mail.body = data

f.close()


#mail.display(True)
mail.Send()

#pyautogui.press('f9')