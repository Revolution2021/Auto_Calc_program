import openpyxl as px
from datetime import datetime 
from dateutil.relativedelta import relativedelta
import os
import glob
from openpyxl.styles import Font
from time import sleep
import win32com.client


def cleandata():
	excel = win32com.client.Dispatch("Excel.Application")
	excel.DisplayAlerts = False
	
	# 絶対パスを取得
	filex=os.path.abspath("modified4x_summary.xlsx")
	filey=os.getcwd()

	#絶対パスの表示
	print(filex)
	print(filey)
	
#wb1=excel.Workbooks.Open(r"filex")
	wb1=excel.Workbooks.Open(filex)

#wb1=excel.Workbooks.Open(r"C:\Users\y-nishikawa\Desktop\testdata\modified4x_summary.xlsx")
	ws1=wb1.worksheets[0]

	xlUp=-4162
	lastrow = ws1.Cells(ws1.Rows.Count, 2).End(xlUp).Row
	print(lastrow)

	xlAscending = 1
	xlDescendig = 2
	xlYes = 1

	ws1.Range(ws1.Range("A2"),ws1.Cells(lastrow,8)).Sort(Key1=ws1.Range("H2"), Order1=xlDescendig, Header=xlYes)

	wb1.Save()
	wb1.Close()



