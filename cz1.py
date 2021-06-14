import os
import pandas as pd
import shutil
import openpyxl as px
import glob
from time import sleep

def preparation():
	filez=glob.glob('*.csv')
	print(filez)

	max2=range(len(filez))

	for i in max2:
		base, ext = os.path.splitext(filez[i])
		if ext == '.csv':
			print(filez[i])

			p1="copy_"+base+"."+"csv"
			print(p1)
			shutil.copyfile(filez[i],p1)
			basename2_without_ext=os.path.splitext(filez[i])
			print(basename2_without_ext)
	


	
		i=i+1
			
 