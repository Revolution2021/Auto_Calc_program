import os
import pandas as pd
import csv
import codecs
import glob


def displayx():


	filez1=glob.glob('copy_XX*.csv')
	print(filez1)
	max2=range(len(filez1))

	for i in max2:
		base, ext = os.path.splitext(filez1[i])

		with codecs.open(base+"."+"csv","r","utf-8","ignore") as filex:
			df= pd.read_table(filex, delimiter=",")
			print(df)

		i=i+1
	


        