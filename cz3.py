import pandas as pd
import glob
import os

def transtoexcel():

	filez3=glob.glob('copy_*.csv')
	print(filez3)
	print(len(filez3))
	max3=range(len(filez3))

	for i in max3:
		base, ext = os.path.splitext(filez3[i])

		read_file = pd.read_csv (base+"."+"csv")
		read_file.to_excel (base+"."+"xlsx", index = None, header=True)
	
	i=i+1