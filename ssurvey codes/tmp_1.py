import pandas as pd
import os
import time
import glob
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import math
import numpy as np
from itertools import islice

def chunks(data, SIZE=10000):
    it = iter(data)
    for i in range(0, len(data), SIZE):
        yield {k:data[k] for k in islice(it, SIZE)}


path = r'D:\tmp'


excel_files = glob.glob(os.path.join(path, "*.xlsx"))

merge_excels=[]

for f in excel_files:
    
    # read the csv file
    df = pd.read_excel(f)
    rrr= pd.DataFrame.to_dict(df)
    merge_excels.append(rrr)  
   
final_df_all_fine_grained = pd.concat(merge_excels)

fr = set(merge_excels[0].items()) & set(merge_excels[1].items())

# for item in merge_excels:

#     item.pop('Unnamed: 0', None)












a = { 'x' : 1, 'y' : 2, 'z' : 3 }
b = { 'u' : 1, 'v' : 2, 'w' : 3, 'x'  : 1, 'y': 2 }
 
setA = set( a )
setB = set( b )
 
tt=setA.intersection( setB )  


for item in setA.intersection(setB):
    print (item)
 









