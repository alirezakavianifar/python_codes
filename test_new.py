import pandas as pd
import glob
import os
import random

pathsave = r'E:\test'
files=[]

def removeExcelFiles(pathsave):
    filelist = (glob.glob(os.path.join(pathsave, "*.xlsx")))
    for f in filelist:
        files.append((bool(random.getrandbits(1)),f))
        os.remove(f)


excels_to_be_removed=[]

for item in files:
    if not item[0]:
        excels_to_be_removed.append(item[1])
        

