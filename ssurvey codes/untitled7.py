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




df_1 = pd.read_csv(r"C:\Users\Msi\Downloads\scopus(10).csv")

df_2 = pd.read_excel(r"C:\Users\Msi\OneDrive\Research\Survey process\db\fine-grained search\based on words\refs_final14000721.xlsx")

df_2=df_2["References"]

df_2_filtered  = df_2.loc[((df_2['Year'] == 2020) | (df_2['Year'] == 2021) ) ]

df_2_filtered = df_2_filtered[["Title","References"]]

df_2_filtered = df_2_filtered[df_2_filtered['References'].notna()]

df_2_filtered["References"] = df_2_filtered["References"].str.split(';')

dic={}

for index, row in df_2_filtered.iterrows():
    dic[row["Title"]]=row["References"]

df_3 = df_2.merge(df_1, how="outer", left_on="Link", right_on="Link")

df_1.to_excel(r"C:\Users\Msi\Downloads\1.xlsx")
boolean = df_1['Link'].duplicated().any()


df_3_f = df_3[~df_3['Title_x'].notnull()]

df_3_f = df_3_f.dropna(axis=1, how='all')


fd_add = pd.read_excel(r"C:\Users\Msi\OneDrive\Research\Survey process\db\fine-grained search\based on words\55.xlsx")

df_2 = df_2["Title"].tolist()
search_term=""
for title in df_2:
    if title == df_2[-1]:
        search_term+='OR TITLE("{0}")'.format(title)
    else:
        search_term+=' OR TITLE("{0}")'.format(title)
        
        
df_4 = pd.DataFrame(t_f_a, columns="titles")        

print(t_f_a[0])