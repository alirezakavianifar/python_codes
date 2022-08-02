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

titles = pd.read_csv(r"C:\Users\Msi\OneDrive\Research\Survey process\db\fine-grained search\based on words\final_titles.csv")                
k=18000
# Extract References from scopus file

df_2_filtered = titles[['Title','Year','References']]

df_2_filtered["References"] = df_2_filtered["References"].str.split(";")

dic={}

for index, row in df_2_filtered.iterrows():
    dic[row["Title"]] = row["References"]
    
dictlist=[]    
for key, value in dic.items():
    temp = [key,value]
    dictlist.append(temp)    

lst=[]        
for item in chunks(dic, 15):
    lst.append(item)       
    



# Scrape google scholar
profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir", r'D:\SOFTWARE\db\downloads')
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/xml; application/json;text/csv;charset=UTF-8")


path = r"C:/Users/Msi/Desktop/geckodriver.exe"
driver = webdriver.Firefox(firefox_profile=profile, executable_path=path)
driver.get("""https://scholar.google.com/""")  
driver.find_element_by_xpath("/html/body/div/div[7]/div[1]/div[2]/form/div/input").send_keys("self-adaptive")     
element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "gs_hdr_tsb")))
element.click()
time.sleep(1)
try:
    element = driver.find_element(By.XPATH,'/html/body/div[1]/div[10]/div[2]/div/form/h1')
    print("waiting")
    time.sleep(70)
    element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "ZJ-fn3w_QLIJ")))
except:
    print("error")  
i=0    
for key, value in lst[9].items():
    dic_tmp = {}
    lst_tmp = []
    for index,item in enumerate(value):
        print(index)
        driver.find_element_by_xpath("/html/body/div/div[8]/div[1]/div/form/div/input").clear()
        time.sleep(1)
        driver.find_element_by_xpath("/html/body/div/div[8]/div[1]/div/form/div/input").send_keys(item)
        
        element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "gs_hdr_tsb")))
        element.click()
        time.sleep(1)
        try:
            element = WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[10]/div[2]/div[3]/div[2]/div[1]/div/h3/a')))
            lst_tmp.append(element.text)
        except:
            continue
    dic_tmp[key] = lst_tmp
    
    df = pd.DataFrame.from_dict(dic_tmp)
    df.to_excel(r"D:\tmp\refs_of_{0}.xlsx".format(k))
    k=k+1    
    time.sleep(5)    

driver.close()        