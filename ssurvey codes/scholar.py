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

def scrape_references():
    
    titles = pd.read_csv(r"C:\Users\Msi\OneDrive\Research\Survey process\db\fine-grained search\based on words\final_titles.csv")
    
    # titles  = titles.loc[((titles['Year'] == 2017) | (titles['Year'] == 2016) | (titles['Year'] == 2015) | (titles['Year'] == 2014) |(titles['Year'] == 2013) | (titles['Year'] == 2012) ) ]
                      
    
    # Extract References from scopus file
    
    df_2_filtered = titles['References']
    
    l_s = []
    
    for i in df_2_filtered:
        i = str(i)
        l_s.append(i.split(";"))
        
        
    ar_titles=[]    
    for items in l_s:
        for  item in items:
            ar_titles.append(item) 
    
    
    ar_titles = [x for x in ar_titles if x != 'nan']
    
    # End of Extract References from scopus file
    
    # Scrape google scholar
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", r'D:\SOFTWARE\db\downloads')
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/xml; application/json;text/csv;charset=UTF-8")
    
    
    path = r"C:/Users/Msi/Desktop/geckodriver.exe"
    driver = webdriver.Firefox(firefox_profile=profile, executable_path=path)
    
    
    t_of_a = []
    driver.get("""https://scholar.google.com/""")
    
    
    driver.find_element_by_xpath("/html/body/div/div[7]/div[1]/div[2]/form/div/input").send_keys(ar_titles[5000])
    
    element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "gs_hdr_tsb")))
    element.click()
    
    time.sleep(5)
    try:
        element = driver.find_element(By.XPATH,'/html/body/div[1]/div[10]/div[2]/div/form/h1')
        print("waiting")
        time.sleep(70)
        element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "ZJ-fn3w_QLIJ")))
        t_of_a.append(element.text)
    except:
        print("error")    
    
    
    for index,item in enumerate(ar_titles):
       print(index)
       driver.find_element_by_xpath("/html/body/div/div[8]/div[1]/div/form/div/input").clear()
       time.sleep(1)
       driver.find_element_by_xpath("/html/body/div/div[8]/div[1]/div/form/div/input").send_keys(item)
    
       element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "gs_hdr_tsb")))
       element.click()
       time.sleep(1)
       try:
           element = WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[10]/div[2]/div[3]/div[2]/div[1]/div/h3/a')))
           t_of_a.append(element.text)
       except:
           continue
        
    time.sleep(5)    
    driver.close()

# End of Scrape google scholar

df_t_of_a = pd.DataFrame(t_of_a, columns=["Title"])

df_cc = final_df_all_fine_grained['Title'].value_counts(dropna=False)

df_cc.to_excel(r'C:\Users\Msi\OneDrive\Research\Survey process\db\fine-grained search\based on words\refs_count.xlsx')

titles = pd.read_excel(r"C:\Users\Msi\OneDrive\Research\Survey process\db\fine-grained search\based on words\refs_count.xlsx")
titles = titles.loc[titles['Count']>9]
titles = titles["Title"].tolist()
search_term=""
for title in titles:
    if title == titles[-1]:
        search_term+='OR TITLE("{0}")'.format(title)
    else:
        search_term+='TITLE("{0}") OR '.format(title)




# 



# Drop rows based on condition
df_ar_titles = pd.DataFrame(ar_titles,columns=["title"])
df_ar_titles = df_ar_titles.drop(df_ar_titles[~df_ar_titles['title'].str.contains("2018")].index)
ar_titles = df_ar_titles['title'].tolist()
# End of Extract References from scopus file



# 
df_refs = pd.DataFrame(ar_titles,columns=["refs"])


df_refs['refs'] = df_refs['refs'].str[:28]
df_refs['refs'] = df_refs['refs'].str.replace(r"\(.*\)","")
df_refs['refs'] = df_refs['refs'].str.replace(r"\(.*\)","")

df_refs = df_refs.drop(df_refs[df_refs['refs'].str.contains("Proceedings")].index)


titles['Titles'] = titles['Title'].str[:28]


result = df_refs.merge(title, how="inner", left_on='refs', right_on='Title'  )

fd_add = pd.read_excel(r"C:\Users\Msi\Desktop\add.xlsx")

fd_list = fd_add["Title_x"].tolist()
search_term=""
for title in t_of_a:
    if title == t_of_a[-1]:
        search_term+='OR TITLE("{0}")'.format(title)
    else:
        search_term+=' OR TITLE("{0}")'.format(title)
        
        
        
        
rw=[]
for item in r: 
    rw.append(str(r).replace(" ","",1))
# 