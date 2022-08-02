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
from selenium.common.exceptions import NoSuchElementException        

titles = pd.read_csv(r"C:\Users\Msi\OneDrive\Research\Survey process\db\fine-grained search\based on words\final_titles.csv")

# titles  = titles.loc[((titles['Year'] == 2017) | (titles['Year'] == 2016) | (titles['Year'] == 2015) | (titles['Year'] == 2014) |(titles['Year'] == 2013) | (titles['Year'] == 2012) ) ]
                  

# Extract References from scopus file

df_2_filtered = titles['Title']

ar_titles=[]    

for i in df_2_filtered:
    ar_titles.append(i) 
   

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


driver.find_element_by_xpath("/html/body/div/div[7]/div[1]/div[2]/form/div/input").send_keys("self-adaptive")

element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "gs_hdr_tsb")))
element.click()

try:
    element = driver.find_element(By.ID, 'gs_captcha_f')

    while(element.is_displayed()):
        print("waiting")
except NoSuchElementException:
    print("hello")
time.sleep(2) 

for index,item in enumerate(ar_titles[80:]):
   print(index)
   driver.find_element_by_id("gs_hdr_tsi").clear()
   time.sleep(1)
   driver.find_element_by_id("gs_hdr_tsi").send_keys(item)

   element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "gs_hdr_tsb")))
   element.click()
   try:
    element = driver.find_element(By.ID, 'recaptcha-anchor-label')

    while(element.is_displayed()):
        print("waiting")
   except NoSuchElementException:
        try:
           element = WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.XPATH, '//*[@title="Cite"]')))
           element.click()
           
           time.sleep(1)
           
           element = WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[4]/div/div[2]/div/div[1]/table/tbody/tr[4]/td/div')))
           t_of_a.append(element.text)
           
           element = WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[4]/div/div[1]/a/span[1]')))
           element.click()
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