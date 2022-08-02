import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import os
import time
import glob
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import math

profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir", r'D:\SOFTWARE\db\downloads')
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/xml; application/json;text/csv;charset=UTF-8")


path = r"C:/Users/Msi/Desktop/geckodriver.exe"
driver = webdriver.Firefox(firefox_profile=profile)


driver.get("""https://web.telegram.org/""")
time.sleep(45)
driver.find_element(By.XPATH,"/html/body/div[1]/div[1]/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[1]/ul/li[1]/div[1]").click()
time.sleep(4)
#driver.find_element(By.XPATH,"/html/body/div[1]/div[1]/div[2]/div/div/div[4]/div/div[1]/div/div[7]/div[1]/div[1]").send_keys("مظنه")
messages = driver.find_elements(By.CLASS_NAME,"text-content")
texts=[]
for m in messages:
    texts.append(m.text)
element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div/div/div[4]/div/div[1]/div/div[7]/div[1]/div[1]"))) 
element.send_keys(Keys.RETURN)
    
element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "bdd-elsPrimaryBtn")))
element.click()
    
element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Advanced document search')]")))
element.click()
    
element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "contentEditLabel")))
element.click()
    
element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "searchfield")))
element.send_keys(search_term.format(2000))
    
element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, """/html/body/div[1]/div/div[1]/div[2]/div/div[3]/div/div[2]
                                                                          /div/form/div/div[1]/
                                                                          div/div/div[2]/div/section/div[2]/ul/li[5]/button/span[1]""")))
element.click()
    
time.sleep(3)

for year in years[1:]:
    
    # items_fount = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CLASS_NAME, "resultsCount")))
    # items_fount = int(items_fount.text)
    # pages = int(math.ceil(items_fount/200))
    
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "selectAllCheck")))
    element.click()
    
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/span[1]/button[2]/span")))
    element.click()
    
    time.sleep(4)
    
    # element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/div[4]/div/div/div[2]/div[2]/div/div/table/thead/tr/th[3]/span/label")))
    # element.click()
    
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/div[4]/div/div/div[2]/div[1]/div/ul/li[5]/label")))
    element.click()
    
    time.sleep(1)
    
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/div[4]/div/div/div[3]/div/button[2]/span")))
    element.click()
    
    time.sleep(3) 
    
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/nav/ul/li[1]/button/span[2]")))
    element.click()
    
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "searchfield")))
    element.clear()
    element.send_keys(search_term.format(year))
    
    element.send_keys(Keys.RETURN)
    time.sleep(4)
           
time.sleep(25)
driver.close()







# Merge Excel files

path = r'D:\SOFTWARE\db\downloads'

excel_files = glob.glob(os.path.join(path, "*.csv"))

merge_excels=[]

for f in excel_files:
    
    # read the csv file
    df = pd.read_csv(f)
    merge_excels.append(df)  
   
final_df_all_fine_grained = pd.concat(merge_excels)

final_df_all_fine_grained.drop_duplicates(keep=False,inplace=True)


save_path=r"D:\SOFTWARE\db\downloads\final_df_all_fine_grained from 2000-2022.xlsx"
writer = pd.ExcelWriter(save_path, engine='xlsxwriter',options={'strings_to_urls': False})
final_df_all_fine_grained.to_excel(writer)
writer.close()