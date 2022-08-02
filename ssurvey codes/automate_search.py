import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import os
import time
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
driver = webdriver.Firefox(firefox_profile=profile, executable_path=path)
driver.get("https://id.elsevier.com/as/authorization.oauth2?platSite=SC%2Fscopus&ui_locales=en-US&scope=openid+profile+email+els_auth_info+els_analytics_info+urn%3Acom%3Aelsevier%3Aidp%3Apolicy%3Aproduct%3Aindv_identity&response_type=code&redirect_uri=https%3A%2F%2Fwww.scopus.com%2Fauthredirect.uri%3FtxGid%3Da4b003c835f092e8852ae8421b54149f&state=userLogin%7CtxId%3D38592E3A6232942AF46A296D8D975A5C.i-06bcda286c6a4ad1e%3A6&authType=SINGLE_SIGN_IN&prompt=login&client_id=SCOPUS")
driver.find_element_by_id("bdd-email").send_keys("nora.sayed@student.uts.edu.au")
driver.find_element_by_id("bdd-elsPrimaryBtn").click()
driver.find_element_by_id("bdd-password").send_keys("7536148198")

element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "bdd-elsPrimaryBtn")))
element.click()

element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Advanced document search')]")))
element.click()

element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "contentEditLabel")))
element.click()

element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "searchfield")))
element.send_keys("TITLE-ABS-KEY ( self-adaptive )  AND  ( LIMIT-TO ( PUBYEAR ,  2019 ) )  AND  ( LIMIT-TO ( SUBJAREA ,  ""COMP"" ) )  ")

element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[2]/div/div[3]/div/div[2]/div/form/div/div[1]/div/div/div[2]/div/section/div[2]/ul/li[5]/button/span[1]")))
element.click()

time.sleep(3)

items_fount = int(driver.find_element(By.CLASS_NAME,'resultsCount').text)
pages = int(math.ceil(items_fount/200))

search_result_table = driver.find_element(By.ID, "srchResultsList")

checkboxes = search_result_table.find_elements(By.CLASS_NAME,"checkbox-label")

for checkbox in checkboxes:
    checkbox.click()

element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/span[1]/button[2]/span")))
element.click()
time.sleep(2)
element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/div[4]/div/div/div[2]/div[1]/div/ul/li[5]/label")))
element.click()
time.sleep(1)    
element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/div[4]/div/div/div[3]/div/button[2]/span")))
element.click()

time.sleep(3) 

if pages>1:
    
    for i in range(2,pages+2):
        old_url = driver.current_url
        link_location = '/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[2]/div/div/section[1]/div/div[4]/div/div[2]/ul/li[{0}]/a'.format(i)
        element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, link_location)))
        element.click()
        time.sleep(3)
        new_url = driver.current_url
        
        if old_url == new_url:
            continue
        try:
            search_result_table = driver.find_element(By.ID, "srchResultsList")
        
            checkboxes = search_result_table.find_elements(By.CLASS_NAME,"checkbox-label")
    
                
            for checkbox in checkboxes[1:]:
                checkbox.click()
    
            element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/span[1]/button[2]/span")))
            element.click()
            time.sleep(2)
            element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/div[4]/div/div/div[2]/div[1]/div/ul/li[5]/label")))
            element.click()
            time.sleep(1)
            element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/div[4]/div/div/div[3]/div/button[2]/span")))
            element.click()
        
            time.sleep(3) 
            
        except:
            print("error")        
time.sleep(16)
driver.close()