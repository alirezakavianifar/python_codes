from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import time
import glob
import os.path
import xlwings as xw
import pyodbc
import gozareshat
import pandas as pd
from datetime import date
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from pathlib import Path




def download_it():
    WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.ID, 'B31879048549323421')))
    driver.find_element(By.ID, 'B31879048549323421').click()
    try:
        element = driver.find_element(By.XPATH, '/html/body/span/span[1]')
        while (element.is_displayed()):
            print("Waiting for the file to be downloaded")
        driver.back()    
    except:        
        driver.back()



# filelist = glob.glob(os.path.join('J:\ezhar-temp', "*.xlsx"))
# for f in filelist:
#     os.remove(f)

gozareshat.loginSanim(r'J:\ezhar-temp')
driver=gozareshat.driver
driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/span/span').click()
time.sleep(1)
driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/button').click()
#try:
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/div/div/div[2]/ul/li[1]/div/span[1]/a')))
driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/div/div/div[2]/ul/li[1]/div/span[1]/a').click()
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div[2]/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a')))
driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div[2]/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a').click()
time.sleep(2)
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a')))
driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a').click()
time.sleep(2)

if Path(r'J:\ezhar-temp\Excel.xlsx').is_file()==False:
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a').click()
    download_it()

if Path(r'J:\ezhar-temp\Excel(1).xlsx').is_file()==False:
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[5]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[5]/a').click()
    download_it()
    
if Path(r'J:\ezhar-temp\Excel(2).xlsx').is_file()==False:
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[6]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[6]/a').click()
    download_it()

if Path(r'J:\ezhar-temp\Excel(3).xlsx').is_file()==False:
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[7]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[7]/a').click()
    download_it()
    
if Path(r'J:\ezhar-temp\Excel(4).xlsx').is_file()==False:
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[8]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[8]/a').click()
    download_it()
    
if Path(r'J:\ezhar-temp\Excel(5).xlsx').is_file()==False:    
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[9]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[9]/a').click()
    download_it()

if Path(r'J:\ezhar-temp\Excel(6).xlsx').is_file()==False:    
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[10]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[10]/a').click()
    download_it()

if Path(r'J:\ezhar-temp\Excel(7).xlsx').is_file()==False:    
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[11]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[11]/a').click()
    download_it()

if Path(r'J:\ezhar-temp\Excel(8).xlsx').is_file()==False:    
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[12]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[12]/a').click()
    download_it()

if Path(r'J:\ezhar-temp\Excel(9).xlsx').is_file()==False:
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[13]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[13]/a').click()
    download_it()

if Path(r'J:\ezhar-temp\Excel(10).xlsx').is_file()==False:    
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[14]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[14]/a').click()
    download_it()
    
if Path(r'J:\ezhar-temp\Excel(11).xlsx').is_file()==False:    
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[15]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[15]/a').click()
    download_it()

if Path(r'J:\ezhar-temp\Excel(12).xlsx').is_file()==False:    
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[16]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[16]/a').click()
    download_it()
    
if Path(r'J:\ezhar-temp\Excel(13).xlsx').is_file()==False:    
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[17]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[17]/a').click()
    download_it()
    
if Path(r'J:\ezhar-temp\Excel(14).xlsx').is_file()==False:    
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[18]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[18]/a').click()
    download_it()

if Path(r'J:\ezhar-temp\Excel(15).xlsx').is_file()==False:    
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[19]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[19]/a').click()
    download_it()
    
if Path(r'J:\ezhar-temp\Excel(16).xlsx').is_file()==False:    
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[20]/a')))
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[20]/a').click()
    download_it()
################################################################################################################################    
    