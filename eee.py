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
from selenium.common.exceptions import NoSuchElementException




def iterate_through(i):
    try:

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div[2]/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a')))
        driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div[2]/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a').click()
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a')))
        driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a').click()
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[{0}]/a'.format(i))))
        driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[{0}]/a'.format(i)).click()
        WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.ID, 'B31879048549323421')))
        driver.find_element(By.ID, 'B31879048549323421').click()
        element = driver.find_element(By.XPATH, '/html/body/span/span[1]')
        while (element.is_displayed()):
            print("Waiting for the file to be downloaded")
    except:        
    # driver.back()
        i = i + 1
        if (i>20):
            print("done")
        else:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'B15719546717735129')))
            driver.find_element(By.ID, 'B15719546717735129').click()   
            iterate_through(i)
                


filelist = glob.glob(os.path.join('J:\ezhar-temp', "*.xlsx"))
for f in filelist:
    os.remove(f)

gozareshat.loginSanim(r'J:\ezhar-temp')
driver=gozareshat.driver

driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/span/span').click()
time.sleep(1)
driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/button').click()
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 't_MenuNav_1_0i')))
driver.find_element(By.ID, 't_MenuNav_1_0i').click()    
i = 4    
iterate_through(i)
    # WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[{0}]/a'.format(i))))
    # driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[{0}]/a'.format(i)).click()
    # WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.ID, 'B31879048549323421')))
    # driver.find_element(By.ID, 'B31879048549323421').click()
    # element = driver.find_element(By.XPATH, '/html/body/span/span[1]')
    # while (element.is_displayed()):
    #     print("Waiting for the file to be downloaded")
    # driver.back()

    
# element = driver.find_element(By.XPATH, '/html/body/span/span[1]')
# while (element.is_displayed()):
#     print("Waiting for the file to be downloaded")
# time.sleep(2)
# driver.back()
# print("done")




