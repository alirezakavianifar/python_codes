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


filelist = glob.glob(os.path.join('J:\ezhar-temp', "*.xlsx"))
for f in filelist:
    os.remove(f)

gozareshat.loginSanim(r'J:\ezhar-temp')
driver=gozareshat.driver
driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/span/span').click()
time.sleep(1)
driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/button').click()
driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/div/div/div[2]/ul/li[1]/div/span[1]/a').click()
time.sleep(2)
driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div[2]/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[1]').click()
time.sleep(2)
#table = driver.find_element(By.TAG_NAME, 'table')

j = 0
for i in range(2,47):
    xpath= "/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[{0}]/td[8]/a".format(i)
    driver.find_element(By.XPATH,xpath).click()
    time.sleep(4)
    driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[1]/div[1]/div[3]/div/button').click()
    driver.find_element(By.XPATH, '/html/body/div[3]/div/ul/li[8]/div/span[1]/button').click()
    time.sleep(2)
    driver.find_element(By.XPATH, '/html/body/div[6]/div[2]/div[2]/div/div/div/div[2]/label/span').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '/html/body/div[6]/div[2]/ul/li/span[1]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '/html/body/div[6]/div[3]/div/button[2]').click()
    time.sleep(1)
    while len(glob.glob1('J:\ezhar-temp', '*.xlsx')) == j:
        print('still waiting...')
    j+=1
    driver.back()
    time.sleep(2)
j = 0    
# specifying the path to csv files
path = "J:\ezhar-temp"

# csv files in the path
file_list = glob.glob(path + "/*.xlsx")

# list of excel files we want to merge.
# pd.read_excel(file_path) reads the excel
# data into pandas dataframe.
excl_list = []

for file in file_list:
	excl_list.append(pd.read_excel(file))

# create a new dataframe to store the
# merged excel file.
excl_merged = pd.DataFrame()

for excl_file in excl_list:
	
	# appends the data into the excl_merged
	# dataframe.
	excl_merged = excl_merged.append(
	excl_file, ignore_index=True)

# exports the dataframe into excel file with
# specified name.
excl_merged.to_excel(r'J:\final\tashkhis-98-{0}.xlsx'.format(date.today()), index=False)
gozareshat.removeExcelFiles(r'J:\ezhar-temp')


