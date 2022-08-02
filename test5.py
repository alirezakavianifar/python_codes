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


filelist = glob.glob(os.path.join('J:\ezhar-temp', "*.xlsx"))
for f in filelist:
    os.remove(f)

gozareshat.loginSanim(r'J:\ezhar-temp')
driver=gozareshat.driver
driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/span/span').click()
time.sleep(1)
driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/button').click()
#try:
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/div/div/div[2]/ul/li[1]/div/span[1]/a')))
driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/div/div/div[2]/ul/li[1]/div/span[1]/a').click()
##############################################################################################################################
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div[2]/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a')))
driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div[2]/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a').click()
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a')))
driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[4]/a').click()
WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.ID, 'B31879048549323421')))
driver.find_element(By.ID, 'B31879048549323421').click()
time.sleep(2)
element = driver.find_element(By.XPATH, '/html/body/span/span[1]')
while (element.is_displayed()):
    print("Waiting for the file to be downloaded")
time.sleep(2)
driver.back()
print("done")

except TimeoutException:
    print("Loading took too much time!")
    

    time.sleep(2)
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
excl_merged.to_excel(r'J:\final\ghatee-mash-98-{0}.xlsx'.format(date.today()), index=False)
gozareshat.removeExcelFiles(r'J:\ezhar-temp')


