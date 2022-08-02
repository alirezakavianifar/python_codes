from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import time
import glob
import os.path
import xlwings as xw
import pyodbc
from xlutils.copy import copy
import xlrd
import pandas as pd
import datetime
#

filelist = glob.glob(os.path.join('C:\Temp\EtebarSanji', "*.xls"))
for f in filelist:
    os.remove(f)

filelist = glob.glob(os.path.join('C:\Temp\EtebarSanji', "*.xlsx"))
for f in filelist:
    os.remove(f)
# پاک کردن فایل های اکسل پوشه temp

# تنظیمات وب درایور فایرفاکس
fp = webdriver.FirefoxProfile()
fp.set_preference('browser.download.folderList', 2)
fp.set_preference('browser.download.manager.showWhenStarting', False)
fp.set_preference('browser.download.dir', r'C:\Temp\EtebarSanji\\')

fp.set_preference('browser.helperApps.neverAsk.openFile',
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream')
fp.set_preference('browser.helperApps.neverAsk.saveToDisk',
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream')
fp.set_preference('browser.helperApps.alwaysAsk.force', False)
fp.set_preference('browser.download.manager.alertOnEXEOpen', False)
fp.set_preference('browser.download.manager.focusWhenStarting', False)
fp.set_preference('browser.download.manager.useWindow', False)
fp.set_preference('browser.download.manager.showAlertOnComplete', False)
fp.set_preference('browser.download.manager.closeWhenDone', False)

driver = webdriver.Firefox(fp, executable_path="D:\driver\geckodriver.exe")
driver.window_handles
driver.switch_to.window(driver.window_handles[0])
#

# اجرای فرایند اتوماسیون
driver.get("http://management.tax.gov.ir/Public/Login")
driver.implicitly_wait(5)
txtUserName = driver.find_element_by_id('TextBoxUsername').send_keys('1756914443')
txtPassword = driver.find_element_by_id('TextboxPassword').send_keys('1756914443')
txtPassword = driver.find_element_by_name('ButtonLogin').click()
time.sleep(2)
driver.find_element(By.XPATH, '/html/body/form/center/div[1]/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input').send_keys('1756914443')
driver.find_element(By.XPATH, '/html/body/form/center/div[1]/table/tbody/tr[2]/td[2]/table/tbody/tr/td[2]/input').click()
time.sleep(1)
df=pd.read_excel(r'C:\Users\Administrator\Desktop\بررسی\check.xlsx')
df['وضعیت']='no'
for index, row in df.iterrows():
    driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[1]/td[2]/div/div/table/tbody/tr[2]/td[1]/input').send_keys(str(row["کد ملی"]).split('.')[0])
    driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[1]/td[2]/div/div/table/tbody/tr[2]/td[2]/a/span').click()
    time.sleep(1)
    for item in driver.find_elements_by_class_name('header'):
        if item.text=='جستجو بر اساس شماره ملی حقیقی در ثبت نام' or item.text=='جستجو بر اساس شناسه ملی حقوقی در ثبت نام':
            print('ok')
            df.at[index,'وضعیت']='ok'

print(df)
df.to_excel(r'C:\Users\Administrator\Desktop\بررسی\checkMainShahrestan.xlsx')
