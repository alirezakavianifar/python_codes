from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import time
import glob
import os.path
import xlwings as xw
import pyodbc
#from xlutils.copy import copy
import xlrd
import pandas as pd
import datetime
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

#

filelist = glob.glob(os.path.join('J:\Temp\186', "*.xls"))
for f in filelist:
    os.remove(f)

filelist = glob.glob(os.path.join('J:\Temp\186', "*.xlsx"))
for f in filelist:
    os.remove(f)
# پاک کردن فایل های اکسل پوشه temp

# تنظیمات وب درایور فایرفاکس
fp = webdriver.FirefoxProfile()
fp.set_preference('browser.download.folderList', 2)
fp.set_preference('browser.download.manager.showWhenStarting', False)
fp.set_preference('browser.download.dir', r'J:\Temp\EtebarSanji\\')

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
driver.get("http://govahi186.tax.gov.ir/Login.aspx")
driver.implicitly_wait(5)

element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "MainContent_txtUserName")))
element.send_keys("1757400389")

element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "MainContent_txtPassword")))
element.send_keys("14579Ali.")

time.sleep(12)

element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[3]/div[2]/div/div[1]/div[2]/div/input")))
element.click()

url="http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ReportSelectReqCount.aspx?type=1"
menu = driver.find_element(By.XPATH,"/html/body/form/div[3]/div[3]/div[1]/div/div[1]/ul/li[3]/a")
hidden_submenu =driver.find_element_by_xpath('/html/body/form/div[3]/div[3]/div[1]/div/div[1]/ul/li[3]/ul/li[1]/a')

  
actions = ActionChains(driver)
actions.move_to_element(menu)
# time.sleep(2)
actions.click(hidden_submenu)
actions.perform()











