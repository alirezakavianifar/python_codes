from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
import os
import glob
import os.path
import xlwings as xw


filelist = glob.glob(os.path.join('C:\Temp', "*.xlsx"))
for f in filelist:
    os.remove(f)

fp = webdriver.FirefoxProfile()
fp.set_preference('browser.download.folderList', 2)
fp.set_preference('browser.download.manager.showWhenStarting', False)
fp.set_preference('browser.download.dir', r'C:\Temp\\')
fp.set_preference('browser.helperApps.neverAsk.openFile',
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
fp.set_preference('browser.helperApps.neverAsk.saveToDisk',
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
fp.set_preference('browser.helperApps.alwaysAsk.force', False)
fp.set_preference('browser.download.manager.alertOnEXEOpen', False)
fp.set_preference('browser.download.manager.focusWhenStarting', False)
fp.set_preference('browser.download.manager.useWindow', False)
fp.set_preference('browser.download.manager.showAlertOnComplete', False)
fp.set_preference('browser.download.manager.closeWhenDone', False)

driver = webdriver.Firefox(fp, executable_path="D:\driver\geckodriver.exe")
driver.window_handles
driver.switch_to.window(driver.window_handles[0])
driver.get("https://mgmt.tax.gov.ir/ords/f?p=100:101:16540338045165:::::")
driver.implicitly_wait(30)
txtUserName = driver.find_element_by_id('P101_USERNAME').send_keys('1970619521')
txtPassword = driver.find_element_by_id('P101_PASSWORD').send_keys('123456')

driver.find_element(By.ID, 'B1700889564218640').click()
driver.find_element(By.ID, '313_menubar_1i').click()
driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/div/div/ul/li[1]/div/span[1]/a').click()
driver.find_element(By.XPATH, '/html/body/form/div[2]/div/div[2]/div/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[8]/td[5]/a').click()