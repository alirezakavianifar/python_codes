from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import time
import glob
import os.path
import xlwings as xw
import pyodbc
import pandas as pd
import datetime

driver = ""
cursor = ""
cnxn = ""


def loginSanim():
    global driver
    fp = webdriver.FirefoxProfile()
    fp.set_preference('browser.download.folderList', 2)
    fp.set_preference('browser.download.manager.showWhenStarting', False)
    #fp.set_preference('browser.download.dir')
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
    driver.get("https://scholar.google.com/")
    driver.implicitly_wait(20)


    driver.find_element(By.ID, 'gs_hdr_tsi').send_keys("self-adaptive reuse")
    driver.find_element(By.XPATH, '/html/body/div/div[7]/div[1]/div[2]/form/button/span/span[1]').click()







