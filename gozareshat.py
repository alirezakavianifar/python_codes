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


def loginSanim(pathsave):
    global driver
    fp = webdriver.FirefoxProfile()
    fp.set_preference('browser.download.folderList', 2)
    fp.set_preference('browser.download.manager.showWhenStarting', False)
    fp.set_preference('browser.download.dir', pathsave)
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
    driver.implicitly_wait(20)
    txtUserName = driver.find_element_by_id('P101_USERNAME').send_keys('1970619521')
    txtPassword = driver.find_element_by_id('P101_PASSWORD').send_keys('123456')

    driver.find_element(By.ID, 'B1700889564218640').click()


def connectToSql():
    global cursor
    global cnxn
    server = 'SQL Server'
    database = 'test'
    username = 'sa'
    password = '14579Ali.'
    cnxn = pyodbc.connect(
        'DRIVER={SQL Server};SERVER=.' + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
    cursor = cnxn.cursor()


def removeExcelFiles(pathsave):
    filelist = glob.glob(os.path.join(pathsave, "*.xlsx"))
    for f in filelist:
        os.remove(f)
        
        
import pandas as pd


df = pd.read_excel(r'D:\projects\report\excelfiles\1395\ezhar\Excel.xlsx')



str_1 = ''

insert_c = ''
insert_p = ''
for c in df.columns:
    str_1 += '[%s] NVARCHAR(MAX) NOT NULL,' % c
    insert_c += '[%s],' % c
    insert_p += '?,'
    
insert_p = insert_p.rsplit(',', 1)[0]

insert_c = insert_c.rsplit(',', 1)[0]

    
str_2 = 'create table [dbo].[testDb].[tblGhateeSaderShode] \n ([id] INT IDENTITY(1,1), PRIMARY KEY ([id]),%s)' % str_1

insert_1 = """INSERT INTO [dbo].[tbltesting] \n
(%s) \n
VALUES (%s)
""" % (insert_c, insert_p)
