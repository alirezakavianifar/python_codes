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
import shutil
from helpers_v2 import make_dir_if_not_exists
import multiprocessing


DOWNLOADED_FILES = 0   
n_retries = 0

date_0 = '13900101'
date_97_1 = '13970101'
date_97_2 = '13970102'
date_98_1 = '13980101'
date_98_2 = '13980102'
date_99_1 = '13990101'
date_99_2 = '13990102'
date_00_1 = '14000101'
date_00_2 = '14000102'
date_01_1 = '14010101'
date_01_2 = '14010102'
until_date = '14010419'

urls={'1-b':['http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=0&Edare=undefined&reqtl=1&rwndrnd=0.391666473501828'.format(date_0,until_date)],
      
      '2-b':['http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=1&Edare=undefined&reqtl=1&rwndrnd=0.3663317252369561'.format(date_0,date_97_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=1&Edare=undefined&reqtl=1&rwndrnd=0.3663317252369561'.format(date_97_2,date_98_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=1&Edare=undefined&reqtl=1&rwndrnd=0.3663317252369561'.format(date_98_2,date_99_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=1&Edare=undefined&reqtl=1&rwndrnd=0.3663317252369561'.format(date_99_2,date_00_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=1&Edare=undefined&reqtl=1&rwndrnd=0.3663317252369561'.format(date_00_2,date_01_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=1&Edare=undefined&reqtl=1&rwndrnd=0.3663317252369561'.format(date_01_2,until_date),
      ],
      '3-b':['http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=2&Edare=undefined&reqtl=1&rwndrnd=0.19619883107398017'.format(date_0,until_date)],
      '4-b':['http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=3&Edare=undefined&reqtl=1&rwndrnd=0.9409760878581537'.format(date_0,until_date)],
      
      '5-b':['http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=4&Edare=undefined&reqtl=1&rwndrnd=0.5427791035981446'.format(date_0,date_97_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=4&Edare=undefined&reqtl=1&rwndrnd=0.5427791035981446'.format(date_97_2,date_98_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=4&Edare=undefined&reqtl=1&rwndrnd=0.5427791035981446'.format(date_98_2,date_99_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=4&Edare=undefined&reqtl=1&rwndrnd=0.5427791035981446'.format(date_99_2,date_00_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=4&Edare=undefined&reqtl=1&rwndrnd=0.5427791035981446'.format(date_00_2,date_01_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=4&Edare=undefined&reqtl=1&rwndrnd=0.5427791035981446'.format(date_01_2,until_date),
      ],
      
      '6-b':['http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=5&Edare=undefined&reqtl=1&rwndrnd=0.12348184643733573'.format(date_0,until_date)],
      '7-b':['http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=6&Edare=undefined&reqtl=1&rwndrnd=0.9987754619692941'.format(date_0,until_date)],
      '8-b':['http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=8&Edare=undefined&reqtl=1&rwndrnd=0.2656927736407235'.format(date_0,until_date)],
      
      
      '1-a':['http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=0&Edare=undefined&reqtl=3&rwndrnd=0.2053884823939629'.format(date_0,until_date)],
      
      '2-a':['http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=1&Edare=undefined&reqtl=3&rwndrnd=0.5633793857911045'.format(date_0,date_97_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=1&Edare=undefined&reqtl=3&rwndrnd=0.5633793857911045'.format(date_97_2,date_98_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=1&Edare=undefined&reqtl=3&rwndrnd=0.5633793857911045'.format(date_98_2,date_99_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=1&Edare=undefined&reqtl=3&rwndrnd=0.5633793857911045'.format(date_99_2,date_00_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=1&Edare=undefined&reqtl=3&rwndrnd=0.5633793857911045'.format(date_00_2,date_01_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=1&Edare=undefined&reqtl=3&rwndrnd=0.5633793857911045'.format(date_01_2,until_date),
      ],
      
      '3-a':['http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ShowSelectReqCount.aspx?F={0}&T={1}&S=7&Edare=undefined&reqtl=3&rwndrnd=0.7202288185326258'.format(date_0,until_date)],
      
      
      '1-b-r':['http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=1&Edare=undefined&reqtl=1&rwndrnd=0.6731190588613609'.format(date_0,until_date)],
      
      '2-b-r':['http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=2&Edare=undefined&reqtl=1&rwndrnd=0.3354817228826432'.format(date_0,date_97_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=2&Edare=undefined&reqtl=1&rwndrnd=0.3354817228826432'.format(date_97_2,date_98_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=2&Edare=undefined&reqtl=1&rwndrnd=0.3354817228826432'.format(date_98_2,date_99_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=2&Edare=undefined&reqtl=1&rwndrnd=0.3354817228826432'.format(date_99_2,date_00_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=2&Edare=undefined&reqtl=1&rwndrnd=0.3354817228826432'.format(date_00_2,date_01_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=2&Edare=undefined&reqtl=1&rwndrnd=0.3354817228826432'.format(date_01_2,until_date),
      ],
      
      '3-b-r':['http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=3&Edare=undefined&reqtl=1&rwndrnd=0.3154066478901216'.format(date_0,until_date)],
      
      '1-a-r':['http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=7&Edare=undefined&reqtl=3&rwndrnd=0.6972614531535117'.format(date_0,until_date)],
      
      '2-a-r':['http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=8&Edare=undefined&reqtl=3&rwndrnd=0.8946791182948994'.format(date_0,date_97_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=8&Edare=undefined&reqtl=3&rwndrnd=0.8946791182948994'.format(date_97_2,date_98_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=8&Edare=undefined&reqtl=3&rwndrnd=0.8946791182948994'.format(date_98_2,date_99_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=8&Edare=undefined&reqtl=3&rwndrnd=0.8946791182948994'.format(date_99_2,date_00_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=8&Edare=undefined&reqtl=3&rwndrnd=0.8946791182948994'.format(date_00_2,date_01_1),
      'http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=8&Edare=undefined&reqtl=3&rwndrnd=0.8946791182948994'.format(date_01_2,until_date),
      ],
      
      '3-a-r':['http://govahi186.tax.gov.ir/StimolSoftReport/MoreThan15DaysDetails/ShowMoreThan15DaysDetails.aspx?F={0}&T={1}&S=9&Edare=undefined&reqtl=3&rwndrnd=0.33102715206520283'.format(date_0,until_date)]}



def retry(func):
    def try_it(*args, **kwargs):
        global n_retries
        try:
            result = func(*args, **kwargs)
            return result
        
        except Exception as e:
            n_retries += 1
            print(e)
            if n_retries < 3:
                try_it(*args, **kwargs)
                
    return try_it 



    
def save_process(driver):
    global DOWNLOADED_FILES
    
    save = driver.find_element(By.ID, 'StiWebViewer1_SaveLabel')
        
    if (save.is_displayed()):
        actions = ActionChains(driver)
        actions.move_to_element(save).perform()
        hidden_submenu = driver.find_element(By.XPATH, '/html/body/form/div[3]/span/div[1]/div/table/tbody/tr/td[2]/table/tbody/tr/td[3]/div/div[2]/div/table/tbody/tr/td/table[12]/tbody/tr/td[5]')
        actions.move_to_element(hidden_submenu).perform()
        hidden_submenu.click()
            
        element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "StiWebViewer1_StiWebViewer1ExportDataOnly")))
        element.click()
            
        element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "StiWebViewer1_StiWebViewer1ExportObjectFormatting")))
        element.click()
        time.sleep(1)
        element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[3]/span/table[1]/tbody/tr/td/table/tbody/tr[2]/td[2]/table/tbody/tr[6]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]")))
        element.click()
        time.sleep(2)
        DOWNLOADED_FILES += 1

@retry    
def scrape_186(urls, location=None):
    global DOWNLOADED_FILES        

 
    file_lists=[r'D:\Temp\186',r'D:\Temp\186\bank',r'D:\Temp\186\asnaf',r'D:\Temp\186\asnaf\reply',r'D:\Temp\186\bank\reply']
    make_dir_if_not_exists(file_lists)
    # saving_dir=r'D:\Temp\186\{0}'.format(str(location))
    # make_dir_if_not_exists([saving_dir])
    for f in file_lists:
        
        filelist = glob.glob(os.path.join(f, "*.xlsx"))

        for f in filelist:
            os.remove(f)
    # پاک کردن فایل های اکسل پوشه t    # تنظیمات وب درایور فایرفاکس
    fp = webdriver.FirefoxProfile()
    fp.set_preference('browser.download.folderList', 2)
    fp.set_preference('browser.download.manager.showWhenStarting', False)
    fp.set_preference('browser.download.dir', r'D:\Temp\186\\')

    fp.set_preference('browser.helperApps.neverAsk.openFile',
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream')
    fp.set_preference('browser.helperApps.neverAsk.saveToDisk',
                      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream;application/excel')
    fp.set_preference('browser.helperApps.alwaysAsk.force', False)
    fp.set_preference('browser.download.manager.alertOnEXEOpen', False)
    fp.set_preference('browser.download.manager.focusWhenStarting', False)
    fp.set_preference('browser.download.manager.useWindow', False)
    fp.set_preference('browser.download.manager.showAlertOnComplete', False)
    fp.set_preference('browser.download.manager.closeWhenDone', False)

    driver = webdriver.Firefox(fp, executable_path="D:\drivers\geckodriver.exe")
    driver.window_handles
    driver.switch_to.window(driver.window_handles[0])
    raise Exception('error')
#

    # اجرای فرایند اتوماسیون
    driver.get("http://govahi186.tax.gov.ir/Login.aspx")
    driver.implicitly_wait(5)
    
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "MainContent_txtUserName")))
    element.send_keys("1757400389")
    
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "MainContent_txtPassword")))
    element.send_keys("14579Ali.")
    
    # time.sleep(12)
    
    element = driver.find_element(By.ID, "lblUser")
    
    while (element.is_displayed() == False):
        print("waiting for the login to be completed")
    # element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[3]/div[2]/div/div[1]/div[2]/div/input")))
    # element.click()
    element = driver.find_element(By.ID, "lblUser")
    
    if (element.is_displayed()):
        driver.get("http://govahi186.tax.gov.ir/StimolSoftReport/StatusOfReqsInEdarat/ReportSelectReqCount.aspx?type=1")
        if n_retries == 0:
            index = 0
        else:
            index = DOWNLOADED_FILES + 1
        for key, values in urls.items():
            for value in values:
                driver.get(value)
                time.sleep(1)
                save_process(driver)
                while len(glob.glob1(r'D:\Temp\186', '*.xlsx')) == i:
                    print("waiting")
                    
    time.sleep(45)    
    
    
        
    src = os.path.join(r'D:\Temp\186', "گزارشوضعیتدرخواستها.xlsx")
    dst = os.path.join(r'D:\Temp\186', "1.xlsx")
    os.rename(src, dst)

    for i in range(1,7):
          src = os.path.join(r'D:\Temp\186', "گزارشوضعیتدرخواستها({0}).xlsx".format(i))
          dst = os.path.join(r'D:\Temp\186', "{0}.xlsx".format(i+1))
          os.rename(src, dst)
         
    # for i in range(1,9):
    #     src = os.path.join(r'D:\Temp\186', "{0}.xlsx".format(i))
    #     dst = os.path.join(r'D:\Temp\186\bank', "{0}.xlsx".format(i))
    #     shutil.move(src, dst)
        
        
    # for i in range(8,11):
    #      src = os.path.join(r'D:\Temp\186', "گزارشوضعیتدرخواستها({0}).xlsx".format(i))
    #      dst = os.path.join(r'D:\Temp\186', "{0}.xlsx".format(i+1))
    #      os.rename(src, dst)
         
    # for i in range(1,4):
    #     src = os.path.join(r'D:\Temp\186', "{0}.xlsx".format(i+8))
    #     dst = os.path.join(r'D:\Temp\186\asnaf', "{0}.xlsx".format(i))
    #     shutil.move(src, dst)



        
    # src = os.path.join(r'D:\Temp\186', "Report.xlsx")
    # dst = os.path.join(r'D:\Temp\186', "1.xlsx")
    # os.rename(src, dst)
    # for i in range(1,3):
    #      src = os.path.join(r'D:\Temp\186', "Report({0}).xlsx".format(i))
    #      dst = os.path.join(r'D:\Temp\186', "{0}.xlsx".format(i+1))
    #      os.rename(src, dst)
         
    # for i in range(1,4):
    #     src = os.path.join(r'D:\Temp\186', "{0}.xlsx".format(i))
    #     dst = os.path.join(r'D:\Temp\186\bank\reply', "{0}.xlsx".format(i))
    #     shutil.move(src, dst)
        
        
    # for i in range(3,6):
    #      src = os.path.join(r'D:\Temp\186', "Report({0}).xlsx".format(i))
    #      dst = os.path.join(r'D:\Temp\186', "{0}.xlsx".format(i-2))
    #      os.rename(src, dst)
         
    # for i in range(1,4):
    #     src = os.path.join(r'D:\Temp\186', "{0}.xlsx".format(i))
    #     dst = os.path.join(r'D:\Temp\186\asnaf\reply', "{0}.xlsx".format(i))
    #     shutil.move(src, dst)            
        
    # os.replace(r"path/to/current/file.foo", r"path/to/new/destination/for/file.foo")
    driver.close()
                         

if __name__=="__main__":
    scrape_186(urls=urls[:7])

    
   
        