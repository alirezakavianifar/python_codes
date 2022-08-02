import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import time
import os.path
import xlrd
import pandas as pd
import datetime
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException


n_retries = 0
n_downloads = 0

def make_dir_if_not_exists(paths):
    for path in paths:
# Check whether the specified path exists or not
        isExist = os.path.exists(path)
        
        if not isExist:
          
          # Create a new directory because it does not exist 
          os.makedirs(path)
          
def retry(func):
    def try_it(Cls):
        global n_retries
        try:
            result = func(Cls)
            return result
        
        except Exception as e:
            n_retries += 1
            print(e)
            if n_retries < 3:
                print('trying again')
                Cls.driver.close()
                time.sleep(3)
                x = Scrape()
                x.scrape_sanim()
                
    return try_it  
          
def init_driver():
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
    
    return driver


class Scrape:
    
    def __init__(self):
        print('init called')
        self.driver = init_driver()
    @retry   
    def scrape_sanim(self):
        pass
        raise Exception()
        
x = Scrape()
x.scrape_sanim()
