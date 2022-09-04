from selenium import webdriver
from selenium.webdriver.common.by import By



class Login:
    def __init__(self, pathsave):
        self.pathsave = pathsave

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
        
        self.driver = webdriver.Firefox(fp, executable_path="H:\driver\geckodriver.exe")
        self.driver.window_handles
        self.driver.switch_to.window(self.driver.window_handles[0])
        
    def __call__(self):
        return self.driver
    
    def close(self):
        self.driver.close()
        

def login_sanim(driver):
    driver.get("https://mgmt.tax.gov.ir/ords/f?p=100:101:16540338045165:::::")
    driver.implicitly_wait(20)
    txtUserName = driver.find_element_by_id('P101_USERNAME').send_keys('1971385018')
    txtPassword = driver.find_element_by_id('P101_PASSWORD').send_keys('123456')
    
    driver.find_element(By.ID, 'B1700889564218640').click()
    
    return driver

        


    
    
    

        

            
            
    
