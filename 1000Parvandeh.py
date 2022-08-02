from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import time
import glob
import os.path
import xlwings as xw
import pyodbc


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
driver.implicitly_wait(10)
txtUserName = driver.find_element_by_id('P101_USERNAME').send_keys('1970619521')
txtPassword = driver.find_element_by_id('P101_PASSWORD').send_keys('123456')

driver.find_element(By.ID, 'B1700889564218640').click()
driver.find_element(By.ID, '313_menubar_1i').click()
x = driver.find_element(By.ID, '313_menubar_1_2i')
y = driver.find_element(By.ID, '313_menubar_1_2_0i')
actions = ActionChains(driver)
actions.move_to_element(x).perform()
time.sleep(2)
actions.move_to_element(y).click().perform()
time.sleep(5)
# drpOstan=Select(driver.find_element(By.ID,'P151_PARAMETER_GTO'))
# drpOstan.select_by_visible_text('خوزستان')
driver.find_element(By.XPATH,
                    '/html/body/form/div[2]/div/div[2]/div/div/div/div/div/div[1]/div[1]/div[1]/div/div[2]/span[1]/span[1]/span').click()
driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('خوزستان')
driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys(Keys.RETURN)

driver.find_element(By.ID, 'B145063925488180985').click()
time.sleep(12)
driver.find_element(By.ID, 'ANNUAL_TAXTYPE_actions_button').click()
time.sleep(1)
# f=driver.find_element(By.ID, 'ANNUAL_TAXTYPE_actions_menu_14i')
# print(f.is_displayed())
driver.find_element(By.ID, 'ANNUAL_TAXTYPE_actions_menu_14i').click()
time.sleep(2)
driver.find_element(By.ID, 'download_excel_gpv').click()

while len(glob.glob1('C:\Temp','*.xlsx')) == 0:
    print('still waiting...')

time.sleep(2)
if len(glob.glob1('C:\Temp','*.xlsx')) == 1:
    print('شرکت ها')
    driver.find_element(By.XPATH,'//*[@id="select2-P151_TAX_TYPE-container"]').click()
    driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('مالیات بر درآمد شرکت ها')
    driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys(Keys.RETURN)
    driver.find_element(By.ID, 'B145063925488180985').click()
    time.sleep(12)
    driver.find_element(By.ID, 'ANNUAL_TAXTYPE_actions_button').click()
    time.sleep(1)
    driver.find_element(By.ID, 'ANNUAL_TAXTYPE_actions_menu_14i').click()
    time.sleep(2)
    driver.find_element(By.ID, 'download_excel_gpv').click()
#
time.sleep(15)
if len(glob.glob1('C:\Temp','*.xlsx')) == 2:
#
    print('مشاغل')
    driver.find_element(By.XPATH,
                        '//*[@id="select2-P151_TAX_TYPE-container"]').click()
    driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('مالیات بر درآمد مشاغل')
    driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys(Keys.RETURN)
#
    driver.find_element(By.ID, 'B145063925488180985').click()
    time.sleep(12)
    driver.find_element(By.ID, 'ANNUAL_TAXTYPE_actions_button').click()
    time.sleep(1)
    # f=driver.find_element(By.ID, 'ANNUAL_TAXTYPE_actions_menu_14i')
    # print(f.is_displayed())
    driver.find_element(By.ID, 'ANNUAL_TAXTYPE_actions_menu_14i').click()
    time.sleep(2)
    driver.find_element(By.ID, 'download_excel_gpv').click()


time.sleep(30)

for i in os.listdir('C:\Temp'):
    if i=="Excel.xlsx":
        old = os.path.join('C:\Temp', 'Excel.xlsx')
        new = os.path.join('C:\Temp', '2000iris-arzesh.xlsx')
        os.rename(old, new)
    elif i=="Excel(1).xlsx":
        old = os.path.join('C:\Temp', 'Excel(1).xlsx')
        new = os.path.join('C:\Temp', '2000iris-sherkat.xlsx')
        os.rename(old, new)
    elif i=="Excel(2).xlsx":
        old = os.path.join('C:\Temp', 'Excel(2).xlsx')
        new = os.path.join('C:\Temp', '2000iris-mash.xlsx')
        os.rename(old, new)

pathirismash = r'C:\Temp\2000iris-mash.xlsx'
pathirisarzesh = r'C:\Temp\2000iris-arzesh.xlsx'
pathirissherkat = r'C:\Temp\2000iris-sherkat.xlsx'
irismash = xw.Book(pathirismash)
irismash.save()
irismash.app.quit()

irisarzesh = xw.Book(pathirisarzesh)
irisarzesh.save()
irisarzesh.app.quit()
irissherkat = xw.Book(pathirissherkat)
irissherkat.save()
irissherkat.app.quit()

#Run Stored Procedure spMain
server = 'SQL Server'
database = 'test'
username = 'sa'
password = '14579Ali.'
sql = """\
DECLARE	@return_value int;
EXEC	@return_value = [dbo].[spMain];
SELECT	'Return Value' = @return_value;
"""
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=.'+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

cursor.execute(sql)
cursor.execute('commit')
driver.quit()
#
# else:
#     print(len(os.listdir('C:\Temp')))
#     driver.quit()
# print(os.listdir('C:\Temp'))
# print(glob.glob)

# list_of_files = glob.glob('C:\Temp') # * means all if need specific format then *.csv
#
# latest_file = max(list_of_files, key=os.path.getctime)
# print (latest_file)
