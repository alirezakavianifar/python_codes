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

filelist = glob.glob(os.path.join('J:\Temp', "*.xlsx"))
for f in filelist:
    os.remove(f)

gozareshat.loginSanim(r'J:\Temp')
driver=gozareshat.driver
driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/span/span').click()
x=driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/div/div/ul/li[3]/div[1]/span[1]/button')
y=driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/div/div/ul/li[3]/div[2]/div/ul/li[1]/div/span[1]/a')
# x = driver.find_element(By.ID, '313_menubar_1_2i')
# y = driver.find_element(By.ID, '313_menubar_1_2_0i')
actions = ActionChains(driver)
actions.move_to_element(x).perform()
time.sleep(5)
actions.move_to_element(y).click().perform()
time.sleep(15)
# drpOstan=Select(driver.find_element(By.ID,'P151_PARAMETER_GTO'))
# drpOstan.select_by_visible_text('خوزستان')
driver.find_element(By.XPATH,
                    '/html/body/form/div[2]/div/div[2]/div/div/div/div/div/div[1]/div[1]/div[1]/div/div[2]/span[1]/span[1]/span').click()
driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('خوزستان')
driver.find_element(By.XPATH,
                        '//*[@id="select2-P151_TAX_YEAR-container"]').click()
driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('1397')
driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys(Keys.RETURN)

driver.find_element(By.ID, 'B145063925488180985').click()
time.sleep(12)
driver.find_element(By.ID, 'ANNUAL_TAXTYPE_actions_button').click()
time.sleep(5)
# f=driver.find_element(By.ID, 'ANNUAL_TAXTYPE_actions_menu_14i')
# print(f.is_displayed())
driver.find_element(By.ID, 'ANNUAL_TAXTYPE_actions_menu_14i').click()
time.sleep(2)
driver.find_element(By.ID, 'download_excel_gpv').click()

while len(glob.glob1('J:\Temp', '*.xlsx')) == 0:
    print('still waiting...')

time.sleep(2)
if len(glob.glob1('J:\Temp', '*.xlsx')) == 1:
    print('شرکت ها')
    driver.find_element(By.XPATH, '//*[@id="select2-P151_TAX_TYPE-container"]').click()
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
while len(glob.glob1(r'J:\Temp\186', '*.xlsx.part')) == 0:
    print("waiting")
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

time.sleep(20)
if len(glob.glob1('J:\Temp', '*.xlsx')) == 3:
    #
    print('مشاغل 96')
    driver.find_element(By.XPATH,
                        '//*[@id="select2-P151_TAX_YEAR-container"]').click()
    driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('1396')
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

time.sleep(20)
if len(glob.glob1('J:\Temp', '*.xlsx')) == 4:
    #
    print('شرکت ها 96')
    driver.find_element(By.XPATH,
                        '//*[@id="select2-P151_TAX_TYPE-container"]').click()
    driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('مالیات بر درآمد شرکت ها')
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

# 98

if len(glob.glob1('J:\Temp', '*.xlsx')) == 5:
    #
    print('شرکت ها 98')
    driver.find_element(By.XPATH,
                        '//*[@id="select2-P151_TAX_YEAR-container"]').click()
    driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('1398')
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

time.sleep(20)
if len(glob.glob1('J:\Temp', '*.xlsx')) == 6:
    #
    print('مشاغل 98')
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


for i in os.listdir('J:\Temp'):
    if i == "Excel.xlsx":
        old = os.path.join('J:\Temp', 'Excel.xlsx')
        new = os.path.join('J:\Temp', '2000iris-arzesh.xlsx')
        os.rename(old, new)
    elif i == "Excel(1).xlsx":
        old = os.path.join('J:\Temp', 'Excel(1).xlsx')
        new = os.path.join('J:\Temp', '2000iris-sherkat.xlsx')
        os.rename(old, new)
    elif i == "Excel(2).xlsx":
        old = os.path.join('J:\Temp', 'Excel(2).xlsx')
        new = os.path.join('J:\Temp', '2000iris-mash.xlsx')
        os.rename(old, new)
    elif i == "Excel(3).xlsx":
        old = os.path.join('J:\Temp', 'Excel(3).xlsx')
        new = os.path.join('J:\Temp', '2000iris-mash-96.xlsx')
        os.rename(old, new)
    elif i == "Excel(4).xlsx":
        old = os.path.join('J:\Temp', 'Excel(4).xlsx')
        new = os.path.join('J:\Temp', '2000iris-sherkat-96.xlsx')
        os.rename(old, new)
    elif i == "Excel(5).xlsx":
        old = os.path.join('J:\Temp', 'Excel(5).xlsx')
        new = os.path.join('J:\Temp', '2000iris-sherkat-98.xlsx')
        os.rename(old, new)
    elif i == "Excel(6).xlsx":
        old = os.path.join('J:\Temp', 'Excel(6).xlsx')
        new = os.path.join('J:\Temp', '2000iris-mash-98.xlsx')
        os.rename(old, new)      

pathirismash = r'J:\Temp\2000iris-mash.xlsx'
pathirismash96 = r'J:\Temp\2000iris-mash-96.xlsx'
pathirisarzesh = r'J:\Temp\2000iris-arzesh.xlsx'
pathirissherkat = r'J:\Temp\2000iris-sherkat.xlsx'
pathirissherkat96 = r'J:\Temp\2000iris-sherkat-96.xlsx'
pathirismash98 = r'J:\Temp\2000iris-mash-98.xlsx'
pathirissherkat98 = r'J:\Temp\2000iris-sherkat-98.xlsx'



irismash = xw.Book(pathirismash)
irismash.save()
irismash.app.quit()

irismash96 = xw.Book(pathirismash96)
irismash96.save()
irismash96.app.quit()

irisarzesh = xw.Book(pathirisarzesh)
irisarzesh.save()
irisarzesh.app.quit()

irissherkat = xw.Book(pathirissherkat)
irissherkat.save()
irissherkat.app.quit()

irissherkat96 = xw.Book(pathirissherkat96)
irissherkat96.save()
irissherkat96.app.quit()

irissherkat98 = xw.Book(pathirissherkat98)
irissherkat98.save()
irissherkat98.app.quit()

irismash98 = xw.Book(pathirismash98)
irismash98.save()
irismash98.app.quit()

# Run Stored Procedure spMain
server = 'SQL Server'
database = 'test'
username = 'sa'
password = '14579Ali.'
sql = """\
DECLARE	@return_value int;
EXEC	@return_value = [dbo].[spMain];
SELECT	'Return Value' = @return_value;
"""
cnxn = pyodbc.connect(
    'DRIVER={SQL Server};SERVER=.' + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
cursor = cnxn.cursor()

cursor.execute(sql)
cursor.execute('commit')
print('done')
driver.quit()
#
# else:
#     print(len(os.listdir('J:\Temp')))
#     driver.quit()
# print(os.listdir('J:\Temp'))
# print(glob.glob)

# list_of_files = glob.glob('J:\Temp') # * means all if need specific format then *.csv
#
# latest_file = max(list_of_files, key=os.path.getctime)
# print (latest_file)
