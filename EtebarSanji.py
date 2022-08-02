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
driver.find_element(By.XPATH, '/html/body/form/center/div[1]/table/tbody/tr[2]/td[1]/a[5]/div').click()
time.sleep(1)
driver.find_element(By.XPATH,
                    '/html/body/form/table/tbody/tr[2]/td[2]/span[2]/div[2]/table/tbody/tr/td[2]/div/span').click()
time.sleep(2)
driver.find_element(By.XPATH,
                    '/html/body/form/table/tbody/tr[2]/td[2]/span[2]/div[2]/table/tbody/tr/td[2]/div/div/a[1]').click()
time.sleep(2)
driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[2]/td[2]/span[2]/div/a').click()

while len(glob.glob1('C:\Temp\EtebarSanji', '*.xls')) == 0:
    print('still waiting...')
driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[2]/td[1]/div[7]').click()
time.sleep(1)
driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[2]/td[1]/div[8]/a[1]/div').click()
time.sleep(1)
driver.find_element(By.XPATH,
                    '/html/body/form/table/tbody/tr[2]/td[2]/div/table[1]/tbody/tr/td[3]/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/input').click()
time.sleep(120)
#

# خواندن اطلاعات کد اقتصادی و ذخیره آنها در sql

df = pd.concat(pd.read_excel(r"C:\Temp\EtebarSanji\PN721.AllTaxpayers.16.xls", sheet_name=None), ignore_index=True)
print(df)
#
server = 'SQL Server'
database = 'test'
username = 'sa'
password = '14579Ali.'
sql = """\
INSERT INTO [test].[dbo].[tblEtebarSanji]
           ([استان]
           ,[شهرستان]
           ,[شهر]
           ,[آدرس]
           ,[شناسه ملی]
           ,[کد رهگیری]
           ,[نام شرکت/نام واحدصنفی]
           ,[شماره اقتصادی قدیم]
           ,[شماره ثبت]
           ,[مودی اصلی/ مدیرعامل]
           ,[کد پستی]
           ,[وضعیت ثبت نام]
           ,[تاریخ آخرین تغییر وضعیت]
           ,[مرحله]
           ,[شماره پرونده]
           ,[کد حوزه]
           ,[شماره پرونده ارزش افزوده]
           ,[کد حوزه ارزش افزوده]
           ,[نوع مودی]
           ,[بند قانونی]
           ,[تلفن ]
           ,[نام اتحادیه/نوع شرکت]
           ,[نوع مالکیت]
           ,[شرح اولین فعالیت]
           ,[اداره کل مربوطه]
           ,[وضعیت فعالیت]
           ,[شناسه مشارکت]
           ,[درخواست ابلاغ الکترونیک]
           ,[اداره سنیم]
           ,[کدرهگیری ثبت نام قبلی]
           ,[تاریخ پیش ثبت نام]
           ,[تاریخ شروع فعالیت]
           ,[شماره همراه])
     VALUES
          (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
"""
cnxn = pyodbc.connect(
    'DRIVER={SQL Server};SERVER=.' + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
cursor = cnxn.cursor()

cursor.execute("TRUNCATE TABLE [test].[dbo].[tblEtebarSanji]")
for index, row in df.iterrows():
    print(index)
    cursor.execute(sql,row.استان, row.شهرستان,row.شهر,str(row.آدرس),str(row["شناسه ملی"]),str(row["کد رهگیری"]),str(row["نام شرکت/نام واحدصنفی"]),str(row["شماره اقتصادی قدیم"]),
                   str(row["شماره ثبت"]),str(row["مودی اصلی/ مدیرعامل"]),str(row["کد پستی"]),str(row["وضعیت ثبت نام"]),str(row["تاریخ آخرین تغییر وضعیت"]),
                   str(row["مرحله"]),str(row["شماره پرونده"]),str(row["کد حوزه"]),str(row["شماره پرونده ارزش افزوده"]),str(row["کد حوزه ارزش افزوده"]),str(row["نوع مودی"]),
                   str(row["بند قانونی"]),str(row["تلفن "]),str(row["نام شرکت/نام واحدصنفی"]),str(row["نوع مالکیت"]),str(row["شرح اولین فعالیت"]),str(row["اداره کل مربوطه"]),
                   str(row["وضعیت فعالیت"]),str(row["شناسه مشارکت"]),str(row["درخواست ابلاغ الکترونیک"]),str(row["اداره سنیم"]),str(row["کدرهگیری ثبت نام قبلی"]),
                   str(row["تاریخ پیش ثبت نام"]),str(row["تاریخ شروع فعالیت"]),str(row["شماره همراه"]))

cnxn.commit()
cursor.close()
# حذف فایل کد اقتصادی
os.remove(r"C:\Temp\EtebarSanji\PN721.AllTaxpayers.16.xls")
#

#تغییر نام فایل کلی Tax

filesInDir = os.listdir(r"C:\Temp\EtebarSanji")

for i in filesInDir:
    if i[:3] == "Tax":
        old = os.path.join(r'C:\Temp\etebarsanji', i)
        new = os.path.join(r'C:\Temp\etebarsanji', 'TaxTemp.xlsx')
        os.rename(old, new)

# Delete Last row of TaxTemp File
xl = pd.ExcelFile(r"C:\Temp\EtebarSanji\TaxTemp.xlsx")

dfs = xl.parse(xl.sheet_names[0])
dfs = dfs.iloc[:-1]


# Updating the excel sheet with the updated DataFrame

dfs.to_excel(r"C:\Temp\EtebarSanji\Tax.xlsx", sheet_name='Sheet1', index=False)


os.remove(r"C:\Temp\EtebarSanji\TaxTemp.xlsx")

sxl = pd.read_excel(r"C:\Temp\EtebarSanji\Tax.xlsx")
xl['تاریخ'] = str(datetime.date.today()).replace("-", "")
print(xl["تاریخ"])

server = 'SQL Server'
database = 'test'
username = 'sa'
password = '14579Ali.'
sql = """\
INSERT INTO [test].[dbo].[tblEtebarsanjiAmar]
           (
           [شماره حوزه],
           [حوزه معتبر],
           [تعداد کل اسناد در وضعیت 34،44،45],
           [اسناد چاپ شده],
           [اسناد اعتبارسنجی شده],
           [اعتبارسنجی موفق],
           [مغایرت در اعتبارسنجی],
           [در انتظار اعتبارسنجی در وضعیت 44],
           [آماده برای چاپ],
           [چاپ شده و در انتظار اعتبارسنجی],
           [تاریخ]  
           )
     VALUES
          (?,?,?,?,?,?,?,?,?,?,?)
"""
cnxn = pyodbc.connect(
    'DRIVER={SQL Server};SERVER=.' + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
cursor = cnxn.cursor()
for index, row in xl.iterrows():
    cursor.execute(sql,
                   row["شماره حوزه"],
                   row["حوزه معتبر"],
                   row["تعداد کل اسناد در وضعیت 34،44،45"],
                   row["اسناد چاپ شده"],
                   row["اسناد اعتبارسنجی شده"],
                   row["اعتبارسنجی موفق"],
                   row["مغایرت در اعتبارسنجی"],
                   row["در انتظار اعتبارسنجی در وضعیت 44"],
                   row["آماده برای چاپ"],
                   row["چاپ شده و در انتظار اعتبارسنجی"],
                   row["تاریخ"]
                   )

cnxn.commit()
cursor.close()


driver.quit()
#
