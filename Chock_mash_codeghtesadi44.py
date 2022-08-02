
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import time
import glob
import os.path
import xlwings as xw
import pyodbc
import xlrd
import pandas as pd
import datetime
import pandas as pd
import os
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
#

excels = ['*.xls','*.xlsx','*.xls.part']

for item in excels:    
    filelist = glob.glob(os.path.join('D:\Software\mashReports', item))
    for f in filelist:
        os.remove(f)
      
# پاک کردن فایل های اکسل پوشه temp

# تنظیمات وب درایور فایرفاکس
fp = webdriver.FirefoxProfile()
fp.set_preference('browser.download.folderList', 2)
fp.set_preference('browser.download.manager.showWhenStarting', False)
fp.set_preference('browser.download.dir', r'D:\Software\mashReports\\')

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
time.sleep(1)
driver.find_element(By.XPATH, "/html/body/form/center/div[1]/table/tbody/tr[2]/td[1]/a[1]/div").click()
time.sleep(1)


element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/table/tbody/tr[2]/td[1]/a[2]/div")))
element.click()

element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/table/tbody/tr[2]/td[2]/span[2]/div[2]/table/tbody/tr/td[2]/div/span")))
element.click()


element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/table/tbody/tr[2]/td[2]/span[2]/div[2]/table/tbody/tr/td[2]/div/div/a[1]")))
element.click()

element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/table/tbody/tr[2]/td[2]/span[2]/div/a")))
element.click()

time.sleep(5)

while len(glob.glob1('D:\Software\mashReports', '*.xls*')) != 1:
    print('still waiting...')
    
time.sleep(2)

driver.close()   

# Done for scraping code eghtesadi data

 
hozes = '1603,1607,1608,1615'

hozes = hozes.split(',')

for hoze in hozes:
    
    server = 'SQL Server'
    database = 'MASHAGHEL'
    username = 'mash'
    password = '123456'
    sql = """\
    select distinct k.cod_hozeh as 'کدواحد مالياتي',k.k_parvand as 'کلاسه پرونده',k.sal as 'سال عملکرد',cast(m.Family as char(10)) as 'نام خانوادگي',cast(m.Namee as char(10))  as 'نام',cast(m.Father_name as char(10)) as 'نام پدر',t.s_des as 'شرح فعاليت',k.Modi_Seq as 'شماره مودي',m.Eco_No as 'کد اقتصادي',m.Cod_Meli as 'کد ملي',k.sahm as 'درصد سهم',k.bazargan_mashmol as 'مشمول اتاق بازرگاني' ,cast((case k.sherakat_hamsar when 0 then 'ندارد' when 1 then 'دارد' else 'نا معلوم' end) as char(15)) as 'شراکت همسر',cast((case k.end_vaziyat when 0 then 'فعال' when 1 then 'غير فعال' when 2 then 'متوفي' else 'نا معلوم' end) as char(20)) as 'آخرين وضعيت مودي',k.gr_vares_asli as 'گروه وراث اصلي',k.gr_vares_fari as 'گروه وراث فرعي',m.address as 'آدرس',m.cod_post as 'کد پستي' from  KMLINK_inf k left join modi_inf m  on k.modi_seq = m.modi_seq left join ghabln_inf g on k.k_parvand=g.k_parvand and g.sal=k.sal and g.cod_hozeh=k.cod_hozeh  left join tabl_inf t on g.Faliat_sharh=t.s_code and t.g_code=1  where 1=1 and  replace(k.cod_hozeh,' ','') like '%{0}%'  and  replace(m.Eco_No,' ','') like '%%' and  replace(m.Cod_Meli,' ','') like '%%' and  replace(m.Family,' ','') like '%%' and  replace(m.Namee,' ','') like '%%' and  replace(m.Modi_Seq,' ','') like '%%' order by k.cod_hozeh
    """.format(hoze)
    
    r=[]
    try:
            cnxn = pyodbc.connect(
                'DRIVER={SQL Server};SERVER=10.52.0.50\AHWAZ' + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
            
            tableResult = pd.read_sql(sql, cnxn) 
    
    except:
            print("error")
            
            
    tableResult.to_excel(r"D:\Software\mashReports\{0}.xlsx".format(hoze))
    print("done for {0}".format(hoze))        

path = r'D:\Software\mashReports'

excel_files = glob.glob(os.path.join(path, "*.xlsx"))

merge_excels=[]

for f in excel_files:
    
    # read the csv file
    df = pd.read_excel(f)
    merge_excels.append(df)  
   
all_hozes = pd.concat(merge_excels)


save_path=r"D:\Software\mashReports\all_hozes.xlsx"
writer = pd.ExcelWriter(save_path, engine='xlsxwriter',options={'strings_to_urls': False})
all_hozes.to_excel(writer)
writer.close()


df_eghtesadi = pd.read_excel(r'D:\Software\mashReports\PN721.AllTaxpayers.16.xls', sheet_name = None)

df_eghtesadi = pd.concat(df_eghtesadi)

df_eghtesadi["شناسه ملی"]= df_eghtesadi["شناسه ملی"].map(str)

# all_hozes = all_hozes[all_hozes['کد ملي'].notna()]

df_eghtesadi = df_eghtesadi.loc[df_eghtesadi['وضعیت ثبت نام'].isin(['44'])]
df_eghtesadi = df_eghtesadi.sort_values('تاریخ پیش ثبت نام')


save_path=r"D:\Software\mashReports\df_eghtesadi.xlsx"
writer = pd.ExcelWriter(save_path, engine='xlsxwriter',options={'strings_to_urls': False})
df_eghtesadi.to_excel(writer)
writer.close()

c=df_eghtesadi.columns
d = all_hozes.columns

df_new = df_eghtesadi.merge(all_hozes, how="inner", left_on='شناسه ملی', right_on='کد ملي')
