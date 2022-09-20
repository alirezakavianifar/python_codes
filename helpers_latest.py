from selenium import webdriver
import time
import glob
import os.path
import pandas as pd
from datetime import datetime
import argparse
import jdatetime
import xlwings as xw
import shutil
import pandas as pd
import openpyxl
import datetime as dt
from datetime import datetime
from functools import wraps
import math

n_retries = 0
year = 0
report_type = 0



log_folder_name = 'C:\ezhar-temp'
log_excel_name = 'excel.xlsx'
log_dir = os.path.join(log_folder_name, log_excel_name)



def retry(func):
    def try_it():
        global n_retries
        try:
            result = func()
            return result
        
        except Exception as e:
            n_retries += 1
            print(e)
            if n_retries < 5:
                try_it()
                
    return try_it 

        
def retry_with_arguments(driver=None):
    def retry(func):
        def try_it(*args, **kwargs):
            global n_retries
            try:
                result = func(*args, **kwargs)
                return result
            
            except Exception as e:
                print('error occured. please pay close attention')
                n_retries += 1
                # print(e)
                if n_retries < 5:
                    driver.close()
                    time.sleep(4)
                    try_it()
                    
        return try_it 
    return retry


def maybe_make_dir(directory):
    for item in directory:
        
        if not os.path.exists(item):
            os.makedirs(item)
            
def read_multiple_excel_sheets(path):
    df = pd.read_excel(path, sheet_name=None)
    list = []

    for key in df:
        list.append(df[key])
        
    df = pd.concat(list)
    
    return df            
    
def remove_excel_files(files=None,pathsave=None):
    for f in files:
        os.remove(f)
        
        
def merge_multiple_excel_sheets(path,dest):
    lst = []  
    
    df = pd.read_excel(path, sheet_name=None)
    
    for key, value in df.items():
        lst.append(value)
        
    df_all = pd.concat(lst)
    df_all.to_excel(dest)   
    
def merge_multiple_excel_files(path, dest, excel_name='merged'):
    
    dest = dest + '/' + str(excel_name) + '.xlsx'

# csv files in the path
    file_list = glob.glob(path + "/*.xlsx")
    
    excel_files = glob.glob(os.path.join(path, "*.xlsx"))
    merge_excels=[]
    
    for f in excel_files:
         df = pd.read_excel(f)
         merge_excels.append(df)
     
     # Merge all dataframes into one        
    final_df_all_fine_grained = pd.concat(merge_excels)
    
    final_df_all_fine_grained.to_excel(dest, index=False)
    
    print('merging was done successfully')
    
@retry   
def input_info():
    
    parser = argparse.ArgumentParser()
    parser.add_argument('--reportTypes', type=str, help=' Report types:\nezhar = 1\nTashkhisSaderShode = 2\nTashkhisEblaghShode = 3\nGhateeSaderShode = 4\nGhateeEblaghShode = 5\nTabsare100 = 6')
    parser.add_argument('--years', type=str, help='types:\n1395 = 1\n1396 = 2\n1397 = 3\n1398 = 4\n1399 = 5\n1400 = 5\n')
    parser.add_argument('--s',  type=str, nargs='?', default='not-s', help='types:\nt = True\nf = False\n')
    parser.add_argument('--d',  type=str, nargs='?', default='not-d', help='types:\nt = True\nf = False\n')
    parser.add_argument('--c',  type=str, nargs='?', default='not-c', help='types:\nt = True\nf = False\n')
    args = parser.parse_args()


    report_types = args.reportTypes.split(',') # ['1','2','3','4']
    years = args.years.split(',')
    dump_to_sql = args.d
    create_reports = args.c
    scrape = args.s
    
    new_reportTypes = []
    new_years = []
    
    for report_type in report_types:
        if (report_type == '1'):
            new_reportTypes.append('ezhar')
        elif (report_type == '2'):
            new_reportTypes.append('tashkhis_sader_shode')
        elif (report_type == '3'):
            new_reportTypes.append('tashkhis_eblagh_shode')
        elif (report_type == '4'):
            new_reportTypes.append('ghatee_sader_shode')
        elif (report_type == '5'):
            new_reportTypes.append('ghatee_eblagh_shode')
        elif (report_type == '6'):
            new_reportTypes.append('tabsare_100')
        
    for year in years:
        if (year == '1'):
            new_years.append(1395)
        elif (year == '2'):
            new_years.append(1396)
        elif (year == '3'):
             new_years.append(1397)
        elif (year == '4'):
             new_years.append(1398)
        elif (year == '5'):
            new_years.append(1399)
        elif (year == '6'):
             new_years.append(1400)
             
    return new_reportTypes,new_years,scrape, dump_to_sql, create_reports


def get_update_date():
    x = jdatetime.date.today()
    if len(str(x.month)) == 1:
        month = '0%s' % str(x.month)
        
    else:
        month = str(x.month)
    if len(str(x.day)) == 1:
        day = '0%s' % str(x.day)
        
    else:
        day = str(x.day)
        
    year = str(x.year)
        
    update_date =year + month + day
    
    return update_date            
            
def time_it(func):
    def wrapper(*args, **kwargs):
        start = time.time()
        result = func(*args, **kwargs)
        end = time.time()
        print('The operation %s was done successfully' % kwargs['connect_type'])
        print(func.__name__ + ' took ' + str((end - start)) + ' seconds')
        return result
    return wrapper            
            


    


def make_dir_if_not_exists(paths):
    for path in paths:
# Check whether the specified path exists or not
        isExist = os.path.exists(path)
        
        if not isExist:
          
          # Create a new directory because it does not exist 
          os.makedirs(path)
          
          

def check_if_up_to_date(func):
    @wraps(func)
    def try_it(*args, **kwargs):
        
        current_date = int(get_update_date())
        func_name = func.__name__  
        if func_name == 'is_updated_to_save':
            type_of = 'save_excel'
        elif func_name == 'is_updated_to_download':
            type_of = 'download_excel'
        else: 
            type_of = 'save_excel'
            
        
        df = pd.read_excel(log_dir)
        check_date = df['date'].where((df['file_name'] == args[0]) & (df['type_of'] == type_of)).max()  
        time.sleep(1)
        

        if not math.isnan(check_date):

                if (int(check_date) == int(current_date) and func_name != 'save_excel'):
                    print('The %s have already been logged' % args[0])
                    result = func(*args, **kwargs)       
                    return result
                
                elif (int(check_date) != int(current_date) and func_name == 'save_excel'):
                    print('opening excel')
                    result = func(*args, **kwargs)       
                    return result
                else:
                    return False
        
            # return print('The file %s has already been updated \n' % args[0]) 
                     
    return try_it  
  
def log_it(func):
    
    @wraps(func)
    def try_it(*args, **kwargs):
        print('log_it initialized')
        d1 = datetime.now()
        type_of = func.__name__
        
        
        if (func.__name__ == 'save_excel'):
            
            print('opening %s for saving' %args[0])
        result = func(*args, **kwargs)
        
        c_date = get_update_date()
        
   
        if type_of == 'download_excel':
            
            df_1 = pd.DataFrame([[result, c_date, type_of]], columns=['file_name', 'date', 'type_of'])
            
        else:
            
            df_1 = pd.DataFrame([[args[0], c_date, type_of]], columns=['file_name', 'date', 'type_of'])
        
        # create excel file for logging if it does not already exist
        if not os.path.exists(log_dir):
            
            df_1.to_excel(log_dir)
            
        else:
            
            df_2 = pd.read_excel(log_dir, index_col=0)
               
            df_3 = pd.concat([df_1,df_2])
               
            remove_excel_files([log_dir])
            
            df_3.to_excel(log_dir)
                
        d2 = datetime.now()
        d3 = (d2 - d1).total_seconds() / 60
           
        print('it took %s minutes for the %s to be saved' % ("%.2f" % d3 , args[1]))
        print('***********************************************************************\n')

        return result
                 
    return try_it

    


@check_if_up_to_date
def is_updated_to_save(path):
    return True

@check_if_up_to_date
@log_it
def save_excel(excel_file):
    irismash = xw.Book(excel_file)
    irismash.save()
    irismash.app.quit()
    time.sleep(8)
    
         
          
def open_and_save_excel_files(path, report_type=None, year=None, merge=False):
    
    excel_files = glob.glob(os.path.join(path, "*.xlsx"))

    for f in excel_files:
        save_excel(f)

def init_driver(pathsave):
    fp = webdriver.FirefoxProfile()
    fp.set_preference('browser.download.folderList', 2)
    fp.set_preference('browser.download.manager.showWhenStarting', False)
    fp.set_preference('browser.download.dir', pathsave)

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

    driver = webdriver.Firefox(fp, executable_path="H:\driver\geckodriver.exe")
    driver.window_handles
    driver.switch_to.window(driver.window_handles[0])
    
    return driver


def pivot_table():
    path = r'D:\projects\data\iris\test.xlsx'
    df = pd.read_excel(path)

    df = df.loc[df['منبع مالیاتی']]

    sources = list(df['منبع مالیاتی'].unique())

    cols = df.columns

    df = df[df['منبع مالیاتی'].isin(sources)]


    p_df_1 = pd.pivot_table(df, values='کد اداره',
                            index= 'نام اداره', columns=['منبع مالیاتی', 'سال عملکرد'], aggfunc=len, 
                            fill_value=0, margins=True, margins_name='جمع کلی')
