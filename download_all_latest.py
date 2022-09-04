from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import time
import glob
import sys
import os
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from gozaresh_1 import Login, login_sanim
from helpers import maybe_make_dir, input_info, check_if_up_to_date, remove_excel_files, init_driver, log_it
from dump_to_sql_2 import DumpToSQL
from sql_queries_1 import create_sql_table, insert_into, create_sql_table, table_name


n_retries = 0
first_list=[4,8,9,20,21]
second_list=[]
time_out_1 = 2080
time_out_2 = 2080
timeout_fifteen = 15
excel_file_names = ['Excel.xlsx', 'Excel(1).xlsx', 'Excel(2).xlsx']

download_button_ezhar = '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[1]/div[2]/button[3]'
download_button_rest = '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/font/div/div/div[2]/div[1]/div[2]/button[2]'

menu_nav_1 = '//*[@id="t_MenuNav_1_1i"]'
menu_nav_2 = '/html/body/form/header/div[2]/div/ul/li[2]/div/div/div[2]/ul/li[1]/div/span[1]/a'
year_button_1 = '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/font/div[1]/div/div/div/div[2]/div/div/button'
year_button_2 = '/html/body/div[7]/div[2]/div[1]/button'
year_button_3 = '/html/body/div[7]/div[2]/div[2]/div/div[3]/ul/li'
year_button_4 =  '/html/body/div[3]/div/ul/li[8]/div/span[1]/button'
input_1 = '/html/body/span/span/span[1]/input'
td_1 = '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/font/div/div/div[2]/div[2]/div[5]/div[1]/div/div[3]/table/tbody/tr[2]'
td_2 = '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/font/div[2]/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]'
td_3 = '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div[2]/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]'
year_button_5 = '/html/body/div[6]/div[3]/div/button[2]'
year_button_6 = '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[1]/div[1]/div[3]/div/button'
td_4 = '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[1]/table/tr/th[8]/a'
td_5 = '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div[1]/div/div/div/div[2]/div/span/span[1]/span/span[2]'
td_6 = '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/font/div[1]/div/div/div/div[2]/div/span/span[1]/span/span[1]'
td_ezhar = '/html/body/form/div[2]/div/div[2]/main/div[2]/div/div/div/div/div/div/div[2]/div[2]/div[5]/div[1]/div/div[2]/table/tbody/tr[2]/td[%s]/a'
          
def retry(func):
    def try_it(Cls):
        global n_retries
        try:
            result = func(Cls)
            return result
        
        except Exception as e:
            n_retries += 1
            print(e)
            if n_retries < 5:
                print('trying again')
                Cls.driver.close()
                path = Cls.path
                report_type = Cls.report_type
                year = Cls.year
                time.sleep(3)
                x = Scrape(path, report_type, year)
                x.scrape_sanim()
                
    return try_it  
          



@check_if_up_to_date
def is_updated_to_download(path):
    return True

@check_if_up_to_date
def is_updated_to_save(path):
    return True


class Scrape:
    
    def __init__(self, path, report_type, year):
        self.path = path
        self.report_type = report_type
        self.year = year        
        self.driver = init_driver(pathsave=self.path)
        
    
    @log_it    
    def download_excel(self, type_of_excel=None, no_files_in_path=None):
        print(no_files_in_path)
        i=0
        while len(glob.glob1(self.path, '%s' % excel_file_names[no_files_in_path])) == 0:
           if i%60 == 0:   
               print('waiting %s seconds for the file to be downloaded' % i)
           i+=1
           time.sleep(1)
       
        print('****************%s done*******************************' %type_of_excel)
        
        excel_files = glob.glob1(self.path, '*.xlsx')                   
            
        self.driver.back()
        
        return self.path +'\\' + excel_file_names[no_files_in_path]
        
    @retry   
    def scrape_sanim(self):
        global excel_file_names

                    
                    
        self.driver = login_sanim(self.driver)

        
        if self.report_type == 'ezhar':
            download_button = download_button_ezhar
        else:
            download_button = download_button_rest
        

        
        #انتخاب منوی گزارشات اصلی        
        self.driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/span/span').click()
        time.sleep(1)
        self.driver.find_element(By.XPATH, '/html/body/form/header/div[2]/div/ul/li[2]/button').click() 
        if self.report_type != 'tabsare_100':
        
            if (self.report_type == 'ezhar'):
                td_number = 4
            elif (self.report_type == 'tashkhis_sader_shode'):                
                td_number = 8
            elif (self.report_type == 'tashkhis_eblagh_shode'): 
                td_number = 9
            elif (self.report_type == 'ghatee_sader_shode'):                  
                td_number = 20
            elif (self.report_type == 'ghatee_eblagh_shode'):
                
                td_number = 21
    
            
                                                # انتخاب منوی اول از گزارشات اصلی                
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, menu_nav_2)))
            self.driver.find_element(By.XPATH, menu_nav_2).click()
            
            time.sleep(3)
            
                                                                    # انتخاب سال عملکرد                
            WebDriverWait(self.driver, 8).until(EC.presence_of_element_located((By.XPATH, year_button_1)))
            self.driver.find_element(By.XPATH,year_button_1).click()  
            
            self.driver.find_element(By.XPATH,'/html/body/div[7]/div[2]/div[1]/input').send_keys(self.year)
            
            WebDriverWait(self.driver, 8).until(EC.presence_of_element_located((By.XPATH, year_button_2)))
            self.driver.find_element(By.XPATH,year_button_2).click()  
            
            time.sleep(3)
            
            WebDriverWait(self.driver, 8).until(EC.presence_of_element_located((By.XPATH, year_button_3)))
            self.driver.find_element(By.XPATH,year_button_3).click() 
            
            #################################################################################################################################
            
            time.sleep(3)
            
            WebDriverWait(self.driver, time_out_1).until(EC.presence_of_element_located((By.XPATH, '%s/td[4]/a' % td_2)))
            self.driver.find_element(By.XPATH, '%s/td[%s]/a' % (td_2, td_number)).click()
            
            time.sleep(4)
            print('reached here')
                                                                                    # دریافت اظهارنامه ها و تشخیص های صادر شده                
            exists_in_first_list = first_list.count(td_number)
            
            if (exists_in_first_list):
                
                # print(check_if_up_to_date('%s\%s' % (self.path, excel_file_names[0])))

                if not (is_updated_to_download('%s\%s' % (self.path, excel_file_names[0]))):
                # if(uptodate.count(self.path + '\Excel.xlsx') == 0):
                    print('updating for report_type=%s and year=%s' % (self.report_type, self.year))
                    # WebDriverWait(self.driver, time_out_1).until(EC.presence_of_element_located((By.XPATH, '%s/td[4]/a' % td_1)))
                    # self.driver.find_element(By.XPATH, '%s/td[5]/a' % td_1).click()
                    if (self.report_type != 'ezhar'):
                        WebDriverWait(self.driver, timeout_fifteen).until(EC.presence_of_element_located((By.XPATH, '%s/td[4]/a' % td_1)))
                        self.driver.find_element(By.XPATH, '%s/td[4]/a' % td_1).click() 

                    else:
                        WebDriverWait(self.driver, timeout_fifteen).until(EC.presence_of_element_located((By.XPATH, td_ezhar % 5)))
                        self.driver.find_element(By.XPATH, td_ezhar % 5).click() 
                        
                    print('waiting')
                    time.sleep(4)
                    WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, download_button)))
                    self.driver.find_element(By.XPATH, download_button).click()
                    
                    
                    print('*******************************************************************************************')
                    
                    self.download_excel('hoghoghi', 0)
                    
                if not (is_updated_to_download('%s\%s' % (self.path, excel_file_names[1]))):
                    print('updating for report_type=%s and year=%s' % (self.report_type, self.year))
                    if (self.report_type != 'ezhar'):
                        WebDriverWait(self.driver, timeout_fifteen).until(EC.presence_of_element_located((By.XPATH, '%s/td[3]/a' % td_1)))
                        self.driver.find_element(By.XPATH, '%s/td[3]/a' % td_1).click()
                    
                    else:
                        WebDriverWait(self.driver, timeout_fifteen).until(EC.presence_of_element_located((By.XPATH, td_ezhar % 4)))
                        self.driver.find_element(By.XPATH, td_ezhar % 4).click()                        
                    
                    time.sleep(4)
                    WebDriverWait(self.driver, time_out_1).until(EC.presence_of_element_located((By.XPATH, download_button)))
                    self.driver.find_element(By.XPATH, download_button).click()
                    
                    print('*******************************************************************************************')
                    
                    self.download_excel('haghighi', 1)

                    
                if not (is_updated_to_download(excel_2)):
                    print('updating for report_type=%s and year=%s' % (self.report_type, self.year))
                    if (self.report_type != 'ezhar'):
                        WebDriverWait(self.driver, timeout_fifteen).until(EC.presence_of_element_located((By.XPATH, '%s/td[7]/a' % td_1)))
                        self.driver.find_element(By.XPATH, '%s/td[7]/a' % td_1).click()
                    else:
                        WebDriverWait(self.driver, timeout_fifteen).until(EC.presence_of_element_located((By.XPATH, td_ezhar % 8)))
                        self.driver.find_element(By.XPATH, td_ezhar % 8).click()                        
                    
                    time.sleep(4)
                    WebDriverWait(self.driver, time_out_1).until(EC.presence_of_element_located((By.XPATH, download_button)))
                    self.driver.find_element(By.XPATH, download_button).click()
                    
                    print('*******************************************************************************************')
                    
                    self.download_excel('Arzesh Afzoode', 2)

    
            # if there is only one report and no distinction between haghighi, hoghoghi and arzesh afzoode 
            # else:
            #     time.sleep(4)
            #     WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, download_button)))
            #     self.driver.find_element(By.XPATH, download_button).click()
                
            #     i=0
            #     while len(glob.glob1(self.path, '*.xlsx')) == 0:
            #         if i%60 == 0:
            #             print('waiting %s seconds for the file to be downloaded' % i)
            #         i+=1
            #         time.sleep(1)
                            
            #     self.driver.back()
                
            #     time.sleep(150)
                
        else:
            
                # گزارشات مرتبط با تبصره 100 ادارات             
            self.driver.find_element(By.XPATH, menu_nav_1).click()
            time.sleep(2)
            self.driver.find_element(By.XPATH, '//*[@id="t_MenuNav_1_1_0i"]').click()
            
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, td_5)))
            self.driver.find_element(By.XPATH, td_5).click()
    
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, input_1)))
            self.driver.find_element(By.XPATH, input_1).send_keys(self.year)
            
            time.sleep(2)
            
            
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, input_1)))
            self.driver.find_element(By.XPATH, input_1).send_keys(Keys.ENTER)
            
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, '%s/td[4]/a' % td_3)))
            self.driver.find_element(By.XPATH, '%s/td[4]/a' % td_3).click()
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, td_4)))
    
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, year_button_6)))
            self.driver.find_element(By.XPATH, year_button_6).click()
    
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH,year_button_4)))
            self.driver.find_element(By.XPATH,year_button_4).click()
    
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, year_button_5)))
            self.driver.find_element(By.XPATH, year_button_5).click()
    
            self.download_excel('Tabsare 100', 0)
            
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, '%s/td[5]/a' % td_3)))
            self.driver.find_element(By.XPATH, '%s/td[5]/a' % td_3).click()
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, td_4)))
    
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, year_button_6)))
            self.driver.find_element(By.XPATH, year_button_6).click()
    
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH,year_button_4)))
            self.driver.find_element(By.XPATH,year_button_4).click()
    
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, year_button_5)))
            self.driver.find_element(By.XPATH, year_button_5).click()
            
            
            self.download_excel('Tabsare 100', 1)
            
            time.sleep(5)
            
            self.driver.back()
            
            time.sleep(5)
                        
            self.driver.back()
            
            time.sleep(10)
    
            self.driver.back()
    
            self.driver.find_element(By.XPATH, menu_nav_1).click()
            time.sleep(2)
            self.driver.find_element(By.XPATH, '//*[@id="t_MenuNav_1_1_1i"]').click()
            
            time.sleep(40)
                
            WebDriverWait(self.driver, 50).until(EC.presence_of_element_located((By.XPATH, td_6)))
            self.driver.find_element(By.XPATH, td_6).click()
            
            WebDriverWait(self.driver, 45).until(EC.presence_of_element_located((By.XPATH, input_1)))
            self.driver.find_element(By.XPATH, input_1).send_keys(self.year)
            
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, input_1)))
            self.driver.find_element(By.XPATH, input_1).send_keys(Keys.ENTER)
            
            WebDriverWait(self.driver, time_out_1).until(EC.presence_of_element_located((By.XPATH, '%s/td[3]/a' % td_2)))
            self.driver.find_element(By.XPATH, '%s/td[3]/a' % td_2).click()
            
            WebDriverWait(self.driver, time_out_2).until(EC.presence_of_element_located((By.XPATH, download_button)))
            self.driver.find_element(By.XPATH, download_button).click()

            self.download_excel('final file', 2)               
            return
        
        time.sleep(20)
        self.driver.close()


if __name__=="__main__":
    
    
    
    
    report_types, years, initial_scrape, dump_to_sql, create_reports = input_info()
    print('dump_to_sql = %s and create_reports = %s and scrape = %s' %(dump_to_sql, create_reports, initial_scrape))

    
    for year in years:
        
        for report_type in report_types:
            scrape = initial_scrape 
            path = r'C:\ezhar-temp\%s\%s' % (year, report_type)
            maybe_make_dir([path])

        # Update the reports
            print('updating excel files...................................')
            
            # download excel files
            if (scrape == 's'): 
                
                        
                excel_0 = '%s\%s' % (path, excel_file_names[0])
                excel_1 = '%s\%s' % (path, excel_file_names[1])
                excel_2 = '%s\%s' % (path, excel_file_names[2])
                
                hoghoghi_updated = is_updated_to_download(excel_0)
                haghighi_updated = is_updated_to_download(excel_1)
                arzeshafzoode_updated = is_updated_to_download(excel_2)
                
                if (os.path.exists(excel_0) and not (hoghoghi_updated)):
                    remove_excel_files([excel_0])
                    
                if  (os.path.exists(excel_1) and not (haghighi_updated)):
                    remove_excel_files([excel_1])
            
                if (os.path.exists(excel_2) and not (arzeshafzoode_updated)):
                    remove_excel_files([excel_2])
                
                
                if  ((hoghoghi_updated) and
                    (haghighi_updated) and
                    (arzeshafzoode_updated)):
                
                    print('All excel files related to %s year %s are up to date' % (report_type, year))
                    scrape = 'not-s'
                else:
                    x = Scrape(path=path, report_type=report_type, year=year) 
                    x.scrape_sanim()
                    
            
                    
            # Dump excel files into sql table
            
            if (dump_to_sql == 'd'):
                if (is_updated_to_save(os.path.join(path, excel_file_names[0])) and
                is_updated_to_save(os.path.join(path, excel_file_names[1])) and
                is_updated_to_save(os.path.join(path, excel_file_names[2]))):
                
                    print('All excel files related to %s year %s are saved' % (report_type, year))
                    dump_to_sql = 'not-d'
                    
                else: 
                    table = table_name(report_type, year)
                    sql_delete = """
                    BEGIN TRANSACTION
                            BEGIN TRY 
                        
                                IF Object_ID('%s') IS NOT NULL DROP TABLE %s
                        
                                COMMIT TRANSACTION;
                                
                            END TRY
                            BEGIN CATCH
                                ROLLBACK TRANSACTION;
                            END CATCH
                    """ % (table, table)
                    dump = DumpToSQL(report_type, table, year, sql_delete, path)
                    dump.dump_to_sql()

            
            
            # if(len(excel_files_to_be_removed) > 0):
            #     print('Removing excel files...............................')
            #     remove_excel_files(excel_files_to_be_removed,path)
                # remove_excel_files(excel_files_to_be_removed,saved_folder)

                # print('done removing excel files..........................')

                
                ##########################################################
                # Dump excel data into sql database
                
                                
                
                
            if (create_reports == 'c'):
                dump.create_anbare_reports()
                dump.create_Anbare99Mashaghel_reports()
                dump.create_Anbare99Sherkatha_reports()
                dump.create_hesabrasiArzeshAfzoode_reports()                

        # sys.exit()
