import pandas as pd
import os
import time
import glob
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import math

profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir", r'D:\SOFTWARE\db\downloads')
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/xml;text/html;application/x-research-info-systems;application/json;text/csv;charset=UTF-8")

titles = pd.read_excel(r"C:\Users\Msi\Desktop\titles_list.xlsx")
titles.to_excel(r"C:\Users\Msi\OneDrive\Research\Survey process\db\fine-grained search\based on words\titles_list.xlsx")
titles = titles["Title"].tolist()
search_term=""
for title in titles:
    if title == titles[-1]:
        search_term+='OR TITLE("{0}")'.format(title)
    else:
        search_term+=' OR TITLE("{0}")'.format(title)
  
# search_term = """TITLE-ABS-KEY ( "Adaptive monitoring" )  OR  TITLE-ABS-KEY ( "adaptive monitor" )  OR  TITLE-ABS-KEY ( "adaptive monitors" )  OR  TITLE-ABS-KEY ( "adaptable monitoring" )  OR  TITLE-ABS-KEY ( "adaptable monitor" )  OR  TITLE-ABS-KEY ( "monitor adaptation" )  OR  TITLE-ABS-KEY ( "monitoring adaptation" )  OR  TITLE-ABS-KEY ( "reconfigurable monitor" )  OR  TITLE-ABS-KEY ( "reconfigurable monitoring" )  OR  TITLE-ABS-KEY ( "monitoring reconfiguration" )  OR  TITLE-ABS-KEY ( "dynamic monitor" )  OR  TITLE-ABS-KEY ( "dynamic monitors" )  OR  TITLE-ABS-KEY ( "dynamic monitoring" )  OR  TITLE-ABS-KEY ( "monitoring evolution" )  OR  TITLE-ABS-KEY ( "monitor evolution" )  OR  TITLE-ABS-KEY ( "monitors evolution" )  OR  TITLE-ABS-KEY ( "evolving monitoring" )  OR  TITLE-ABS-KEY ( "customized monitor" )  OR  TITLE-ABS-KEY ( "customized monitors" )  OR  TITLE-ABS-KEY ( "customized monitoring" )  OR  TITLE-ABS-KEY ( "customised monitoring" )  OR  TITLE-ABS-KEY ( "monitoring personalization" )  OR  TITLE-ABS-KEY ( "personalized monitors" )  OR  TITLE-ABS-KEY ( "personalized monitoring" )  OR  TITLE-ABS-KEY ( "monitoring personalization" )  OR  TITLE-ABS-KEY ( "reactive monitoring" )  OR  TITLE-ABS-KEY ( "reactive monitors" )  OR  TITLE-ABS-KEY ( "proactive monitoring" )  OR  TITLE-ABS-KEY ( "Tuning self-adaptation" )  OR  TITLE-ABS-KEY ( "real-time self-adaptive" )  OR  TITLE-ABS-KEY ( "Auto-adjusting self-adaptive" )  OR  TITLE-ABS-KEY ( "contextual requirements' adaptation" )  OR  TITLE-ABS-KEY ( "uncertainty at runtime" )  OR  TITLE-ABS-KEY ( "self-adaptive software at runtime" )  OR  TITLE-ABS-KEY ( "Strengthening adaptation" )  OR  TITLE-ABS-KEY ( "meta-adaptation" )  OR  TITLE-ABS-KEY ( "meta adaptation" )  OR  TITLE-ABS-KEY ( "evolution of adaptation" )  OR  TITLE-ABS-KEY ( "learning-based framework" )  OR  TITLE-ABS-KEY ( "contextual requirements at runtime" )  OR  TITLE-ABS-KEY ( "behaviour self-adaptation" )  OR  TITLE-ABS-KEY ( "evolution of the adaptation logic" )  OR  TITLE-ABS-KEY ( "Self-learning" )  OR  TITLE-ABS-KEY ( "self-* planning" )  OR  TITLE-ABS-KEY ( "knowledge evolution" )  OR  TITLE-ABS-KEY ( "Evolution in dynamic software product lines" )  OR  TITLE-ABS-KEY ( "Dynamically evolving" )  OR  TITLE-ABS-KEY ( "Active formal models" )  OR  TITLE-ABS-KEY ( "Improving context-awareness" )  OR  TITLE-ABS-KEY ( "context relevance" )  OR  TITLE-ABS-KEY ( "planning in adaptive systems" )  OR  TITLE-ABS-KEY ( "feature-oriented self-adaptive" )  OR  TITLE-ABS-KEY ( "dynamic evolution" )  OR  TITLE-ABS-KEY ( "dynamic updating of control loops" )  OR  TITLE-ABS-KEY ( "flexible evolution" )  OR  TITLE-ABS-KEY ( "Managing Uncertainty" ) OR  TITLE-ABS-KEY ( "Managing Uncertainty" )  OR  TITLE-ABS-KEY ( "self-* systems" )  OR  TITLE-ABS-KEY ( "requirements@ runtime" )  OR  TITLE-ABS-KEY ( "autonomic managers" )  OR  TITLE-ABS-KEY ( "Context-aware reconfiguration" )  OR  TITLE-ABS-KEY ( "model-driven adaptation" )  OR  TITLE-ABS-KEY ( "self-tuning self-adaptive software systems" )  OR  TITLE-ABS-KEY ( "requirements-driven adaptation" )  OR  TITLE-ABS-KEY ( "Interaction-driven self-adaptation" )  OR  TITLE-ABS-KEY ( "Autonomic middleware" )  OR  TITLE-ABS-KEY ( "dynamic adaptation" )  OR  TITLE-ABS-KEY ( "Model evolution by run-time parameter adaptation" )  OR  TITLE-ABS-KEY ( "run-time policy reconfiguration" )  OR  TITLE-ABS-KEY ( "dynamic composition of autonomic" )  OR  TITLE-ABS-KEY ( "uncertainty in self-* systems" )  OR  TITLE-ABS-KEY ( "stochastic self-* planners" )  OR  TITLE-ABS-KEY ( "Managing uncertainty in self-adaptive" )  OR  TITLE-ABS-KEY ( "uncertainty in self-adaptive" )  OR  TITLE-ABS-KEY ( "Anticipating Uncertainty" )  OR  TITLE-ABS-KEY ( "autonomous managers" )  OR  TITLE-ABS-KEY ( "governing policies" )  OR  TITLE-ABS-KEY ( "Policy-based systems" )  OR  TITLE-ABS-KEY ( "Extending Context Awareness" )  OR  TITLE-ABS-KEY ( "Dynamically discovering optimal configurations" )  OR  TITLE-ABS-KEY ( "Search-based Adaptation" )  OR  TITLE-ABS-KEY ( "Adaptation Planning" )  OR  TITLE-ABS-KEY ( "Preparing for the unexpected" )  OR  TITLE-ABS-KEY ( "dependable self-learning" )  OR  TITLE-ABS-KEY ( "self-adaptation framework" )  OR  TITLE-ABS-KEY ( "search-based software engineering" )  OR  TITLE-ABS-KEY ( "self-adaptive software with search-based optimization" )  OR  TITLE-ABS-KEY ( "Self-adaptation using stochastic search" )  OR  TITLE-ABS-KEY ( "Experiment-Driven Adaptation" )  OR  TITLE-ABS-KEY ( "Learning to Self-adapt" )  OR  TITLE-ABS-KEY ( "evolving self-adaptive" )  OR  TITLE-ABS-KEY ( "Policy-based Self-Adaptive" )  OR  TITLE-ABS-KEY ( "Understanding uncertainty in self-adaptive" )  OR  TITLE-ABS-KEY ( "intelligent and explainable self-adaptive systems" )  OR  TITLE-ABS-KEY ( "Hybrid Planning Using Learning and Model Checking" )  OR  TITLE-ABS-KEY ( "Instance-based learning" )  OR  TITLE-ABS-KEY ( "Hybrid planning in self-adaptive systems" )  OR  TITLE-ABS-KEY ( "selfadaptive systems using probabilistic model-checking" )  OR  TITLE-ABS-KEY ( "Optimal planning for architecture-based self-adaptation" )  OR  TITLE-ABS-KEY ( "adaptive complex systems" )  OR  TITLE-ABS-KEY ( "feature-oriented self-adaptive" )  OR  TITLE-ABS-KEY ( "run-time parameter adaptation" )  OR  TITLE-ABS-KEY ( "adaptive distributed systems" )  OR  TITLE-ABS-KEY ( "quantitative verification and sensitivity analysis at run time" )  OR  TITLE-ABS-KEY ( "Architecture-based self-adaptation with reusable infrastructure" )  OR  TITLE-ABS-KEY ( "Adaptive software needs continuous verification" )  OR  TITLE-ABS-KEY ( "distributed adaptive real-time systems" )  OR  TITLE-ABS-KEY ( "Adaptive knowledge bases" )  OR  TITLE-ABS-KEY ( "dynamic adaptation behaviour in real-time systems" )  OR  TITLE-ABS-KEY ( "dynamic adaptation behaviour" )  OR  TITLE-ABS-KEY ( "Parameterisation and optimisation patterns for MAPE-K feedback loops" )  AND  ( LIMIT-TO ( SRCTYPE ,  "p" )  OR  LIMIT-TO ( SRCTYPE ,  "j" ) )  AND  ( LIMIT-TO ( SUBJAREA ,  "COMP" )  OR  EXCLUDE ( SUBJAREA ,  "cp OR LIMIT-TO DOCTYPE" )  OR  EXCLUDE ( SUBJAREA ,  "t LIMIT-TO LANGUAGE" ) )   AND  ( LIMIT-TO ( PUBYEAR ,  {0} ) ) """
path = r"C:/Users/Msi/Desktop/geckodriver.exe"
driver = webdriver.Firefox(firefox_profile=profile, executable_path=path)


driver.get("""https://id.elsevier.com/as/authorization.oauth2?platSite=SC%2Fscopus&ui_locales=en-US&scope=openid+profile+email+els_auth_info+els_analytics_info+urn%3Acom%3Aelsevier%3Aidp%3Apolicy%3Aproduct%3Aindv_identity&response_type=code&redirect_uri=https%3A%2F%2Fwww.scopus.com%2Fauthredirect.uri%3FtxGid%3D90827c055d03af308422f06278903f27&state=userLogin%7CtxId%3DB46EE4D1844E7A2E4A5ED52B4CCEE9CA.i-0ec6f9fb368021ee2%3A4&authType=SINGLE_SIGN_IN&prompt=login&client_id=SCOPUS""")
time.sleep(3)
driver.find_element_by_id("bdd-email").send_keys("nora.sayed@student.uts.edu.au")
driver.find_element_by_id("bdd-elsPrimaryBtn").click()
driver.find_element_by_id("bdd-password").send_keys("7536148198")


element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "bdd-elsPrimaryBtn")))
element.click()
    
element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Advanced document search')]")))
element.click()
    
element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "contentEditLabel")))
element.click()
    
element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "searchfield")))
    
element.send_keys(search_term)
    
element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, """/html/body/div[1]/div/div[1]/div[2]/div/div[3]/div/div[2]
                                                                          /div/form/div/div[1]/
                                                                          div/div/div[2]/div/section/div[2]/ul/li[5]/button/span[1]""")))
element.click()


# element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "selectAllCheck")))
# element.click()    
# time.sleep(3)

for title in titles[1:]:
    
    # items_fount = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CLASS_NAME, "resultsCount")))
    # items_fount = int(items_fount.text)
    # pages = int(math.ceil(items_fount/200))
    
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "selectAllCheck")))
    element.click()
    
    time.sleep(1)
    
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[5]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/span[1]/button[2]/span")))
    element.click()
    
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[5]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/div[4]/div/div/div[2]/div[1]/div/ul/li[4]/label")))
    element.click()
    
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[5]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/div[4]/div/div/div[3]/div/button[2]/span")))
    element.click()
    
    
    # element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/div[4]/div/div/div[2]/div[2]/div/div/table/thead/tr/th[3]/span/label")))
    # element.click()
    
    # element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[2]/div/div/section[1]/div/div[1]/div[1]/div/div/span/div[4]/div/div/div[2]/div[1]/div/ul/li[5]/label")))
    # element.click()
    
    # time.sleep(1)
    
    
    time.sleep(3) 
    
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/nav/ul/li[1]/button/span[2]")))
    element.click()
    
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "searchfield")))
    element.clear()
    element.send_keys('TITLE-ABS-KEY ("{0}")'.format(title))
    
    element.send_keys(Keys.RETURN)
    time.sleep(4)
           
time.sleep(25)

driver.close()



# Merge Excel files

path = r'D:\SOFTWARE\db\downloads'

excel_files = glob.glob(os.path.join(path, "*.csv"))

merge_excels=[]

for f in excel_files:
    
    # read the csv file
    df = pd.read_csv(f)
    merge_excels.append(df)  
   
final_df_all_fine_grained = pd.concat(merge_excels)

final_df_all_fine_grained.drop_duplicates(keep=False,inplace=True)


save_path=r"D:\SOFTWARE\db\downloads\final_df_all_fine_grained from 2000-2022.xlsx"
writer = pd.ExcelWriter(save_path, engine='xlsxwriter',options={'strings_to_urls': False})
final_df_all_fine_grained.to_excel(writer)
writer.close()