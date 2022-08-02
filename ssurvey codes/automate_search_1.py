import requests
import pandas as pd
import os
import glob


low=['Dynamically evolving',
      'self-adaptation','self-managing','adaptation',
      'autonomic','model-driven adaptation','dynamically adaptive systems','Distributed adaptive',
      'self-improvement','context-aware',
      'self-learning','Model evolution','requirements@runtime','models@runtime','models@run.time',
      'uncertainty','known unknown','knowledge evolution','dynamic software product lines',
      'machine learning','Auto-adjusting','Autonomic middleware','requirements-driven']

# low=['Dynamically evolving']

years=['2000','2001','2002','2003','2004','2005','2006','2007','2008','2009','2010',
       '2011','2012','2013','2014','2015','2016','2017','2018','2019','2020','2021']

dfs=[]

for item in low:
    try:
    
        for year in years:

            request_str='https://dblp.org/search/publ/api?q={0}%20year%3A{1}%3A&h=1000&format=json'.format(item,year)
            response = requests.get(request_str)
            if response.status_code == 200:
                data = response.json()
                if bool(data.items()):
                    
                    for (key, value) in data.items():
                        level_1=value
                    
                    for (key, value) in level_1.items():
                        level_2=value
                        
                    for (key, value) in level_2.items():
                        level_3=value    
                    
                    infos=[]
                    
                    if isinstance(level_3, str):
                            continue
                    else:
                        for itm in level_3:
                            for (key,value) in itm.items():
                                if key == 'info':
                                    infos.append(value)
                                
                    df = pd.DataFrame(infos)
                    
                    dfs.append(df)
                    # print("done for {0} year {1}".format_map(item,year))
                else:
                    print(item,year)
    except:
        print("error happened")
        continue
            
    final_df = pd.concat(dfs)
    save_path=r"D:\SOFTWARE\db\{0} keyword from 2000-2021.xlsx".format(item)
    final_df.to_excel(save_path)             



# Combining individual excel files

path = r'D:\SOFTWARE\db'

excel_files = glob.glob(os.path.join(path, "*.xlsx"))

merge_excels=[]

for f in excel_files:
    
    # read the csv file
    df = pd.read_excel(f)
    merge_excels.append(df)  
   
final_df_all = pd.concat(merge_excels)

final_df_all.drop_duplicates(keep=False,inplace=True)

final_df_all.drop_duplicates(subset ="url",
                     keep = False, inplace = True)

# final_df_all = final_df_all.loc[final_df_all['type'].isin(['Journal Articles','Conference and Workshop Papers'])]


save_path=r"D:\SOFTWARE\db\final_df_all from 2000-2021.xlsx"
writer = pd.ExcelWriter(save_path, engine='xlsxwriter',options={'strings_to_urls': False})
final_df_all.to_excel(writer)
writer.close()
