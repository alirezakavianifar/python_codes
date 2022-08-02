import requests
import pandas as pd
import os
import glob


a = final_df_all['title'].unique()

titles = final_df_all["title"].str.lower()
# duplicate_titles = final_df_all[titles.isin(titles[titles.duplicated()])]
final_df_all['title']=final_df_all['title'].str.lower()

low=['self-adaptive','self-healing','self-protecting','self-optimizing',
     'self-optimization','adaptive feedback','Dynamically evolving'
     'self-adaptation','self-managing','pattern','decenralize','reconfiguration',
     ' reinforcement','learning-based','adaptation','reference model'
     'autonomic','contexual','runtime','reusability','model-driven adaptation',
     'dynamically adaptive systems','Distributed adaptive','dynamic updating'
     'self-improvement','self-awareness','context-aware','model-based','adaptive monitor',
     'self-learning','Model evolution','evolution',
     'self-manag','monitor','loop','system','software','agent','uncertainty','known unknown','knowledge evolution','dynamic software product lines',
     'unknown','reuse','evolutionary','search','machine learning','tuning','Auto-adjusting','run-time','Autonomic middleware','requirements-driven']

ltw=['self-adaptive','self-healing','self-protecting','self-optimizing',
     'self-optimization','adaptive feedback','Dynamically evolving'
     'self-adaptation','self-managing','pattern','decenralize','reconfiguration',
     ' reinforcement','learning-based','adaptation','reference model'
     'autonomic','contexual','runtime','reusability','model-driven adaptation',
     'dynamically adaptive systems','Distributed adaptive','dynamic updating'
     'self-improvement','self-awareness','context-aware','model-based','adaptive monitor',
     'self-learning','Model evolution','evolution'
     'self-manag','monitor','loop','system','software','agent','uncertainty','known unknown','knowledge evolution','dynamic software product lines',
     'unknown','reuse','evolutionary','search','machine learning','tuning','Auto-adjusting','run-time','Autonomic middleware','requirements-driven']

for item in low:
    s_1 = final_df_all[final_df_all['title'].str.contains(item, na=False)]
    for word in ltw:
        s_2 = s_1[s_1['title'].str.contains(word, na=False)]
        save_path=r"D:\SOFTWARE\db\fine-grained search\{0} - {1} from 2000-2021.xlsx".format(item,word)
        writer = pd.ExcelWriter(save_path, engine='xlsxwriter',options={'strings_to_urls': False})
        s_2.to_excel(writer)
        writer.close()



path = r'D:\SOFTWARE\db\fine-grained search'

excel_files = glob.glob(os.path.join(path, "*.xlsx"))

merge_excels=[]

for f in excel_files:
    
    # read the csv file
    df = pd.read_excel(f)
    merge_excels.append(df)  
   
final_df_all_fine_grained = pd.concat(merge_excels)

final_df_all_fine_grained.drop_duplicates(keep=False,inplace=True)

final_df_all_fine_grained.drop_duplicates(subset ="url",
                     keep = False, inplace = True)

final_df_all_fine_grained = final_df_all_fine_grained.loc[final_df_all_fine_grained['type'].isin(['Journal Articles','Conference and Workshop Papers'])]


save_path=r"D:\SOFTWARE\db\final_df_all_fine_grained from 2000-2021.xlsx"
writer = pd.ExcelWriter(save_path, engine='xlsxwriter',options={'strings_to_urls': False})
final_df_all_fine_grained.to_excel(writer)
writer.close()