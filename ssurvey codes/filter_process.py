import pandas as pd
import os
import time
import glob
import numpy as np


s_source_fltered_based_on_presence_of_selected_words = r'D:\SOFTWARE\db\downloads\based on words\s_source_fltered_based_on_presence_of_selected_words.xlsx'


df_new = pd.read_excel(r"C:\Users\Msi\Desktop\Articles based on Thesess.xlsx")
df_new['Title'] = df_new['Title'].str.lower()
s_source['Title'] = s_source['Title'].str.lower()

list=df_new['Title'].tolist()
s_2 = s_source[s_source['Title'].str.contains('|'.join(list))]

s_source_new = pd.concat([s_source, s_2]).drop_duplicates(keep=False)

writer = pd.ExcelWriter(s_path, engine='xlsxwriter',options={'strings_to_urls': False})
s_source_new.to_excel(writer)
writer.close()


s_source_filtered = s_source_new.loc[(s_source_new['Year'] == 2017) | (s_source_new['Year'] == 2018) | (s_source_new['Year'] == 2019) | 
                                     (s_source_new['Year'] == 2020) | (s_source_new['Year'] == 2021) | (s_source_new['Year'] == 2022) ][['Title','Abstract','Year','Source title']]
s_source_filtered.to_excel(r"D:\SOFTWARE\db\downloads\based on words\s_source_filtered.xlsx")


#  Function Check if values in the dictionary matches any of the predefined words
def if_matches(my_dict):
    d={}
    for word in searchfor:
        if word in my_dict:
            d[word]=my_dict.get(word)
        else:    
            continue
    return d


#  Function for Count number of occurences of each word in the abstract
def count_words(word_list):
    d = {}
    for word in word_list:
        d[word] = d.get(word,0) + 1
    return d

#  General word count function
def word_count(str):
    counts = dict()
    words = str.split()

    for word in words:
        if word in counts:
            counts[word] += 1
        else:
            counts[word] = 1

    return counts


# Remove extra characters fron the abstracts
for char in '.,\n':
    s_source_filtered["Abstract"]=s_source_filtered["Abstract"].transform(func = lambda x: x.replace('  ',' '))
    
# split string into a list of words
s_source_filtered["splitted_abstract"]=s_source_filtered["Abstract"].transform(func = lambda x: x.split())

#  Count number of occurences of each word in the abstract
s_source_filtered["c_splitted_abstract"]=s_source_filtered["splitted_abstract"].transform(count_words)

# Check if values in the dictionary matches any of the predefined words
s_source_filtered["c_splitted_abstract_if_matches"]=s_source_filtered["c_splitted_abstract"].transform(if_matches)
s_source_filtered["c_splitted_abstract_if_matches_true"]=s_source_filtered["c_splitted_abstract_if_matches"].transform(func = lambda x:bool(x))
df_filtered_final = s_source_filtered.loc[s_source_filtered['c_splitted_abstract_if_matches_true'] == True]


s_source_filtered['Abstracts_counts_filtered'].replace({}, "", inplace=True)

df_filtered_final.to_excel(s_source_fltered_based_on_presence_of_selected_words)    


