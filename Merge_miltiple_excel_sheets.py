import pandas as pd

df = pd.read_excel(r"C:\Users\1756914443\Downloads\PN721.AllTaxpayers.16 (1).xls", sheet_name=None)


lst = []

for key, value in df.items():
    lst.append(value)
    
    
df_all = pd.concat(lst)

df_all.to_excel(r"C:\Users\1756914443\Downloads\final.xlsx")    
    