import pandas as pd

data = pd.read_excel(r'C:\Users\1756914443\Desktop\ezafe.xls')
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)
print(data.columns)

data = data