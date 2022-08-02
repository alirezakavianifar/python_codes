import pyodbc

for driver in pyodbc.drivers():
    # print(driver)

    if '.xlsx' in driver:
        myDriver = driver

conn_str = (r'DRIVER={'+myDriver+'};'
            r'DBQ=C:\Temp\2000iris-mash.xlsx;'
            r'ReadOnly=1')
cnxn=pyodbc.connect(conn_str, autocommit= True)

cursor = cnxn.cursor()

for worksheet in cursor.tables():
    # print(worksheet)
    if worksheet[2]=='Sheet1$':
        tablename=worksheet[2]


cursor.execute('SELECT * FROM [{}]'.format(tablename))

for row in cursor:
    print(row)