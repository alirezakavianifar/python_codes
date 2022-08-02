from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import time
import glob
import os.path
import xlwings as xw
import pyodbc
import gozareshat

driver=gozareshat.driver
pathirismash = r'J:\Temp\2000iris-mash.xlsx'
pathirismash96 = r'J:\Temp\2000iris-mash-96.xlsx'
pathirisarzesh = r'J:\Temp\2000iris-arzesh.xlsx'
pathirissherkat = r'J:\Temp\2000iris-sherkat.xlsx'
pathirissherkat96 = r'J:\Temp\2000iris-sherkat-96.xlsx'

irismash = xw.Book(pathirismash)
irismash.save()
irismash.app.quit()

irismash96 = xw.Book(pathirismash96)
irismash96.save()
irismash96.app.quit()

irisarzesh = xw.Book(pathirisarzesh)
irisarzesh.save()
irisarzesh.app.quit()

irissherkat = xw.Book(pathirissherkat)
irissherkat.save()
irissherkat.app.quit()

irissherkat96 = xw.Book(pathirissherkat96)
irissherkat96.save()
irissherkat96.app.quit()

# Run Stored Procedure spMain
server = 'SQL Server'
database = 'test'
username = 'sa'
password = '14579Ali.'
sql = """\
DECLARE	@return_value int;
EXEC	@return_value = [dbo].[spMain];
SELECT	'Return Value' = @return_value;
"""
cnxn = pyodbc.connect(
    'DRIVER={SQL Server};SERVER=.' + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
cursor = cnxn.cursor()

cursor.execute(sql)
cursor.execute('commit')
print('done')
driver.quit()