import glob
import pandas as pd
import xlrd
import openpyxl
import numpy as np

folder_path = '//.../Giorgio/PREVENTIVI/Preventivi 2015 Giorgio'
file_list = glob.glob(folder_path + ("/*.xls" or "/*.xlsx" ))
#print(file_list)
path=[]
cli=[]


for i in file_list:
    file_excel=xlrd.open_workbook(i) #APRE ogni file nella lista
    #scrivi la path del file
    print(i)
    path.append(i)
    #scrivi il CLIENTE
    sheet = file_excel.sheet_by_index(0)
    v=(sheet.cell(10,8).value)
    print(v)
    cli.append(v)


df = pd.DataFrame(list(zip(path,cli)), columns = ['PATH','CLIENTE'])
print(df)
df.to_excel (r'//...Giorgio/PREVENTIVI/Preventivi 2015 Giorgio/AAA.xlsx', index = False, header=True)
