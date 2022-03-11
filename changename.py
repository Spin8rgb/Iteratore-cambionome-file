import glob
import os
import xlrd
from pathlib import Path

folder_path = '//.../Giorgio/PREVENTIVI/Preventivi 2014 Moreno'
list = glob.glob(folder_path + "/*.xls")
file_list=os.listdir(folder_path)

#while True:
try:
    for i, y in zip(list, file_list):
        file_excel=xlrd.open_workbook(i)
        sheet = file_excel.sheet_by_index(0)
        v=(sheet.cell(10,8).value)
        if v==0.0:
            v="A"
            print(v)
            y=y[0:7]
            print(y)
        else:
            print(v)
            y=y[0:7]
            print(y)

        file_oldname = os.path.join(folder_path, y+".xls")
        file_newname_newfile = os.path.join(folder_path, y+".a0-"+v+".xls")
        os.rename(file_oldname, file_newname_newfile)



except FileNotFoundError:
    print('errore PDF')
except TypeError:
    print('errore altro')
#except xlrd.biffh.XLRDError:
#    print('errore altro')
