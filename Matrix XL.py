import os
import openpyxl
from openpyxl.styles import NamedStyle, Border, Side, Font, GradientFill, Alignment, PatternFill, Color
from openpyxl.writer.excel import save_workbook
import pandas as pd
import easygui
import statistics
import datetime
import calendar
import xlsxwriter


filename_poteri=easygui.fileopenbox()
df1 = pd.read_excel(filename_poteri, header=None)
df1.to_excel(filename_poteri + '.xlsx', index=False, header=False)
wb1 = openpyxl.load_workbook(filename_poteri + '.xlsx')
sh1 = wb1.active

# Identifying period (month, year)
dateCell=str(sh1['A5'].value)
monthNum = dateCell[5:7]
year = dateCell[8:10]
month_list={'01':'Январь', '02':'Февраль', '03':'Март', '04':'Апрель', '05':'Май', '06':'Июнь',
            '07':'Июль', '08':'Август', '09':'Сентябрь', '10':'Октябрь', '11':'Ноябрь', '12':'Декабрь'}
monthName=month_list[monthNum]
period=str(monthName+' 20'+year)

bdList = []
nameList = []
# Генератор вложенных списков
# Lists for prices
prData = [[0 for x in range(3)] for y in range(25)]
# data = [i for i in itertools.repeat(1, 5)]  Another way of generating a list

def listCor(x,y,z): # Prices list editing
    prData[x].pop(y)
    prData[x].insert(y,z)

rowP = 9
for i in range(25):
    bd = sh1.cell(row=rowP, column=1).value
    name = sh1.cell(row=rowP, column=2).value
    bdList.append(bd)
    nameList.append(name)
    pr_zak = sh1.cell(row=rowP, column=10).value # Закупка
    pr_roz = sh1.cell(row=rowP, column=11).value # Розница
    nats = sh1.cell(row=rowP, column=12).value # Наценка
    # pot = sh1.cell(row=rowP, column=13).value  # may be do this one separately
    listCor(i,0,pr_zak)
    listCor(i,1,pr_roz)
    listCor(i, 2, nats)
    rowP += 1

bdPrices = dict(zip(bdList, prData))
commList = dict(zip(bdList, nameList))
wb1.save(str(filename_poteri + '.xlsx'))
wb1.close()
os.remove(filename_poteri+".xlsx")

#===============================================================

filename_matrix=easygui.fileopenbox()
df2 = pd.read_excel(filename_matrix, header=None)
df2.to_excel(filename_matrix + '.xlsx', index=False, header=False)
wb2 = openpyxl.load_workbook(filename_matrix + '.xlsx')
sh2 = wb2.active
listLen = sh2.max_row

branchList = ['Б33', 'БД1', 'БД3', 'БД4']
bdDict = {}
for i in bdList:
    for j in branchList:
        bdDict[i]=j

# bdDict = dict(bdList, dict(branchList))

"""
def matrixData(w,x,y,z): # Prices list editing
    bdDict[w][x].insert(y,z)

for i in bdList:
    for j in range(1, listLen):
        if i == sh2.cell(row=j, column=1).value:
            m_b33 = sh2.cell(row=j, column=24).value
            m_bd1 = sh2.cell(row=j, column=38).value
            m_bd3 = sh2.cell(row=j, column=66).value
            m_bd4 = sh2.cell(row=j, column=80).value
            sr_pr_b33 = sh2.cell(row=j, column=27).value
            sr_pr_bd1 = sh2.cell(row=j, column=41).value
            sr_pr_bd3 = sh2.cell(row=j, column=69).value
            sr_pr_bd4 = sh2.cell(row=j, column=83).value
            dolya_b33 = sh2.cell(row=j, column=30).value
            dolya_bd1 = sh2.cell(row=j, column=44).value
            dolya_bd3 = sh2.cell(row=j, column=72).value
            dolya_bd4 = sh2.cell(row=j, column=86).value
            matrixData(i, 'Б33', 0, m_b33)
            matrixData(i, 'Б33', 1, sr_pr_b33)
            matrixData(i, 'Б33', 2, dolya_b33)
            matrixData(i, 'БД1', 0, m_bd1)
            matrixData(i, 'БД1', 1, sr_pr_bd1)
            matrixData(i, 'БД1', 2, dolya_bd1)
            matrixData(i, 'БД3', 0, m_bd3)
            matrixData(i, 'БД3', 1, sr_pr_bd3)
            matrixData(i, 'БД3', 2, dolya_bd3)
            matrixData(i, 'БД4', 0, m_bd4)
            matrixData(i, 'БД4', 1, sr_pr_bd4)
            matrixData(i, 'БД4', 2, dolya_bd4)


# bdDict = {'Б33':[100, 20, 2, 0.45], 'БД1':[90, 15, 1, 0.35], 'БД3':[110, 25, 3, 0.45]}

# df=pd.read_excel(filename_matrix, header=None)
# df.to_excel(filename_matrix+'.xlsx', index=False, header=False)
# filename_base=list(filename_matrix)
#
# wbm1=xlrd.open_workbook(filename_matrix, formatting_info=True)
# sh_m1elete_rows(1, amount=8)
# os.remove(filename_+'.xlsx')


# Identifying period (month, year)
# period=list(sh_m1['A5'].value)
# month01=period[17]
# month02=period[18]
# month_num=month01+month02
# year01=period[20]
# year02=period[21]
# year='20'+year01+year02
# month_list={'01':'Январь', '02':'Февраль', '03':'Март', '04':'Апрель', '05':'Май', '06':'Июнь', '07':'Июль', '08':'Август', '09':'Сентябрь', '10':'Октябрь', '11':'Ноябрь', '12':'Декабрь'}
# month_name=month_list[month_num]
# Продумать имя файла здесь
# filename_matrix2=str(filename_matrix+month_name+' '+year+'.xlsx')
# wbm2=xlsxwriter.Workbook(filename_matrix2)
# wbm2.close()
#
# wbm2 = openpyxl.load_workbook(filename_matrix2)
# sh_m2 = wbm2.worksheets[0]
# #
# sh_m1.delete_rows(1, amount=8)

# ro=1
# for i in range(1, 4):
# sh_m.cell(1,1).EntireRow.Delete()

# wbm2.save(str(filename_matrix2))

dicti = {'Б33':[100, 20, 2, 0.45], 'БД1':[90, 15, 1, 0.35], 'БД3':[110, 25, 3, 0.45]}

"""