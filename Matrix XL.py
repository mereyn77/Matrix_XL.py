import os
import openpyxl
from openpyxl.styles import NamedStyle, Border, Side, Font, GradientFill, Alignment, PatternFill, Color
from openpyxl.writer.excel import save_workbook
import pandas as pd
import easygui
import xlsxwriter
import copy # For copying nested dictionaries
import datetime as DT # For make a list of dates


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
    bd = str(sh1.cell(row=rowP, column=1).value)
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
"""
#===============================================================

filename_matrix=easygui.fileopenbox()
df2 = pd.read_excel(filename_matrix, header=None)
df2.to_excel(filename_matrix + 'M.xlsx', index=False, header=False)
wb2 = openpyxl.load_workbook(filename_matrix + 'M.xlsx')
sh2 = wb2.active
listLen = sh2.max_row

bdList2 = bdList[:]
bdDict = dict.fromkeys(bdList2, 0)

for i in bdList2:
    bdDict[i] = dict()
    bdDict[i]['Б33'] = []
    bdDict[i]['БД1'] = []
    bdDict[i]['БД3'] = []
    bdDict[i]['БД4'] = []

def matrixData(w,x,y,z): # Prices list editing
    bdDict[w][x].insert(y,z)

for i in bdList2:
    for j in range(1, listLen):
        if i == str(sh2.cell(row=j, column=1).value):
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

wb2.close()
os.remove(filename_matrix + 'M.xlsx')

"""
# =========================================================================


filename_ved1=easygui.fileopenbox()
df3 = pd.read_excel(filename_ved1, header=None)
df3.to_excel(filename_ved1 + 'V1.xlsx', index=False, header=False)
wbv1 = openpyxl.load_workbook(filename_ved1 + 'V1.xlsx')
sh_v1 = wbv1.active
listLen_v1 = sh_v1.max_row

# Identifying period (month, year)
dateCell2=str(sh_v1['A5'].value)

startD = int(dateCell2[2:4])
startM = int(dateCell2[5:7])
startY = 2000 + int(dateCell2[8:10])
endD = int(dateCell2[14:16])
endM = int(dateCell2[17:19])
endY = 2000 + int(dateCell2[20:])

startDate = DT.datetime(startY, startM, startD)
endDate = DT.datetime(endY, endM, endD)
res = pd.date_range( # Creating a list of dates of the year's period
    min(startDate, endDate),
    max(startDate, endDate)
).strftime('%d.%m.%y').tolist()

bdList3 = bdList[:]
data_stockDict = dict.fromkeys(bdList3, 0)

rowV = 8
for i in range(8, listLen_v1):
    dateList = []
    stockList = []
    init_stock = sh_v1.cell(row=i, column=5).value
    bd_d = str(sh_v1.cell(row=i, column=1).value)
    if init_stock != None:
        data_stockDict[bd_d][res(0)] = init_stock
    else:
        for k in res:
            if k != bd_d:
                data_stockDict[bd_d][k] = init_stock # Is it correct?
            else:
                pr = sh_v1.cell(row=i, column=6).value
                ras = sh_v1.cell(row=i, column=7).value
                data_stockDict[bd_d][k] = init_stock + pr - ras








# for i in bdList3:
#     dataDict[i] = {}
#     bdDict[i]['Б33'] = []
#     bdDict[i]['БД1'] = []
#     bdDict[i]['БД3'] = []
#     bdDict[i]['БД4'] = []


wbv1.close()
os.remove(filename_ved1 + 'V1.xlsx')

