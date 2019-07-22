import os
import time

import openpyxl
import pandas as pd
from openpyxl.utils.cell import get_column_letter

s = time.time()
# converts to excel file
for files in os.listdir("output"):
    csvIn = pd.read_csv("output/" + files, delimiter=",")
    excelout = pd.ExcelWriter("excel_output/" + files.split(".")[0] + ".xlsx", engine='xlsxwriter')
    csvIn.to_excel(excelout, sheet_name="Original Data", index=False)
    # csvIn.to_excel(excelout, sheet_name="3h Average", index=False)
    wb = excelout.book
    ws = wb.add_worksheet('3h Average')
    ws1 = wb.add_worksheet('Regional Hour Max')
    ws1.write(1, 1, "22222")
    excelout.save()

for excelFiles in os.listdir("excel_output"):
    wb = openpyxl.load_workbook("excel_output/" + excelFiles)
    rawdata = wb['Original Data']
    avg_3h = wb['3h Average']
    numberRow = rawdata.max_row
    numberColumn = rawdata.max_column

    # copies row/column from old set
    for r in range(1, numberRow + 1):
        for c in range(1, 3):
            if r is 1:
                for allC in range(1, numberColumn):
                    avg_3h[get_column_letter(allC) + str(r)] = '=(\'Original Data\'!' + get_column_letter(allC) + str(
                        r) + ')'
            if r is 2 or r is 3:
                for allC in range(3, numberColumn + 1):
                    avg_3h[get_column_letter(allC) + str(r)] = ''
            avg_3h[get_column_letter(c) + str(r)] = '=(\'Original Data\'!' + get_column_letter(c) + str(r) + ')'

    for i in range(0, numberRow - 3):
        for o in range(3, numberColumn):
            avg_3h[get_column_letter(o) + str(i + 4)] = '=IF((COUNTA(\'Original Data\'!' + get_column_letter(o) + str(
                i + 2) + ':' + get_column_letter(o) \
                                                        + str(i + 4) + '))>1,ROUND(AVERAGE(ROUND(\'Original Data\'!' + \
                                                        get_column_letter(o) + str(i + 2) + \
                                                        ',0),ROUND(\'Original Data\'!' + get_column_letter(o) + str(
                i + 3) + \
                                                        ',0),ROUND(\'Original Data\'!' + get_column_letter(o) + str(
                i + 4) + ',0)),0),\" \")'

    wb.save("excel_output/" + excelFiles)

e = time.time()
print(e - s)
# "=IF((COUNTA(raw!C2:C4))>1,(AVERAGE(ROUND(raw!C2,0),ROUND(raw!C3,0),ROUND(raw!C4,0))),\" \")"
# https://automatetheboringstuff.com/chapter12/
