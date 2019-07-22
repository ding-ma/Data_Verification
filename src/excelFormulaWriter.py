import os
import time

import openpyxl
import pandas as pd
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
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
    excelout.save()

for excelFiles in os.listdir("excel_output"):

    PurpleFill = PatternFill(bgColor="56007a")
    RedFill = PatternFill(bgColor="EE1111")
    YellowFill = PatternFill(bgColor="EECE00")
    GreenFill = PatternFill(bgColor="00CE15")

    wb = openpyxl.load_workbook("excel_output/" + excelFiles)
    rawdata = wb['Original Data']
    avg_3h = wb['3h Average']
    numberRow = rawdata.max_row
    numberColumn = rawdata.max_column

    for sheets in wb.worksheets:
        if excelFiles.startswith("PM25"):
            sheets.conditional_formatting.add('C2:' + get_column_letter(numberColumn) + str(numberRow),
                                              CellIsRule(operator='greaterThanOrEqual', formula=['50'], stopIfTrue=True,
                                                         fill=PurpleFill))

            sheets.conditional_formatting.add('C2:' + get_column_letter(numberColumn) + str(numberRow),
                                              CellIsRule(operator='between', formula=['35', '49.9999'], stopIfTrue=True,
                                                         fill=RedFill))

            sheets.conditional_formatting.add('C2:' + get_column_letter(numberColumn) + str(numberRow),
                                              CellIsRule(operator='between', formula=['30', '34.9999'], stopIfTrue=True,
                                                         fill=YellowFill))

            sheets.conditional_formatting.add('C2:' + get_column_letter(numberColumn) + str(numberRow),
                                              CellIsRule(operator='between', formula=['25', '29.9999'], stopIfTrue=True,
                                                         fill=GreenFill))

        if excelFiles.startswith("O3"):
            sheets.conditional_formatting.add('C2:' + get_column_letter(numberColumn) + str(numberRow),
                                              CellIsRule(operator='greaterThanOrEqual', formula=['100'],
                                                         stopIfTrue=True,
                                                         fill=PurpleFill))

            sheets.conditional_formatting.add('C2:' + get_column_letter(numberColumn) + str(numberRow),
                                              CellIsRule(operator='between', formula=['85', '99.9999'], stopIfTrue=True,
                                                         fill=RedFill))

            sheets.conditional_formatting.add('C2:' + get_column_letter(numberColumn) + str(numberRow),
                                              CellIsRule(operator='between', formula=['72', '84.9999'], stopIfTrue=True,
                                                         fill=YellowFill))

            sheets.conditional_formatting.add('C2:' + get_column_letter(numberColumn) + str(numberRow),
                                              CellIsRule(operator='between', formula=['62', '71.9999'], stopIfTrue=True,
                                                         fill=GreenFill))

            # todo add for NO2

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

    # formulae for 3h avg
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
