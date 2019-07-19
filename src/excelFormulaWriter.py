import openpyxl
from openpyxl.utils.cell import get_column_letter

wb = openpyxl.load_workbook("test.xlsx")
rawdata = wb['raw']
avg_3h = wb['AVG']
numberRow = rawdata.max_row
numberColumn = rawdata.max_column

# copies row/column from old set
for r in range(1, numberRow + 1):
    for c in range(1, 3):
        if r is 1:
            for allC in range(1, numberColumn):
                avg_3h[get_column_letter(allC) + str(r)] = '=(raw!' + get_column_letter(allC) + str(r) + ')'
        avg_3h[get_column_letter(c) + str(r)] = '=(raw!' + get_column_letter(c) + str(r) + ')'

for i in range(0, numberRow - 3):
    for o in range(3, numberColumn):
        avg_3h[get_column_letter(o) + str(i + 4)] = '=IF((COUNTA(raw!' + get_column_letter(o) + str(
            i + 2) + ':' + get_column_letter(o) \
                                                    + str(i + 4) + '))>1,ROUND(AVERAGE(ROUND(raw!' + get_column_letter(
            o) + str(i + 2) + \
                                                    ',0),ROUND(raw!' + get_column_letter(o) + str(i + 3) + \
                                                    ',0),ROUND(raw!' + get_column_letter(o) + str(
            i + 4) + ',0)),0),\" \")'

wb.save("test.xlsx")

# "=IF((COUNTA(raw!C2:C4))>1,(AVERAGE(ROUND(raw!C2,0),ROUND(raw!C3,0),ROUND(raw!C4,0))),\" \")"
# https://automatetheboringstuff.com/chapter12/
