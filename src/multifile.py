import os
import pandas as pd
import os

import pandas as pd

# convert csv to excel
# todo add date recognition for files
for files in os.listdir("output"):
    csvIn = pd.read_csv("output/" + files, delimiter=",")
    excelout = pd.ExcelWriter("excel_output/" + files.split(".")[0] + ".xlsx", engine='xlsxwriter')
    csvIn.to_excel(excelout, sheet_name="Original Data", index=False)
    csvIn.to_excel(excelout, sheet_name="3h Average", index=False)
    excelout.save()
# single file
# path = "excel_output/NO2_2019_06.xlsx"
# book = openpyxl.load_workbook(path)
# avg_3h = book['3h AVG']
# print(avg_3h['A1'].value)
