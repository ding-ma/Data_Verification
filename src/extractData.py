import csv
import datetime
import os
import shutil
import time

import openpyxl
import pandas as pd
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.utils.cell import get_column_letter

s = time.time()
path = os.getcwd()

directories = ["O3", "PM25", "NO2"]
shutil.rmtree(path + "/monthfiles")
for i in directories:
    if not os.path.exists(path + "/monthfiles/"):
        os.mkdir(path + "/" + "monthfiles")
    if not os.path.exists(path + "/monthfiles/" + i):
        os.mkdir(path + "/monthfiles/" + i)

NAPS_PB_File = open("NAPS_PB.csv", "r")
reader = csv.reader(NAPS_PB_File)
converter = list(reader)
EC_Code = []
NAPS_ID = []

for x in range(len(converter)):
    line = converter[x]
    EC = line[0]
    NAPS = line[1]
    EC_Code.append(EC)
    NAPS_ID.append(NAPS)

NAPS_PB_dict = dict(zip(NAPS_ID, EC_Code))
PB_NAPS_dict = dict(zip(EC_Code, NAPS_ID))

# unix path : /fs/home/fs1/ords/oth/airq_central/frc002/Data/CAS/Observations/Station/

# CMVQ
# 50126
print("Enter the start date in YYYY/MM/DD followed by enter")
start = input()
print("Enter the end date in YYYY/MM/DD followed by enter ")
end = input()

sdatelist = start.strip().split("/")
edatelist = end.strip().split("/")

startDate = datetime.datetime(int(sdatelist[0]), int(sdatelist[1]), int(sdatelist[2]))
endDate = datetime.datetime(int(edatelist[0]), int(edatelist[1]), int(edatelist[2]))

# startDate = datetime.datetime(2019, 6, 1)
# endDate = datetime.datetime(2019, 6, 30)

delta = endDate - startDate

filelstCSV = ["O3_" + startDate.strftime("%Y") + "_" + startDate.strftime("%m") + ".csv",
              "NO2_" + startDate.strftime("%Y") + "_" + startDate.strftime("%m") + ".csv",
              "PM25_" + startDate.strftime("%Y") + "_" + startDate.strftime("%m") + ".csv"]

filelstExcel = ["O3_" + startDate.strftime("%Y") + "_" + startDate.strftime("%m") + ".xlsx",
                "NO2_" + startDate.strftime("%Y") + "_" + startDate.strftime("%m") + ".xlsx",
                "PM25_" + startDate.strftime("%Y") + "_" + startDate.strftime("%m") + ".xlsx"]
listofDate = []

for i in range(delta.days + 1):
    listofDate.append((startDate + datetime.timedelta(days=i)).strftime("%Y%m%d"))

stationlst = []
datelst = []
hourlst = []
O3lst = []
NO2lst = []
PM25lst = []
stationsLstNO2 = open("stationsNO2.txt", "r").read().strip().split(",")
stationsLstO3 = open("stationsO3.txt", "r").read().strip().split(",")
stationsLstPM25 = open("stationsPM25.txt", "r").read().strip().split(",")


def getPerMonth(lst, polluant):
    for station in lst:
        stationID = PB_NAPS_dict[station]
        if os.path.exists(
                "/fs/home/fs1/ords/oth/airq_central/frc002/Data/CAS/Observations/Station/" + stationID) is True:
            for files in sorted(
                    os.listdir("/fs/home/fs1/ords/oth/airq_central/frc002/Data/CAS/Observations/Station/" + stationID)):
                for days in listofDate:
                    if files.endswith(stationID + "_" + days + ".csv"):
                        f = open(
                            "/fs/home/fs1/ords/oth/airq_central/frc002/Data/CAS/Observations/Station/" + stationID + "/" + files,
                            "r")
                        csvFile = list(csv.reader(f))
                        for z in range(len(csvFile)):
                            row = csvFile[z]
                            if row != ['station', 'Date', 'UTC', 'AQHI', 'O3', 'NO2', 'PM2.5', 'PM10', 'SO2', 'H2S',
                                       'CO',
                                       'NO', 'TRS']:
                                station = row[0]
                                stationlst.append(station)
                                date = row[1]
                                datelst.append(date)
                                hour = row[2]
                                hourlst.append(hour)
                                NO2 = row[5]
                                NO2lst.append(NO2)
                                O3 = row[4]
                                O3lst.append(O3)
                                PM25 = row[6]
                                PM25lst.append(PM25)

        if polluant is 1:
            df2 = pd.DataFrame(data={"Date(DD/MM/YYYY)": datelst, "Concentration": PM25lst, "Hour(UTC)": hourlst})
            df2.to_csv("monthfiles/PM25/" + startDate.strftime("%Y%m") + "_" + stationID + "PM25.csv", sep=",",
                       index=False)

        if polluant is 2:
            df1 = pd.DataFrame(data={"Date(DD/MM/YYYY)": datelst, "Concentration": NO2lst, "Hour(UTC)": hourlst})
            df1.to_csv("monthfiles/NO2/" + startDate.strftime("%Y%m") + "_" + stationID + "NO2.csv", sep=",",
                       index=False)

        if polluant is 3:
            df = pd.DataFrame(data={"Date(DD/MM/YYYY)": datelst, "Concentration": O3lst, "Hour(UTC)": hourlst})
            df.to_csv("monthfiles/O3/" + startDate.strftime("%Y%m") + "_" + stationID + "O3.csv", sep=",", index=False)

        monthtemplate = pd.DataFrame(data={"Date(DD/MM/YYYY)": datelst, "Hour(UTC)": hourlst})
        monthtemplate.to_csv("monthTemplate.csv", sep=",", index=False)
        stationlst.clear()
        datelst.clear()
        hourlst.clear()
        NO2lst.clear()
        O3lst.clear()
        PM25lst.clear()


def generateMonthReport(stat, polluant):
    template = open("monthTemplate.csv", "r")
    csvr = list(csv.reader(template))
    hlst = []
    daylst = []
    for q in range(len(csvr)):
        r = csvr[q]
        d = r[0]
        h = r[1]
        daylst.append(d)
        hlst.append(h)

    cclst = []

    finalFile = pd.DataFrame(data={"Date(DD/MM/YYYY)": daylst[1:], "Hours(UTC)": hlst[1:]})
    for station in stat:
        stationID = PB_NAPS_dict[station]
        for file in os.listdir("monthfiles/" + polluant):
            if file.endswith(stationID + polluant + ".csv"):
                F = open("monthfiles/" + polluant + "/" + file, "r")
                CSVFiles = list(csv.reader(F))

                for t in range(len(CSVFiles)):
                    rows = CSVFiles[t]
                    concentration = rows[0]
                    cclst.append(concentration)
                    # print(rows)

        if len(cclst[1:]) != len(daylst[1:]):
            continue
        finalFile[station] = cclst[1:]
        finalFile.to_csv(
            "output/" + polluant + "_" + startDate.strftime("%Y") + "_" + startDate.strftime("%m") + ".csv", sep=",",
            index=False)
        cclst.clear()


print("Starting Request...")
getPerMonth(stationsLstPM25, 1)
print("25% done")
getPerMonth(stationsLstNO2, 2)
getPerMonth(stationsLstO3, 3)
print("50% done")
generateMonthReport(stationsLstNO2, "NO2")
generateMonthReport(stationsLstO3, "O3")
print("75% done")
generateMonthReport(stationsLstPM25, "PM25")

for files in filelstCSV:
    # for files in os.listdir("output"):
    csvIn = pd.read_csv("output/" + files, delimiter=",")
    excelout = pd.ExcelWriter("excel_output/" + files.split(".")[0] + ".xlsx", engine='xlsxwriter')
    csvIn.to_excel(excelout, sheet_name="Original Data", index=False)
    csvIn.to_excel(excelout, sheet_name="3h Average", index=False)
    excelout.save()

for excelFiles in filelstExcel:
    # for excelFiles in os.listdir("excel_output"):
    wb = openpyxl.load_workbook("excel_output/" + excelFiles)
    rawdata = wb['Original Data']
    avg_3h = wb['3h Average']
    numberRow = rawdata.max_row
    numberColumn = rawdata.max_column

    PurpleFill = PatternFill(bgColor="56007a")
    RedFill = PatternFill(bgColor="EE1111")
    YellowFill = PatternFill(bgColor="EECE00")
    GreenFill = PatternFill(bgColor="00CE15")

    # Fill colors
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

    for i in range(0, numberRow - 3):
        for o in range(3, numberColumn + 1):
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
print("Job done, see --> " + path + "/excel_output")
