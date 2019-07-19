import csv
import datetime
import os
import shutil

import pandas as pd

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
# print("Enter the start date in YYYY/MM/DD followed by enter")
# start = input()
# print("Enter the end date in YYYY/MM/DD followed by enter ")
# end = input()
#
# sdatelist = start.strip().split("/")
# edatelist = end.strip().split("/")
#
# startDate = datetime.datetime(int(sdatelist[0]),int(sdatelist[1]), int(sdatelist[2]))
# endDate = datetime.datetime(int(edatelist[0]), int(edatelist[1]), int(edatelist[2]))

startDate = datetime.datetime(2019, 6, 1)
endDate = datetime.datetime(2019, 6, 30)
delta = endDate - startDate
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
print("Job done, see --> " + path + "/output")

# if polluant is 1:
#     df2 = pd.DataFrame(data={"Date(DD/MM/YYYY)": datelst, "Concentration": PM25lst, "Hour(UTC)": hourlst})
#     df2.to_excel("monthfiles/PM25/" + startDate.strftime("%Y%m") + "_" + stationID + "PM25.xlsx", sep=",",
#                  index=False)
#
# if polluant is 2:
#     df1 = pd.DataFrame(data={"Date(DD/MM/YYYY)": datelst, "Concentration": NO2lst, "Hour(UTC)": hourlst})
#     df1.to_excel("monthfiles/NO2/" + startDate.strftime("%Y%m") + "_" + stationID + "NO2.xlsx", sep=",",
#                  index=False)
#
# if polluant is 3:
#     df = pd.DataFrame(data={"Date(DD/MM/YYYY)": datelst, "Concentration": O3lst, "Hour(UTC)": hourlst})
#     df.to_excel("monthfiles/O3/" + startDate.strftime("%Y%m") + "_" + stationID + "O3.xlsx", sep=",", index=False)
#
#     monthtemplate = pd.DataFrame(data={"Date(DD/MM/YYYY)": datelst, "Hour(UTC)": hourlst})
#     monthtemplate.to_excel("monthTemplate.xlsx", sep=",", index=False)
