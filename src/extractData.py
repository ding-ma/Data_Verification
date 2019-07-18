import csv
import datetime
import os

import pandas as pd

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

startDate = datetime.datetime(2019, 6, 1)
endDate = datetime.datetime(2019, 6, 30)

delta = endDate-startDate
listofDate = []

for i in range(delta.days+1):
    listofDate.append((startDate+datetime.timedelta(days=i)).strftime("%Y%m%d"))

stationlst = []
datelst =[]
hourlst=[]
O3lst=[]
NO2lst = []
PM25lst = []
stationsLstNO2 = open("stationsNO2.txt", "r").read().strip().split(",")
stationsLstO3 = open("stationsO3.txt", "r").read().strip().split(",")
stationsLstPM25 = open("stationsPM25.txt", "r").read().strip().split(",")


def getPerMonth(lst, polluant):
    for station in stationsLstNO2:
        stationID = PB_NAPS_dict[station]
        if os.path.exists(
                "/fs/home/fs1/ords/oth/airq_central/frc002/Data/CAS/Observations/Station/" + stationID) is True:
            for files in sorted(
                    os.listdir("/fs/home/fs1/ords/oth/airq_central/frc002/Data/CAS/Observations/Station/" + stationID)):
                for days in listofDate:
                    if files.endswith(stationID + "_" + days + ".csv"):
                        print(files)
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

        df1 = pd.DataFrame(data={"Date(DD/MM/YYYY)": datelst, "Concentration": NO2lst, "Hour(UTC)": hourlst})
        df1.to_csv("monthfiles/NO2/" + startDate.strftime("%Y%m") + "_" + stationID + "NO2.csv", sep=",", index=False)
        monthtemplate = pd.DataFrame(data={"Date(DD/MM/YYYY)": datelst, "Hour(UTC)": hourlst})
        monthtemplate.to_csv("monthTemplate.csv", sep=",", index=False)
        stationlst.clear()
        datelst.clear()
        hourlst.clear()
        NO2lst.clear()


for station in stationsLstO3:
    stationID = PB_NAPS_dict[station]
    if os.path.exists("/fs/home/fs1/ords/oth/airq_central/frc002/Data/CAS/Observations/Station/" + stationID) is True:
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
                        if row != ['station', 'Date', 'UTC', 'AQHI', 'O3', 'NO2', 'PM2.5', 'PM10', 'SO2', 'H2S', 'CO',
                                   'NO', 'TRS']:
                            station = row[0]
                            stationlst.append(station)
                            date = row[1]
                            datelst.append(date)
                            hour = row[2]
                            hourlst.append(hour)
                            O3 = row[4]
                            O3lst.append(O3)

    df = pd.DataFrame(data={"Date(DD/MM/YYYY)": datelst, "Concentration": O3lst, "Hour(UTC)": hourlst})
    df.to_csv("monthfiles/O3/" + startDate.strftime("%Y%m") + "_" + stationID + "O3.csv", sep=",", index=False)
    monthtemplate = pd.DataFrame(data={"Date(DD/MM/YYYY)": datelst, "Hour(UTC)": hourlst})
    monthtemplate.to_csv("monthTemplate.csv", sep=",", index=False)
    stationlst.clear()
    datelst.clear()
    hourlst.clear()
    O3lst.clear()

for station in stationsLstPM25:
    stationID = PB_NAPS_dict[station]
    if os.path.exists("/fs/home/fs1/ords/oth/airq_central/frc002/Data/CAS/Observations/Station/" + stationID) is True:
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
                        if row != ['station', 'Date', 'UTC', 'AQHI', 'O3', 'NO2', 'PM2.5', 'PM10', 'SO2', 'H2S', 'CO',
                                   'NO', 'TRS']:
                            station = row[0]
                            stationlst.append(station)
                            date = row[1]
                            datelst.append(date)
                            hour = row[2]
                            hourlst.append(hour)
                            PM25 = row[6]
                            PM25lst.append(PM25)

    df2 = pd.DataFrame(data={"Date(DD/MM/YYYY)": datelst, "Concentration": PM25lst, "Hour(UTC)": hourlst})
    df2.to_csv("monthfiles/PM25/" + startDate.strftime("%Y%m") + "_" + stationID + "PM25.csv", sep=",", index=False)

    monthtemplate = pd.DataFrame(data={"Date(DD/MM/YYYY)": datelst, "Hour(UTC)": hourlst})
    monthtemplate.to_csv("monthTemplate.csv", sep=",", index=False)
    stationlst.clear()
    datelst.clear()
    hourlst.clear()
    PM25lst.clear()
