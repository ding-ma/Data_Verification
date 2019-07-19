import csv
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

stationsLstO3 = open("stationsO3.txt", "r").read().strip().split(",")
stationsLstNO2 = open("stationsNO2.txt", "r").read().strip().split(",")
stationsLstPM25 = open("stationsPM25.txt", "r").read().strip().split(",")

template = open("monthTemplate.csv", "r")
csvr = list(csv.reader(template))
hourlst = []
daylst = []
for q in range(len(csvr)):
    r = csvr[q]
    d = r[0]
    h = r[1]
    daylst.append(d)
    hourlst.append(h)

cclst = []
finalFile = pd.DataFrame(data={"Date(DD/MM/YYYY)": daylst[1:], "Hours(UTC)": hourlst[1:]})


def generateMonthReport(stat, polluant):
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
        finalFile.to_csv("output/monthlyreport" + polluant + ".csv", sep=",", index=False)
        cclst.clear()


generateMonthReport(stationsLstNO2, "NO2")
generateMonthReport(stationsLstO3, "O3")
generateMonthReport(stationsLstPM25, "PM25")
