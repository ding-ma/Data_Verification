import csv
import os
import datetime

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
endDate = datetime.datetime(2019, 6, 5)
delta = endDate-startDate
listofDate = []

for i in range(delta.days+1):
    listofDate.append((startDate+datetime.timedelta(days=i)).strftime("%Y%m%d"))

stationlst = []
datelst =[]
hourlst=[]
O3lst=[]
for files in os.listdir("set_of_data/50126/"):
    for days in listofDate:
        if files.endswith("50126_" + days + ".csv"):
            f = open("set_of_data/50126/"+files, "r")
            print(files)
            csvFile = list(csv.reader(f))
            for z in range(len(csvFile)):
                row = csvFile[z]
                station = row[0]
                stationlst.append(station)
                print(station)
                date = row[1]
                datelst.append(date)
                hour = row[2]
                hourlst.append(hour)
                O3 = row[4]
                O3lst.append(O3)
                NO2 = row[5]
                PM25 = row[6]

    output = open("monthfiles/"+startDate.strftime("%Y%m")+".csv","w+")
    for s,d,h,o3 in zip(stationlst,datelst,hourlst,O3lst):
        print(s+","+d+","+h+","+o3+"\n")
        output.write(s+","+d+","+h+","+o3+"\n")

# stationsLst = open("stations.txt", "r").read().strip().split(",")
# for station in stationsLst:
#     stationID = PB_NAPS_dict[station]
#     if os.path.exists("set_of_data/"+ PB_NAPS_dict[station]) is True:
#         for files in os.listdir("set_of_data/"+stationID):
#             for days in listofDate:
#                 if files.endswith(stationID+"_" + days + ".csv"):
#                     f = open("set_of_data/"+stationID+"/" + files, "r")
#                     csvFile = list(csv.reader(f))
#                     for z in range(len(csvFile)):
#                         row = csvFile[z]
#                         station = row[0]
#                         date = row[1]
#                         hour = row[2]
#                         O3 = row[4]
#                         NO2 = row[5]
#                         PM25 = row[6]
#                         print(stationID, date, hour, O3)

