import csv

PM25File = open("StationPM25.csv", "r")
reader = list(csv.reader(PM25File))
PM25_StationIDlst = []
PM25_StationRegionlst = []

for x in range(len(reader)):
    line = reader[x]
    pmID = line[0]
    pmRegion = line[1]
    PM25_StationIDlst.append(pmID)
    PM25_StationRegionlst.append(pmRegion)

tempset = set()
regionDictionary = dict(zip(PM25_StationIDlst, PM25_StationRegionlst))
for w in PM25_StationIDlst:
    a = regionDictionary[w]
    tempset.add(PM25_StationRegionlst.index(a))
print(PM25_StationRegionlst)
print(regionDictionary["CGSS"])
x = list(sorted(tempset))
# for i, p in enumerate(x):
#     try:
#         print(x[i+1])
#     except:
#         x.append(x[-1])
#         print(x[i+1])
#         break
#
