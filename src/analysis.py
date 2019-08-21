import pandas as pd

df = pd.read_csv("fromserver/PM25Anal.csv")
df.set_index(['Date(DD/MM/YYYY)', 'Hours(UTC)'], inplace=True)
stationlst = list(df.columns.values.tolist())

PM25threshold = 35
for station in stationlst:
    condition = (df[station] > PM25threshold) & \
                (df[station].shift(-1) > PM25threshold) & \
                (df[station].shift(-2) > PM25threshold)

    consecutive_count = df[condition].count().head(1)[0]
    # print(df[condition].count().head(1))
    # if consecutive_count > 0:
    #     print(consecutive_count,station)

print(df.loc[:, 'CMSS'])
