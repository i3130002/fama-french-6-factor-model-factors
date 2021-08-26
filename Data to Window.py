import pandas as pd

data = pd.read_excel(r'monthly_market_data - Copy (2).xlsx')

exportdata = pd.DataFrame({"month": [], "date": []})

date = data['date']
date.drop_duplicates(keep='first', inplace=True)

date = list(date)
date.sort()
data = data.reset_index(drop=True)

for i, d in enumerate(date):
    data_d = data[data['date'] == d]
    data_d = data_d.reset_index(drop=True)
    olol = pd.DataFrame({"month": [i], "date": list(data_d["date"])[0]})
    for j, c in enumerate(list(data_d['TickerKey'])):
        olol[c] = data_d['AdjClose'][j]
    exportdata = exportdata.append(olol, ignore_index=True)

exportdata.to_excel("output.xlsx", index=False)
