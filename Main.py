import pandas as pd

data = pd.read_excel(r'monthly_market_data - Copy.xlsx')


class CompanyInfo:
    def __init__(self, tickerKey, name, month, marketcap, close):
        self.tickerKey = tickerKey
        self.name = name
        self.month = month
        self.rollingAverage = None
        self.marketcap = marketcap
        self.Close = close


companyInfoList = []

for i, companyInfoPandas in data.iterrows():
    companyInfoList.append(CompanyInfo(tickerKey=companyInfoPandas["TickerKey"],
                                       name=companyInfoPandas["TickerNamePooyaFA"],
                                       month=companyInfoPandas["DayKeyFA"],
                                       marketcap=companyInfoPandas["marketcap"],
                                       close=companyInfoPandas["Close"],
                                       ))

pass
