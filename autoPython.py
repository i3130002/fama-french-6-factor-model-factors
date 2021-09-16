# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %% [markdown]
# # Prepare excel read with pandas

# %%
import pandas as pd
# data = pd.read_excel(r'monthly_market_data - Copy.xlsx')
data = pd.read_excel(r'monthly_market_data - Copy (2).xlsx')

# print(data)

# %% [markdown]
# # Read Excel to memnory and create classes

# %%
import operator

class CompanyInfo:
    def __init__(self, tickerKey, name, month,numberOfOutstandingShares, marketcap,b2m, adjClose,inv,op,sourceSubsectorCode
):
        self.tickerKey = tickerKey
        self.name = name
        self.month = month
        self.rollingAverage = None
        self.numberOfOutstandingShares = numberOfOutstandingShares if not pd.isnull(numberOfOutstandingShares) else None
        self.marketcap = marketcap if not pd.isnull(marketcap) else None
        self.b2m = b2m if not pd.isnull(b2m) else None
        self.closeAdjust = adjClose
        self.momentum = None
        self.size = None
        self.RMWSize = None
        self.CMASize = None
        self.BperMSize = None
        self.currentYield = None
        self.inv = inv if not pd.isnull(inv) else None
        self.op = op if not pd.isnull(op) else None
        self.sourceSubsectorCode = sourceSubsectorCode


    def __repr__(self):
         return self.__str__()

    def __str__(self):
        return "tickerKey:"+str(self.tickerKey) +             "\t month:" + str(self.month) +             "\t closeAdjust:" + str(self.closeAdjust) +             "\t currentYield:" + str(self.currentYield) +             "\t numberOfOutstandingShares:" + str(self.numberOfOutstandingShares) +             "\t marketcap:" + str(self.marketcap) +             "\t b2m:" + str(self.b2m) +             "\t rollingAverage:" + str(self.rollingAverage) +             "\t momentum:" + str(self.momentum) +             "\t size:" + str(self.size)+"\n"




companyInfoList = []
companyInfoDict  = {}

for i, companyInfoPandas in data.iterrows():
    ci = CompanyInfo(tickerKey=companyInfoPandas["TickerKey"],
                                       name=companyInfoPandas["TickerNamePooyaFA"],
                                       month=companyInfoPandas["DayKeyFA"]// 100,
                                       numberOfOutstandingShares=companyInfoPandas["NumberOfOutstandingShares"],
                                       marketcap=companyInfoPandas["marketcap"],
                                       b2m=companyInfoPandas["B2M"],
                                       adjClose=companyInfoPandas["AdjClose"],
                                       inv=companyInfoPandas["INV"],
                                       op=companyInfoPandas["OP"],
                                       sourceSubsectorCode=companyInfoPandas["SourceSubsectorCode"],
                                       )
    companyInfoList.append(ci)
    if companyInfoPandas["TickerKey"] not in companyInfoDict:
      companyInfoDict[companyInfoPandas["TickerKey"]]={}
    companyInfoDict[companyInfoPandas["TickerKey"]][companyInfoPandas["DayKeyFA"]// 100]=ci
    
print(companyInfoList[:5])

# %% [markdown]
# # Dict access sample and test

# %%
print(list(companyInfoDict.keys())[:10])
print(list(companyInfoDict[1].keys())[:10])
print(companyInfoDict[1][139712])

# %% [markdown]
# # Extract company names

# %%
companyTickerSet = list(set([companyInfo.tickerKey for companyInfo in companyInfoList]))
print(companyTickerSet[0:5])

# %% [markdown]
# # Extract company months

# %%
allMonths = list(set([companyInfo.month for companyInfo in companyInfoList]))
allMonths.sort()
print(allMonths)
print("Total:",len(allMonths))

# %% [markdown]
# # Fill missing close adjust

# %%
def calculateCloseAdjust(companyHistoricalData, allMonths, monthIndex):
  privuseMonthCompanyInfo = [x for x in companyHistoricalData if x.month == allMonths[monthIndex - 1]]
  if len(privuseMonthCompanyInfo) == 0:
    return None
  privuseMonthCompanyInfo:CompanyInfo = privuseMonthCompanyInfo[0]
  nextMonthCompanyInfo = [x for x in companyHistoricalData if x.month == allMonths[monthIndex + 1 ]]

  if len(nextMonthCompanyInfo) == 1:
    nextMonthCompanyInfo:CompanyInfo = nextMonthCompanyInfo[0]
    return ((nextMonthCompanyInfo.closeAdjust / privuseMonthCompanyInfo.closeAdjust) ** (1/2)) * privuseMonthCompanyInfo.closeAdjust

  if monthIndex + 2 >= len(allMonths):
    return None
  secondNextMonthCompanyInfo = [x for x in companyHistoricalData if x.month == allMonths[monthIndex + 2 ]]
  if len(secondNextMonthCompanyInfo) == 0:
    return None
  secondNextMonthCompanyInfo:CompanyInfo = secondNextMonthCompanyInfo[0]
  return ((secondNextMonthCompanyInfo.closeAdjust / privuseMonthCompanyInfo.closeAdjust) ** (1/3)) * privuseMonthCompanyInfo.closeAdjust
  
  # print(calculatedCloseAdjust)

_missingData = []
for tickerKey in companyTickerSet:
  companyHistoricalData = [companyInfo for companyInfo in companyInfoList if companyInfo.tickerKey == tickerKey]
  startOfData = False
  for month in range(1,len(allMonths)-1):
    currentCompanyInfo = [x for x in companyHistoricalData if x.month == allMonths[month]]
    if len(currentCompanyInfo) > 1:
        raise Exception("A company cannot have two similar months:"+ tickerKey+" month:"+ allMonths[month])
    if len(currentCompanyInfo) == 1:
        startOfData = True
        continue
    if len(currentCompanyInfo) == 0:
      closeAdjust = calculateCloseAdjust(companyHistoricalData,allMonths,month)
      if closeAdjust is None: 
        if startOfData :
          _missingData.append(["closeAdjust for tick",tickerKey," Month",allMonths[month]," Is None"])
        continue
      if type(closeAdjust) != float :
        print("different type " ,(closeAdjust))
      privuseMonthCompanyInfo = [x for x in companyHistoricalData if x.month == allMonths[month - 1]][0]

      adjustedCompanyInfo = CompanyInfo(tickerKey,companyHistoricalData[0].name,allMonths[month],None,None,None,None,None,None,companyHistoricalData[0].sourceSubsectorCode)
      adjustedCompanyInfo.closeAdjust = closeAdjust
      print(tickerKey,allMonths[month],privuseMonthCompanyInfo) # log adjusted

      adjustedCompanyInfo.marketcap = closeAdjust * privuseMonthCompanyInfo.numberOfOutstandingShares
      adjustedCompanyInfo.numberOfOutstandingShares = privuseMonthCompanyInfo.numberOfOutstandingShares
      companyInfoList.append(adjustedCompanyInfo)
      # print(tickerKey,allMonths[month]) # log adjusted

print(_missingData)
# print([companyInfo for companyInfo in _missingData if companyInfo[1] == 9])
#//todo this has problem with 2 missing data


# %%
print(_missingData)

# %% [markdown]
# # Test data

# %%

dataToShow = [companyInfo for companyInfo in companyInfoList if companyInfo.tickerKey == 9]
dataToShow.sort(key=operator.attrgetter("month"), reverse=False)
print(dataToShow[:10])

# %% [markdown]
# # Yield calculations

# %%
for tickerKey in companyTickerSet:
  companyHistoricalData = [companyInfo for companyInfo in companyInfoList if companyInfo.tickerKey == tickerKey]
  for month in range(1,len(allMonths)):
    lastMonthData = [companyInfo for companyInfo in companyHistoricalData if companyInfo.month == allMonths[month - 1]]
    if len(lastMonthData)==0:
      continue
    thisMonthData = [companyInfo for companyInfo in companyHistoricalData if companyInfo.month == allMonths[month]]
    if len(thisMonthData)==0:
      continue
    lastMonthData:CompanyInfo = lastMonthData[0]
    thisMonthData:CompanyInfo = thisMonthData[0]
    thisMonthData.currentYield = thisMonthData.closeAdjust / lastMonthData.closeAdjust - 1
    

dataToShow = [companyInfo for companyInfo in companyInfoList if companyInfo.tickerKey == 9]
dataToShow.sort(key=operator.attrgetter("month"), reverse=False)
print(dataToShow[:10])

# %% [markdown]
# # Calculate rolling average

# %%
def getRollingDataWindow(companyHistoricalData,allMonths,rollingWindowSize,endMonthIndex):
  rollingDataWindow = []
  for monthIndex in range(endMonthIndex+1 - rollingWindowSize,endMonthIndex+1):
    currentCompanyInfo = [x for x in companyHistoricalData if x.month == allMonths[monthIndex]]
    if len(currentCompanyInfo) == 0:
        return rollingDataWindow
    if len(currentCompanyInfo) > 1:
        raise Exception("A company cannot have two similar months:"+ tickerKey+" month:"+ allMonths[monthIndex])
    rollingDataWindow.append(currentCompanyInfo[0])
  return rollingDataWindow


# %%
def average(dataList):
    if len(dataList)==0:
        return None
    data = [info.currentYield for info in dataList]
    if None in data:
      return None
    # print(data)
    return sum(data) / len(dataList)

# print(companyInfoList)
rollingAverageWindowSize = 12     #Can be change
for tickerKey in companyTickerSet:
    # print("tickerKey",tickerKey)
    companyHistoricalData = [companyInfo for companyInfo in companyInfoList if companyInfo.tickerKey == tickerKey]
    for endMonthIndex in range(rollingAverageWindowSize-1,len(allMonths)):
      # print("endMonthIndex",endMonthIndex)
      rollingDataWindow = getRollingDataWindow(companyHistoricalData,allMonths, rollingAverageWindowSize,endMonthIndex)
      if len(rollingDataWindow) != rollingAverageWindowSize: continue
      currentCompanyInfo = [x for x in companyHistoricalData if x.month == allMonths[endMonthIndex]]
      currentCompanyInfo:CompanyInfo = currentCompanyInfo[0]
      currentCompanyInfo.rollingAverage = average(rollingDataWindow)

print(companyInfoList[:10])

# %% [markdown]
# # set size

# %%
for month in allMonths:
    companiesInMonth = [x for x in companyInfoList if x.month== month]
    companiesInMonth.sort(key=operator.attrgetter("marketcap"), reverse=True)
    for company in companiesInMonth[0:len(companiesInMonth) // 2]:
        company.size = "Big"
    for company in companiesInMonth[len(companiesInMonth) // 2:]:
        company.size = "Small"

print ([x for x in companiesInMonth if x.size ][-10:])


# %%
for month in allMonths:
    companiesInMonth = [x for x in companyInfoList if x.month== month and x.rollingAverage is not None]
    companiesInMonth.sort(key=operator.attrgetter("rollingAverage"), reverse=True)
    for company in companiesInMonth[0:len(companiesInMonth) // 2]:
        company.momentum = "High"
    for company in companiesInMonth[len(companiesInMonth) // 2:]:
        company.momentum = "Low"

print ([x for x in companyInfoList if x.rollingAverage ][-10:])

# %% [markdown]
# # MomentomFactor factor (MOM)

# %%


def findCompanyInfo(companyInfoListToSearch, month, tickerKey):
    for companyInfo in companyInfoListToSearch:
        if companyInfo.month == month and companyInfo.tickerKey == tickerKey:
            return companyInfo
    return None


MoMFactor = {}

for index,month in enumerate(allMonths[:-1]):
    # Get companies of corrent month
    companiesInMonth = [x for x in companyInfoList if x.month== month]
    companiesInNextMonth = [x for x in companyInfoList if x.month== allMonths[index+1]]
    # Get companis for these filters SL SH BL BH
    SL = [company for company in companiesInMonth if company.size == "Small" and company.momentum == "Low"]
    SH = [company for company in companiesInMonth if company.size == "Small" and company.momentum == "High"]
    BL = [company for company in companiesInMonth if company.size == "Big" and company.momentum == "Low"]
    BH = [company for company in companiesInMonth if company.size == "Big" and company.momentum == "High"]
    # Get companies next month info
    BH_average = [findCompanyInfo(companiesInNextMonth,allMonths[index+1],company.tickerKey) for company in BH]
    BL_average = [findCompanyInfo(companiesInNextMonth,allMonths[index+1],company.tickerKey) for company in BL]
    SH_average = [findCompanyInfo(companiesInNextMonth,allMonths[index+1],company.tickerKey) for company in SH]
    SL_average = [findCompanyInfo(companiesInNextMonth,allMonths[index+1],company.tickerKey) for company in SL]
    # Get companies next month currentYield
    BH_average = average([x for x in BH_average if x is not None])
    BL_average = average([x for x in BL_average if x is not None])
    SH_average = average([x for x in SH_average if x is not None])
    SL_average = average([x for x in SL_average if x is not None])
    
    # If all were filled, Add to momentomFactor
    if (BH_average and BL_average and SH_average and SL_average ):
        MoMFactor[month]=(0.5 * (SH_average + BH_average) - 0.5 * (SL_average + BL_average))

print(MoMFactor)


# %%
# Display
 
output = ""
companyTickerSet.sort()
output+="Data\TickerKey,"

for company in companyTickerSet:
  output+=str(company) + ","
output += "\n"
for month in allMonths:
    
    output+=str(month) + ","
    for company in companyTickerSet[:10]:
      dataToShow = [x for x in companyInfoList if x.month== month and x.tickerKey==company]
      if len(dataToShow)!=1:
        output+= ","
      else:
        output+=str(dataToShow[0].currentYield) + ","
    output += "\n"
    

print(output[:1000])

# %% [markdown]
# # Book to market Calc

# %%
#  Fill all marketcap based on the last one
for companyTicker in companyTickerSet:
  for index,month in enumerate(allMonths):
    if month not in companyInfoDict[companyTicker]:
      firstData:CompanyInfo = companyInfoDict[companyTicker][list(companyInfoDict[companyTicker].keys())[0]]
      c = CompanyInfo(companyTicker,firstData.name,month,None,None,None,None,None,firstData.sourceSubsectorCode)
      if index > 0 and allMonths[index-1] in companyInfoDict[companyTicker]:
        c.marketcap = companyInfoDict[companyTicker][allMonths[index-1]].marketcap
        c.b2m = companyInfoDict[companyTicker][allMonths[index-1]].b2m
        c.inv = companyInfoDict[companyTicker][allMonths[index-1]].inv
        c.op = companyInfoDict[companyTicker][allMonths[index-1]].op
      companyInfoDict[companyTicker][month] = c
  # print(companyInfoDict[companyTicker][139712])

for month in sorted(companyInfoDict[1].keys())[:10]:
  print(companyInfoDict[9][month])


# %%
for companyTicker in companyTickerSet:
  lastB2MOfTheYear = None
  lastINVOfTheYear = None
  lastOpOfTheYear = None
  for month in allMonths:
    if month % 100 == 5:
      lastB2MOfTheYear = companyInfoDict[companyTicker][month].b2m
      lastINVOfTheYear = companyInfoDict[companyTicker][month].inv
      lastOpOfTheYear = companyInfoDict[companyTicker][month].op
    companyInfoDict[companyTicker][month].b2m = lastB2MOfTheYear     
    companyInfoDict[companyTicker][month].inv = lastINVOfTheYear     
    companyInfoDict[companyTicker][month].op = lastOpOfTheYear     


# %%
# for month in sorted(companyInfoDict[1].keys())[:10]:
#   print(companyInfoDict[9][month])


# %%
# Book To market ranking

for month in allMonths:
  b2mMonthList = []
  for companyTicker in companyTickerSet:
    if companyInfoDict[companyTicker][month].b2m == None:
      continue
    if companyInfoDict[companyTicker][month].currentYield == None:
      continue
    b2mMonthList.append(companyInfoDict[companyTicker][month].b2m)
  if len(b2mMonthList) == 0 :
    continue
  b2mMonthList.sort()
  # print(len(b2mMonthList),int(len(b2mMonthList)*0.3))
  lowBookToMarket = b2mMonthList[int(len(b2mMonthList)*0.3)]
  highBookToMarket = b2mMonthList[int(-len(b2mMonthList)*0.3)]

  for companyTicker in companyTickerSet:
    if companyInfoDict[companyTicker][month].b2m == None:
      continue
    if companyInfoDict[companyTicker][month].currentYield == None:
      continue
    if companyInfoDict[companyTicker][month].b2m <= lowBookToMarket:
      companyInfoDict[companyTicker][month].BperMSize = "Low"
    elif companyInfoDict[companyTicker][month].b2m >= highBookToMarket:
      companyInfoDict[companyTicker][month].BperMSize = "High"
    else:
      companyInfoDict[companyTicker][month].BperMSize = "Neutral"


# %%
# RMW (Operational profit) ranking

for month in allMonths:
  dataMonthList = []
  for companyTicker in companyTickerSet:
    if companyInfoDict[companyTicker][month].op == None:
      continue
    if companyInfoDict[companyTicker][month].currentYield == None:
      continue
    dataMonthList.append(companyInfoDict[companyTicker][month].op)
  if len(dataMonthList) == 0 :
    continue
  dataMonthList.sort()
  # print(len(dataMonthList),int(len(dataMonthList)*0.3))
  weakValue = dataMonthList[int(len(dataMonthList)*0.3)]
  robustValue = dataMonthList[int(-len(dataMonthList)*0.3)]
  # if month == 139902:
  #   print(dataMonthList)
  for companyTicker in companyTickerSet:
    if companyInfoDict[companyTicker][month].op == None:
      continue
    if companyInfoDict[companyTicker][month].currentYield == None:
      continue
    if companyInfoDict[companyTicker][month].op <= weakValue:
      companyInfoDict[companyTicker][month].RMWSize = "Weak"
    elif companyInfoDict[companyTicker][month].op >= robustValue:
      companyInfoDict[companyTicker][month].RMWSize = "Robust"
    else:
      companyInfoDict[companyTicker][month].RMWSize = "Neutral"


# %%

for companyTicker in companyTickerSet[:10]:
  print(companyInfoDict[companyTicker][139902].RMWSize)


# %%
# CMA (Investment) ranking

for month in allMonths:
  dataMonthList = []
  for companyTicker in companyTickerSet:
    if companyInfoDict[companyTicker][month].inv == None:
      continue
    if companyInfoDict[companyTicker][month].currentYield == None:
      continue    
    dataMonthList.append(companyInfoDict[companyTicker][month].inv)
  if len(dataMonthList) == 0 :
    continue
  dataMonthList.sort()
  # print(len(dataMonthList),int(len(dataMonthList)*0.3))
  weakValue = dataMonthList[int(len(dataMonthList)*0.3)]
  robustValue = dataMonthList[int(-len(dataMonthList)*0.3)]
  # if month == 139902:
  #   print(dataMonthList)
  for companyTicker in companyTickerSet:
    if companyInfoDict[companyTicker][month].inv == None:
      continue
    if companyInfoDict[companyTicker][month].currentYield == None:
      continue
    if companyInfoDict[companyTicker][month].inv <= weakValue:
      companyInfoDict[companyTicker][month].CMASize = "Conservative"
    elif companyInfoDict[companyTicker][month].inv >= robustValue:
      companyInfoDict[companyTicker][month].CMASize = "Aggressive"
    else:
      companyInfoDict[companyTicker][month].CMASize = "Neutral"


# %%
def getNextMonth(currentMonth):
  return allMonths[allMonths.index(currentMonth)+1]

def getNextMonthCompaniesHavingYeild(companyInfoList):
  response = []
  for companyInfo in companyInfoList:
    nextMoonth = companyInfoDict[companyInfo.tickerKey][getNextMonth(companyInfo.month)]
    if nextMoonth.currentYield != None:
      response.append(nextMoonth)
  return response


# %%
HMLFactor = {}
SMBBTMFactor = {}
for month in allMonths[:-1]:
  SH = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].BperMSize == "High" and         companyInfoDict[companyTicker][month].size == "Small"]
  SN = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].BperMSize == "Neutral" and         companyInfoDict[companyTicker][month].size == "Small"]
  SL = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].BperMSize == "Low" and         companyInfoDict[companyTicker][month].size == "Small"]
  BH = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].BperMSize == "High" and         companyInfoDict[companyTicker][month].size == "Big"]
  BN = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].BperMSize == "Neutral" and         companyInfoDict[companyTicker][month].size == "Big"]
  BL = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].BperMSize == "Low" and         companyInfoDict[companyTicker][month].size == "Big"]
  
  
  SH_average = average(getNextMonthCompaniesHavingYeild(SH))
  SN_average = average(getNextMonthCompaniesHavingYeild(SN))
  SL_average = average(getNextMonthCompaniesHavingYeild(SL))
  BH_average = average(getNextMonthCompaniesHavingYeild(BH))
  BN_average = average(getNextMonthCompaniesHavingYeild(BN))
  BL_average = average(getNextMonthCompaniesHavingYeild(BL))
  
  if BL_average == None:
    continue
  
  SMBBTMFactor[month] = (SH_average + SN_average + SL_average) / 3 - (BH_average + BN_average + BL_average) / 3 
  HMLFactor[month] = (SH_average + BH_average) / 2 - (SL_average + BL_average) / 2 
for month in list(HMLFactor.keys())[:10]:
  print(HMLFactor[month])


# %%
RMWFactor = {}
SMBRMWFactor = {}
for month in allMonths[:-1]:
  SR = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].RMWSize == "Robust" and         companyInfoDict[companyTicker][month].size == "Small"]
  SN = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].RMWSize == "Neutral" and         companyInfoDict[companyTicker][month].size == "Small"]
  SW = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].RMWSize == "Weak" and         companyInfoDict[companyTicker][month].size == "Small"]
  BR = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].RMWSize == "Robust" and         companyInfoDict[companyTicker][month].size == "Big"]
  BN = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].RMWSize == "Neutral" and         companyInfoDict[companyTicker][month].size == "Big"]
  BW = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].RMWSize == "Weak" and         companyInfoDict[companyTicker][month].size == "Big"]
  
  SR_average = average(getNextMonthCompaniesHavingYeild(SR))
  SN_average = average(getNextMonthCompaniesHavingYeild(SN))
  SW_average = average(getNextMonthCompaniesHavingYeild(SW))
  BR_average = average(getNextMonthCompaniesHavingYeild(BR))
  BN_average = average(getNextMonthCompaniesHavingYeild(BN))
  BW_average = average(getNextMonthCompaniesHavingYeild(BW))

  if SR_average == None:
    continue
  SMBRMWFactor[month] = (SR_average + SN_average + SW_average) / 3 - (BR_average + BN_average + BW_average) / 3 
  RMWFactor[month] = (SR_average + BW_average) / 2 - (SW_average + BW_average) / 2 
for month in list(RMWFactor.keys())[:10]:
  print(RMWFactor[month])


# %%
CMAFactor = {}
SMBCMAFactor = {}
for month in allMonths[:-1]:
  SA = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].CMASize == "Aggressive" and         companyInfoDict[companyTicker][month].size == "Small"]
  SN = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].CMASize == "Neutral" and         companyInfoDict[companyTicker][month].size == "Small"]
  SC = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].CMASize == "Conservative" and         companyInfoDict[companyTicker][month].size == "Small"]
  BA = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].CMASize == "Aggressive" and         companyInfoDict[companyTicker][month].size == "Big"]
  BN = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].CMASize == "Neutral" and         companyInfoDict[companyTicker][month].size == "Big"]
  BC = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].CMASize == "Conservative" and         companyInfoDict[companyTicker][month].size == "Big"]
  
  SA_average = average(getNextMonthCompaniesHavingYeild(SA))
  SN_average = average(getNextMonthCompaniesHavingYeild(SN))
  SC_average = average(getNextMonthCompaniesHavingYeild(SC))
  BA_average = average(getNextMonthCompaniesHavingYeild(BA))
  BN_average = average(getNextMonthCompaniesHavingYeild(BN))
  BC_average = average(getNextMonthCompaniesHavingYeild(BC))

  if SA_average == None:
    continue
  SMBCMAFactor[month] = (SA_average + SN_average + SC_average) / 3 - (BA_average + BN_average + BC_average) / 3 
  CMAFactor[month] = (SC_average + BC_average) / 2 - (SA_average + BA_average) / 2 
for month in list(CMAFactor.keys())[:10]:
  print(CMAFactor[month])


# %%
SMBFactor = {}
for month in SMBBTMFactor.keys():
  SMBFactor[month] = (SMBBTMFactor[month] + SMBCMAFactor[month] + SMBRMWFactor[month]) / 3


# %%
# Factors to excel
rows = []
for month in allMonths:
     rows.append([
       month,
       SMBFactor[month] if month in SMBFactor else None,
       HMLFactor[month] if month in HMLFactor else None,
       RMWFactor[month] if month in RMWFactor else None,
       CMAFactor[month] if month in CMAFactor else None,
       MoMFactor[month] if month in MoMFactor else None,
     ])
dfOut = pd.DataFrame(rows, columns=['month',"SMBFactor", "HMLFactor", "RMWFactor", "CMAFactor", "MoMFactor"])
dfOut.to_excel('Factors.xlsx', sheet_name='Factors')


