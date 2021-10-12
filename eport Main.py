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

    def convertToList(self):
      return [getattr(self,atr) for atr in self.__dict__]


companyInfoDict  = {}
def addCompanyInfoToList(companyInfo:CompanyInfo):
    if companyInfo.tickerKey not in companyInfoDict:
      companyInfoDict[companyInfo.tickerKey]={}
    if companyInfo.month not in companyInfoDict[companyInfo.tickerKey]:
      companyInfoDict[companyInfo.tickerKey][companyInfo.month]=companyInfo
    else:
      raise Exception("Company info exists." + str(companyInfo))

def findCompanyInfo(tickerKey, month):
    if tickerKey in companyInfoDict and month in companyInfoDict[tickerKey]:
      return companyInfoDict[tickerKey][month]
    return None

def displayCompanyInfoHeader():
  for tickerKey in companyInfoDict.keys():
      for month in companyInfoDict[tickerKey].keys():
        company:CompanyInfo = findCompanyInfo(tickerKey,month)
        if company is None:
          continue
        attrs = []
        for attribute in vars(company):
          attrs.append(attribute)
        print(attrs)
        return

def displayCompanyInfo(tickerKeys:list,months:list):
    displayCompanyInfoHeader()
    for tickerKey in tickerKeys:
      for month in months:
        company:CompanyInfo = findCompanyInfo(tickerKey,month)
        if company is None:
          continue
        print(company.convertToList())

allMonths = []


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
    addCompanyInfoToList(ci)
    allMonths.append(ci.month)

allMonths = list(set(allMonths))
allMonths.sort()


# %%
# companies to excel
def exportCompanyListToExcel(excelFileName):
  rows = []
  for company in companyInfoDict.keys():
    for month in companyInfoDict[company].keys():
      rows.append(companyInfoDict[company][month].convertToList())

  company_ = None
  for tickerKey in companyInfoDict.keys():
      for month in companyInfoDict[tickerKey].keys():
        company:CompanyInfo = findCompanyInfo(tickerKey,month)
        if company is None:
          continue
        company_ = company
        break
      if company_ is not None:
        break
  dfOut = pd.DataFrame(rows,columns=company_.__dict__)
  dfOut.to_excel(excelFileName+'.xlsx', sheet_name='TMPExport')

# %% [markdown]
# # Dict access sample and test

# %%
print(list(companyInfoDict.keys())[:10])
print(list(companyInfoDict[1].keys())[:10])
print(companyInfoDict[1][139712])

# %% [markdown]
# # Extract company names

# %%
companyTickerSet = list(companyInfoDict.keys())
print(companyTickerSet)

# %% [markdown]
# # Extract company months

# %%
print(allMonths)
print("Total:",len(allMonths))


# %%
displayCompanyInfo(companyTickerSet[:3],allMonths[:3])

# %% [markdown]
# # Fill missing close adjust

# %%
def calculateCloseAdjust(tickerKey, monthIndex):
  privuseMonthCompanyInfo:CompanyInfo = findCompanyInfo(tickerKey, allMonths[monthIndex-1])
  if privuseMonthCompanyInfo is None:
    return None
  nextMonthCompanyInfo:CompanyInfo = findCompanyInfo(tickerKey, allMonths[monthIndex+1])

  if nextMonthCompanyInfo is not None:
    return ((nextMonthCompanyInfo.closeAdjust / privuseMonthCompanyInfo.closeAdjust) ** (1/2)) * privuseMonthCompanyInfo.closeAdjust

  if monthIndex + 2 >= len(allMonths):
    return None
  secondNextMonthCompanyInfo:CompanyInfo = findCompanyInfo(tickerKey, allMonths[monthIndex+2])
  if secondNextMonthCompanyInfo is None:
    return None
  return ((secondNextMonthCompanyInfo.closeAdjust / privuseMonthCompanyInfo.closeAdjust) ** (1/3)) * privuseMonthCompanyInfo.closeAdjust
  
  # print(calculatedCloseAdjust)

_missingData = []
for tickerKey in companyTickerSet:
  # companyHistoricalData = [companyInfo for companyInfo in companyInfoList if companyInfo.tickerKey == tickerKey]
  startOfData = False
  for monthIndex in range(1,len(allMonths)-1):
    currentCompanyInfo = findCompanyInfo(tickerKey,allMonths[monthIndex])
    if currentCompanyInfo is not None:
        startOfData = True
        continue
    if currentCompanyInfo is None:
      closeAdjust = calculateCloseAdjust(tickerKey,monthIndex)
      if closeAdjust is None: 
        if startOfData :
          _missingData.append(["closeAdjust for tick",tickerKey," Month",allMonths[monthIndex]," Is None"])
        continue
      privuseMonthCompanyInfo:CompanyInfo = findCompanyInfo(tickerKey, allMonths[monthIndex-1])

      adjustedCompanyInfo = CompanyInfo(tickerKey,privuseMonthCompanyInfo.name,allMonths[monthIndex],None,None,None,None,None,None,privuseMonthCompanyInfo.sourceSubsectorCode)
      adjustedCompanyInfo.closeAdjust = closeAdjust
      # print(tickerKey,allMonths[month],privuseMonthCompanyInfo) # log before adjusted
      if privuseMonthCompanyInfo.numberOfOutstandingShares == None:
        print("None privuse marketcap" ,privuseMonthCompanyInfo)
        continue;
      adjustedCompanyInfo.marketcap = closeAdjust/privuseMonthCompanyInfo.closeAdjust * privuseMonthCompanyInfo.marketcap
      if privuseMonthCompanyInfo.numberOfOutstandingShares == None:
        print("None numberOfOutstandingShares" ,privuseMonthCompanyInfo)
        continue;
      adjustedCompanyInfo.numberOfOutstandingShares = privuseMonthCompanyInfo.numberOfOutstandingShares
      addCompanyInfoToList(adjustedCompanyInfo)
      # print(tickerKey,allMonths[month]) # log adjusted


# print([companyInfo for companyInfo in _missingData if companyInfo[1] == 9])


# %%
print("{} of {}".format(len(_missingData),data.size))
print(_missingData[:10])

# %% [markdown]
# # Test data

# %%

displayCompanyInfo(companyTickerSet[:3],allMonths[:3])

# %% [markdown]
# # Yield calculations

# %%
for tickerKey in companyTickerSet:
  for monthIndex in range(1,len(allMonths)):
    lastMonthData:CompanyInfo = findCompanyInfo(tickerKey,allMonths[monthIndex-1])
    if lastMonthData is None:
      continue
    thisMonthData:CompanyInfo = findCompanyInfo(tickerKey,allMonths[monthIndex])
    if thisMonthData is None:
      continue
    thisMonthData.currentYield = thisMonthData.closeAdjust / lastMonthData.closeAdjust - 1


# %%
displayCompanyInfo(companyTickerSet[:3],allMonths[:3])

# %% [markdown]
# # Calculate rolling average

# %%
def getRollingDataWindow(tickerKey,rollingWindowSize,endMonthIndex):
  rollingDataWindow = []
  for monthIndex in range(endMonthIndex+1 - rollingWindowSize,endMonthIndex+1):
    currentCompanyInfo = findCompanyInfo(tickerKey,allMonths[monthIndex])
    if currentCompanyInfo is None:
        continue
    rollingDataWindow.append(currentCompanyInfo)
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
    for endMonthIndex in range(rollingAverageWindowSize-1,len(allMonths)):
      # print("endMonthIndex",endMonthIndex)
      rollingDataWindow = getRollingDataWindow(tickerKey, rollingAverageWindowSize,endMonthIndex)
      if len(rollingDataWindow) != rollingAverageWindowSize: continue
      currentCompanyInfo:CompanyInfo = findCompanyInfo(tickerKey,allMonths[endMonthIndex])
      currentCompanyInfo.rollingAverage = average(rollingDataWindow)


# %%
displayCompanyInfo(companyTickerSet[:3],allMonths[:3])

# %% [markdown]
# # Sample Export to Excel

# %%
# Example to excel function
exportCompanyListToExcel("AfterRollingAverage")

# %% [markdown]
# # set size

# %%
def getCompaniesInMonth(month):
  compainesInMonth = []
  for tickerKey in companyInfoDict.keys():
    if month in companyInfoDict[tickerKey]:
      compainesInMonth.append(companyInfoDict[tickerKey][month])
  return compainesInMonth


# %%
import math
for month in allMonths:
    companiesInMonth = [x for x in getCompaniesInMonth(month) if x.rollingAverage is not None and x.marketcap is not None]
    companiesInMonth.sort(key=operator.attrgetter("marketcap"), reverse=True)
    median = math.ceil(len(companiesInMonth) / 2)
    for company in companiesInMonth[0:median]:
        company.size = "Big"
    for company in companiesInMonth[median:]:
        company.size = "Small"

print ([x for x in getCompaniesInMonth(allMonths[-1]) if x.size ][-10:])


# %%

    
    
for index,month in enumerate(allMonths[:-1]):
    companiesInMonth = [x for x in getCompaniesInMonth(month) if x.rollingAverage is not None and x.marketcap is not None]
    companiesInMonth = [x for x in companiesInMonth if findCompanyInfo(x.tickerKey,allMonths[index+1]) is not None and findCompanyInfo(x.tickerKey,allMonths[index+1]).currentYield is not None]
    companiesInMonth.sort(key=operator.attrgetter("rollingAverage"), reverse=True)
    count = math.floor(len(companiesInMonth) / 3) #It might be better to round it up instead of down to increase valid data
    for company in companiesInMonth[0:count]:
        company.momentum = "High"
    for company in companiesInMonth[len(companiesInMonth) - count:]:
        company.momentum = "Low"
    for company in companiesInMonth[count:len(companiesInMonth) - count]:
        company.momentum = "Notural"

print ([x for x in getCompaniesInMonth(allMonths[-2]) if x.momentum is not None ][-10:])

# %% [markdown]
# # MomentomFactor factor (MOM)

# %%


MoMFactor = {}

for index,month in enumerate(allMonths[:-1]):
    # Get companies of corrent month
    companiesInMonth = getCompaniesInMonth(month)
    # Get companis for these filters SL SH BL BH
    SL = [company for company in companiesInMonth if company.size == "Small" and company.momentum == "Low"]
    SH = [company for company in companiesInMonth if company.size == "Small" and company.momentum == "High"]
    BL = [company for company in companiesInMonth if company.size == "Big" and company.momentum == "Low"]
    BH = [company for company in companiesInMonth if company.size == "Big" and company.momentum == "High"]
    # Get companies next month info
    BH_nextMonth = [findCompanyInfo(company.tickerKey,allMonths[index+1]) for company in BH]
    BL_nextMonth = [findCompanyInfo(company.tickerKey,allMonths[index+1]) for company in BL]
    SH_nextMonth = [findCompanyInfo(company.tickerKey,allMonths[index+1]) for company in SH]
    SL_nextMonth = [findCompanyInfo(company.tickerKey,allMonths[index+1]) for company in SL]
    # Get companies next month currentYield
    BH_average = average([x for x in BH_nextMonth if x is not None])
    BL_average = average([x for x in BL_nextMonth if x is not None])
    SH_average = average([x for x in SH_nextMonth if x is not None])
    SL_average = average([x for x in SL_nextMonth if x is not None])
    
    # If all were filled, Add to momentomFactor
    if (BH_average and BL_average and SH_average and SL_average ):
        MoMFactor[month]=(0.5 * (SH_average + BH_average) - 0.5 * (SL_average + BL_average))

print(MoMFactor)

# %% [markdown]
# # Book to market Calc

# %%
#  Fill all marketcap based on the last one
for companyTicker in companyTickerSet:
  for index,month in enumerate(allMonths):
    if month not in companyInfoDict[companyTicker]:
      firstData:CompanyInfo = companyInfoDict[companyTicker][list(companyInfoDict[companyTicker].keys())[0]]
      c = CompanyInfo(companyTicker,firstData.name,month,None,None,None,None,None,None, firstData.sourceSubsectorCode)
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


