# %% [markdown]
# # Prepare excel read with pandas

# %%
import pandas as pd
# data = pd.read_excel(r'monthly_market_data - Copy.xlsx')
data = pd.read_excel(r'CRSP.xlsx')

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
        self.sizeMOM = None
        self.sizeHML = None
        self.sizeRMW = None
        self.sizeCMA = None
        self.sizeLabel = None
        self.CMALabel = None
        self.RMWLabel = None
        self.B2MLabel = None
        self.currentYield = None
        self.inv = inv if not pd.isnull(inv) else None
        self.op = op if not pd.isnull(op) else None
        self.sourceSubsectorCode = sourceSubsectorCode


    def __repr__(self):
         return self.__str__()

    def __str__(self):
        return "tickerKey:"+str(self.tickerKey) + \
            "\t month:" + str(self.month) + \
            "\t closeAdjust:" + str(self.closeAdjust) + \
            "\t currentYield:" + str(self.currentYield) + \
            "\t numberOfOutstandingShares:" + str(self.numberOfOutstandingShares) + \
            "\t marketcap:" + str(self.marketcap) + \
            "\t b2m:" + str(self.b2m) + \
            "\t rollingAverage:" + str(self.rollingAverage) + \
            "\t momentum:" + str(self.momentum) + \
            "\t sizeLabel:" + str(self.sizeLabel) + \
            "\t sizeHML:" + str(self.sizeHML) + \
            "\t sizeRMW:" + str(self.sizeRMW) + \
            "\t sizeCMA:" + str(self.sizeCMA) +"\n"

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
      sortedMonths = sorted(months)
      for month in sortedMonths:
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
    for month in sorted(companyInfoDict[company].keys()):
      rows.append(companyInfoDict[company][month].convertToList())

  company_ = None
  for tickerKey in companyInfoDict.keys():
      for month in sorted(companyInfoDict[company].keys()):
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
  
def calculateMarketcapAdjust(tickerKey, monthIndex):
  privuseMonthCompanyInfo:CompanyInfo = findCompanyInfo(tickerKey, allMonths[monthIndex-1])
  if privuseMonthCompanyInfo is None or privuseMonthCompanyInfo.marketcap is None:
    return None
  nextMonthCompanyInfo:CompanyInfo = findCompanyInfo(tickerKey, allMonths[monthIndex+1])

  if nextMonthCompanyInfo is not None and nextMonthCompanyInfo.marketcap is not None:
    return ((nextMonthCompanyInfo.marketcap / privuseMonthCompanyInfo.marketcap) ** (1/2)) * privuseMonthCompanyInfo.marketcap

  if monthIndex + 2 >= len(allMonths):
    return None
  secondNextMonthCompanyInfo:CompanyInfo = findCompanyInfo(tickerKey, allMonths[monthIndex+2])
  if secondNextMonthCompanyInfo is None or secondNextMonthCompanyInfo.marketcap is None:
    return None
  return ((secondNextMonthCompanyInfo.marketcap / privuseMonthCompanyInfo.marketcap) ** (1/3)) * privuseMonthCompanyInfo.marketcap
  
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
        continue
      adjustedCompanyInfo.marketcap = calculateMarketcapAdjust(tickerKey,monthIndex)
      if adjustedCompanyInfo.marketcap is None:
        print("***\tNo market cap while having closed adjust for TickerKey:" +  str(adjustedCompanyInfo.tickerKey) + " Month:" + str(adjustedCompanyInfo.month))
        continue
      if privuseMonthCompanyInfo.numberOfOutstandingShares == None:
        print("None numberOfOutstandingShares" ,privuseMonthCompanyInfo)
        continue
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
maximumCurrentYield = 0.30
minimumCurrentYield = -0.30
for tickerKey in companyTickerSet:
  for monthIndex in range(1,len(allMonths)):
    lastMonthData:CompanyInfo = findCompanyInfo(tickerKey,allMonths[monthIndex-1])
    if lastMonthData is None:
      continue
    thisMonthData:CompanyInfo = findCompanyInfo(tickerKey,allMonths[monthIndex])
    if thisMonthData is None:
      continue
    thisMonthData.currentYield = max(minimumCurrentYield,min(thisMonthData.closeAdjust / lastMonthData.closeAdjust - 1,maximumCurrentYield))

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

def valueWeightedAverage(dataList):
    if len(dataList)==0:
        return None
    lastMonth = allMonths[allMonths.index(dataList[0].month) - 1]
    data = [info.currentYield * findCompanyInfo(info.tickerKey,lastMonth).marketcap for info in dataList]
    marketcapSum = sum([findCompanyInfo(info.tickerKey,lastMonth).marketcap for info in dataList])
    if None in data:
      return None
    # print(data)
    return sum(data) / marketcapSum

# print(companyInfoList)
rollingAverageWindowSize = 6     #Can be change
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
for index,month in enumerate(allMonths[:-1]):
    companiesInMonth = [x for x in getCompaniesInMonth(month) if x.rollingAverage is not None and x.marketcap is not None]
    companiesInMonth = [x for x in companiesInMonth if findCompanyInfo(x.tickerKey,allMonths[index+1]) is not None and findCompanyInfo(x.tickerKey,allMonths[index+1]).currentYield is not None]
    companiesInMonth.sort(key=operator.attrgetter("marketcap"), reverse=True)
    median = math.ceil(len(companiesInMonth) / 2)
    for company in companiesInMonth[0:median]:
        company.sizeMOM = "Big"
    for company in companiesInMonth[median:]:
        company.sizeMOM = "Small"

print ([x for x in getCompaniesInMonth(allMonths[-1]) if x.sizeMOM ][-10:])

# %%
import math
for index,month in enumerate(allMonths[:-1]):
    companiesInMonth = [x for x in getCompaniesInMonth(month) if x.marketcap is not None]
    companiesInMonth.sort(key=operator.attrgetter("marketcap"), reverse=True)
    median = math.ceil(len(companiesInMonth) / 2)
    for company in companiesInMonth[0:median]:
        company.sizeLabel = "Big"
    for company in companiesInMonth[median:]:
        company.sizeLabel = "Small"

print ([x for x in getCompaniesInMonth(allMonths[-1]) if x.sizeLabel ][-10:])

# %%

    
    
for index,month in enumerate(allMonths[:-1]):
    companiesInMonth = [x for x in getCompaniesInMonth(month) if x.rollingAverage is not None]# and x.marketcap is not None
    # companiesInMonth = [x for x in companiesInMonth if findCompanyInfo(x.tickerKey,allMonths[index+1]) is not None and findCompanyInfo(x.tickerKey,allMonths[index+1]).currentYield is not None]
    companiesInMonth.sort(key=operator.attrgetter("rollingAverage"), reverse=True)
    count = math.floor(len(companiesInMonth) * 0.3) #It might be better to round it up instead of down to increase valid data
    for company in companiesInMonth[0:count]:
        company.momentum = "High"
    for company in companiesInMonth[len(companiesInMonth) - count:]:
        company.momentum = "Low"
    for company in companiesInMonth[count:len(companiesInMonth) - count]:
        company.momentum = "Notural"

print ([x for x in getCompaniesInMonth(allMonths[-2]) if x.momentum is not None ][-10:])

# %%
def avgTriple(first, second, third):
    avg = 0
    count = 0
    if first is not None:
        count += 1
        avg += first
    if second is not None:
        count += 1
        avg += second
    if third is not None:
        count += 1
        avg += third
    if count != 0:
        avg /= count
    return avg

# %% [markdown]
# # MomentomFactor factor (MOM)

# %%


MoMFactor = {}
MoMAverages = {}

for index,month in enumerate(allMonths[:-1]):
    # Get companies of corrent month
    companiesInMonth = getCompaniesInMonth(month)
    # Get companis for these filters SL SH BL BH
    SL = [company for company in companiesInMonth if company.sizeLabel == "Small" and company.momentum == "Low"]
    SN = [company for company in companiesInMonth if company.sizeLabel == "Small" and company.momentum == "Notural"]
    SH = [company for company in companiesInMonth if company.sizeLabel == "Small" and company.momentum == "High"]
    BL = [company for company in companiesInMonth if company.sizeLabel == "Big" and company.momentum == "Low"]
    BN = [company for company in companiesInMonth if company.sizeLabel == "Big" and company.momentum == "Notural"]
    BH = [company for company in companiesInMonth if company.sizeLabel == "Big" and company.momentum == "High"]
    # Get companies next month info
    BH_nextMonth = [findCompanyInfo(company.tickerKey,allMonths[index+1]) for company in BH]
    BN_nextMonth = [findCompanyInfo(company.tickerKey,allMonths[index+1]) for company in BN]
    BL_nextMonth = [findCompanyInfo(company.tickerKey,allMonths[index+1]) for company in BL]
    SH_nextMonth = [findCompanyInfo(company.tickerKey,allMonths[index+1]) for company in SH]
    SN_nextMonth = [findCompanyInfo(company.tickerKey,allMonths[index+1]) for company in SN]
    SL_nextMonth = [findCompanyInfo(company.tickerKey,allMonths[index+1]) for company in SL]
    # Get companies next month currentYield
    BH_average = valueWeightedAverage([x for x in BH_nextMonth if x is not None])
    BN_average = valueWeightedAverage([x for x in BN_nextMonth if x is not None])
    BL_average = valueWeightedAverage([x for x in BL_nextMonth if x is not None])
    SH_average = valueWeightedAverage([x for x in SH_nextMonth if x is not None])
    SN_average = valueWeightedAverage([x for x in SN_nextMonth if x is not None])
    SL_average = valueWeightedAverage([x for x in SL_nextMonth if x is not None])
    
    MoMAverages[month] = [SH_average,SN_average,SL_average,BH_average,BN_average,BL_average]
    MoMFactor[month]=avgTriple(SH_average, BH_average,None) - avgTriple(SL_average, BL_average,None)

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
  print(companyInfoDict[1][month])

# %%
# Accounting data fixed for 12 months
for companyTicker in companyTickerSet:
  lastB2MOfTheYear = None
  lastINVOfTheYear = None
  lastOpOfTheYear = None
  for month in allMonths:
    if month % 100 == 4:
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

for index,month in enumerate(allMonths[:-1]):
  b2mMonthList = []
  for companyTicker in companyTickerSet:
    if companyInfoDict[companyTicker][month].b2m == None :
      # companyInfoDict[companyTicker][month].marketcap == None or \
      # findCompanyInfo(companyTicker,allMonths[index+1]).currentYield == None :
      continue
    b2mMonthList.append(companyInfoDict[companyTicker][month].b2m)
  if len(b2mMonthList) == 0 :
    continue
  b2mMonthList.sort()
  lowBookToMarket = b2mMonthList[int(len(b2mMonthList)*0.3)-1]
  highBookToMarket = b2mMonthList[-int(len(b2mMonthList)*0.3)]

  for companyTicker in companyTickerSet:
    if companyInfoDict[companyTicker][month].b2m == None:
      continue
    # if companyInfoDict[companyTicker][month].currentYield == None:
    #   continue
    if companyInfoDict[companyTicker][month].b2m <= lowBookToMarket:
      companyInfoDict[companyTicker][month].B2MLabel = "Low"
    elif companyInfoDict[companyTicker][month].b2m >= highBookToMarket:
      companyInfoDict[companyTicker][month].B2MLabel = "High"
    else:
      companyInfoDict[companyTicker][month].B2MLabel = "Neutral"



# %%
# RMW (Operational profit) ranking

for index,month in enumerate(allMonths[:-1]):
  dataMonthList = []
  for companyTicker in companyTickerSet:
    if companyInfoDict[companyTicker][month].op == None :
      # companyInfoDict[companyTicker][month].marketcap == None or \
      # findCompanyInfo(companyTicker,allMonths[index+1]).currentYield == None :
      continue
    dataMonthList.append(companyInfoDict[companyTicker][month].op)
  if len(dataMonthList) == 0 :
    continue
  dataMonthList.sort()
  weakValue = dataMonthList[int(len(dataMonthList)*0.3)-1]
  robustValue = dataMonthList[-int(len(dataMonthList)*0.3)]
  for companyTicker in companyTickerSet:
    if companyInfoDict[companyTicker][month].op == None:
      continue
    # if companyInfoDict[companyTicker][month].currentYield == None:
    #   continue
    if companyInfoDict[companyTicker][month].op <= weakValue:
      companyInfoDict[companyTicker][month].RMWLabel = "Weak"
    elif companyInfoDict[companyTicker][month].op >= robustValue:
      companyInfoDict[companyTicker][month].RMWLabel = "Robust"
    else:
      companyInfoDict[companyTicker][month].RMWLabel = "Neutral"

# %%

for companyTicker in companyTickerSet[:10]:
  print(companyInfoDict[companyTicker][139807].RMWLabel)

# %%
# CMA (Investment) ranking

for index,month in enumerate(allMonths[:-1]):
  dataMonthList = []
  for companyTicker in companyTickerSet:
    if companyInfoDict[companyTicker][month].inv == None :
      # companyInfoDict[companyTicker][month].marketcap == None or \
      # findCompanyInfo(companyTicker,allMonths[index+1]).currentYield == None :
      continue   
    dataMonthList.append(companyInfoDict[companyTicker][month].inv)
  if len(dataMonthList) == 0 :
    continue
  dataMonthList.sort()
  weakValue = dataMonthList[int(len(dataMonthList)*0.3)-1]
  robustValue = dataMonthList[-int(len(dataMonthList)*0.3)]
  for companyTicker in companyTickerSet:
    if companyInfoDict[companyTicker][month].inv == None:
      continue
    # if companyInfoDict[companyTicker][month].currentYield == None:
    #   continue
    if companyInfoDict[companyTicker][month].inv <= weakValue:
      companyInfoDict[companyTicker][month].CMALabel = "Conservative"
    elif companyInfoDict[companyTicker][month].inv >= robustValue:
      companyInfoDict[companyTicker][month].CMALabel = "Aggressive"
    else:
      companyInfoDict[companyTicker][month].CMALabel = "Neutral"

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
import math
for index,month in enumerate(allMonths[:-1]):
    companiesInMonth = [x for x in getCompaniesInMonth(month) if x.b2m is not None and x.marketcap is not None]
    companiesInMonth = [x for x in companiesInMonth if findCompanyInfo(x.tickerKey,allMonths[index+1]) is not None and findCompanyInfo(x.tickerKey,allMonths[index+1]).currentYield is not None]
    companiesInMonth.sort(key=operator.attrgetter("marketcap"), reverse=True)
    median = math.ceil(len(companiesInMonth) / 2)
    for company in companiesInMonth[0:median]:
        company.sizeHML = "Big"
    for company in companiesInMonth[median:]:
        company.sizeHML = "Small"

print ([x for x in getCompaniesInMonth(allMonths[-1]) if x.sizeHML ][-10:])

# %%
HMLFactor = {}
HMLAverages = {}
SMBBTMFactor = {}
for month in allMonths[:-1]:
  SH = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].B2MLabel == "High" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Small"]
  SN = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].B2MLabel == "Neutral" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Small"]
  SL = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].B2MLabel == "Low" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Small"]
  BH = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].B2MLabel == "High" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Big"]
  BN = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].B2MLabel == "Neutral" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Big"]
  BL = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].B2MLabel == "Low" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Big"]
  
  
  SH_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(SH))
  SN_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(SN))
  SL_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(SL))
  BH_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(BH))
  BN_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(BN))
  BL_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(BL))
  
  HMLAverages[month] = [SH_average,SN_average,SL_average,BH_average,BN_average,BL_average]
  SMBBTMFactor[month] = avgTriple(SH_average,SN_average,SL_average) - avgTriple(BH_average, BN_average, BL_average)
  HMLFactor[month] = avgTriple(SH_average, BH_average,None)  - avgTriple(SL_average , BL_average,None)
for month in list(HMLFactor.keys())[:10]:
  print(HMLFactor[month])

# %%
import math
for index,month in enumerate(allMonths[:-1]):
    companiesInMonth = [x for x in getCompaniesInMonth(month) if x.op is not None and x.marketcap is not None]
    companiesInMonth = [x for x in companiesInMonth if findCompanyInfo(x.tickerKey,allMonths[index+1]) is not None and findCompanyInfo(x.tickerKey,allMonths[index+1]).currentYield is not None]
    companiesInMonth.sort(key=operator.attrgetter("marketcap"), reverse=True)
    median = math.ceil(len(companiesInMonth) / 2)
    for company in companiesInMonth[0:median]:
        company.sizeRMW = "Big"
    for company in companiesInMonth[median:]:
        company.sizeRMW = "Small"

print ([x for x in getCompaniesInMonth(allMonths[-1]) if x.sizeRMW ][-10:])

# %%
RMWFactor = {}
RMWAverages = {}
SMBRMWFactor = {}
for month in allMonths[:-1]:
  SR = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].RMWLabel == "Robust" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Small"]
  SN = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].RMWLabel == "Neutral" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Small"]
  SW = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].RMWLabel == "Weak" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Small"]
  BR = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].RMWLabel == "Robust" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Big"]
  BN = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].RMWLabel == "Neutral" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Big"]
  BW = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].RMWLabel == "Weak" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Big"]
  
  SR_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(SR))
  SN_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(SN))
  SW_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(SW))
  BR_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(BR))
  BN_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(BN))
  BW_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(BW))

  RMWAverages[month] = [SR_average,SN_average,SW_average,BR_average,BN_average,BW_average]
  SMBRMWFactor[month] = avgTriple(SR_average , SN_average, SW_average) - avgTriple(BR_average, BN_average, BW_average)
  RMWFactor[month] = avgTriple(SR_average, BR_average,None) - avgTriple(SW_average, BW_average,None)
for month in list(RMWFactor.keys())[:10]:
  print(RMWFactor[month])

# %%
import math
for index,month in enumerate(allMonths[:-1]):
    companiesInMonth = [x for x in getCompaniesInMonth(month) if x.op is not None and x.marketcap is not None]
    companiesInMonth = [x for x in companiesInMonth if findCompanyInfo(x.tickerKey,allMonths[index+1]) is not None and findCompanyInfo(x.tickerKey,allMonths[index+1]).currentYield is not None]
    companiesInMonth.sort(key=operator.attrgetter("marketcap"), reverse=True)
    median = math.ceil(len(companiesInMonth) / 2)
    for company in companiesInMonth[0:median]:
        company.sizeCMA = "Big"
    for company in companiesInMonth[median:]:
        company.sizeCMA = "Small"

print ([x for x in getCompaniesInMonth(allMonths[-1]) if x.sizeCMA ][-10:])

# %%
CMAFactor = {}
CMAAverages = {}
SMBCMAFactor = {}
for month in allMonths[:-1]:
  SA = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].CMALabel == "Aggressive" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Small"]
  SN = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].CMALabel == "Neutral" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Small"]
  SC = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].CMALabel == "Conservative" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Small"]
  BA = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].CMALabel == "Aggressive" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Big"]
  BN = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].CMALabel == "Neutral" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Big"]
  BC = [companyInfoDict[companyTicker][month] for companyTicker in companyTickerSet if companyInfoDict[companyTicker][month].CMALabel == "Conservative" and \
        companyInfoDict[companyTicker][month].sizeLabel == "Big"]
  
  SA_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(SA))
  SN_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(SN))
  SC_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(SC))
  BA_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(BA))
  BN_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(BN))
  BC_average = valueWeightedAverage(getNextMonthCompaniesHavingYeild(BC))

  CMAAverages[month] = [SA_average,SN_average,SC_average,BA_average,BN_average,BC_average]
  SMBCMAFactor[month] = avgTriple(SA_average, SN_average, SC_average) - avgTriple(BA_average, BN_average, BC_average)
  CMAFactor[month] = avgTriple(SC_average, BC_average,None)  - avgTriple(SA_average, BA_average,None)
for month in list(CMAFactor.keys())[:10]:
  print(CMAFactor[month])

# %%
SMBFactor = {}
for month in SMBBTMFactor.keys():
  SMBFactor[month] = avgTriple(SMBBTMFactor[month], SMBCMAFactor[month], SMBRMWFactor[month])


# %% [markdown]
# # Export company list to excel

# %%
# Example to excel function
exportCompanyListToExcel("calculations")

# %%
# Factors to excel
rows = []
for month in allMonths:
    data_ = [month,
      SMBFactor[month] if month in SMBFactor else None,
      HMLFactor[month] if month in HMLFactor else None]
    for i in range(6):
      data_.append(HMLAverages[month][i] if month in HMLAverages else None)
    
    data_.append(RMWFactor[month] if month in RMWFactor else None)
    for i in range(6):
      data_.append(RMWAverages[month][i] if month in RMWAverages else None)
    
    data_.append(CMAFactor[month] if month in CMAFactor else None)
    for i in range(6):
      data_.append(CMAAverages[month][i] if month in CMAAverages else None)
      
    data_.append(MoMFactor[month] if month in MoMFactor else None)
    for i in range(6):
      data_.append(MoMAverages[month][i] if month in MoMAverages else None)
      
    rows.append(data_)

dfOut = pd.DataFrame(rows, columns=['month',"SMBFactor", "HMLFactor","SH_average","SN_average","SL_average","BH_average","BN_average","BL_average", "RMWFactor","SR_average","SN_average","SW_average","BR_average","BN_average","BW_average", "CMAFactor","SA_average","SN_average","SC_average","BA_average","BN_average","BC_average", "MoMFactor","SH_average","SN_average","SL_average","BH_average","BN_average","BL_average"])
dfOut.to_excel('Factors.xlsx', sheet_name='Factors')


