# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
import pandas
import math
from sklearn import linear_model
from sklearn.metrics import *


def saveDFToExcel(dataFrame,fileName,sheetName):
  regressionDF = pandas.DataFrame(dataFrame)
  regressionDF.to_excel(fileName + '.xlsx', sheet_name=sheetName,index=False)

df = pandas.read_excel("Factors (with market & industries).xlsx", "Factors")
rollingSize = 6
offsetMonths = 36 
industryCount = 50

for i in range(1,industryCount+1):
  rollingAverage = []
  for monthIndex in range(0,offsetMonths):
    rollingAverage.append(None)
  for monthIndex in range(offsetMonths,len(df)-1):
    rollingWindow = df["Industry"+str(i)][monthIndex-rollingSize+1:monthIndex+1]
    avg = sum(rollingWindow) / rollingSize
    rollingAverage.append(avg)
  rollingAverage.append(None)
  df["rollingAverage"+str(i)] = rollingAverage
  
saveDFToExcel(df,"IndustryMomentumRollingAverages","RollingAverages")


# %%
acceptedIndustriesList = []
for monthIndex in range(offsetMonths, len(df)):
  allRollings = []
  for i in range(1,industryCount+1):
    allRollings.append(df["rollingAverage"+str(i)][monthIndex])
  allRollings.sort()
  lastAcceptable = allRollings[int(industryCount * -0.1)]
  acceptedIndustries = []
  for i in range(1,industryCount+1):
    if df["rollingAverage"+str(i)][monthIndex]>= lastAcceptable:
      acceptedIndustries.append(i)
  acceptedIndustriesList.append(acceptedIndustries)

print(acceptedIndustriesList[:20])


# %%
strategyYield = 1
monthlyYieldHistory = {"month":[],"yield":[],"strategyYield":[]}
for monthIndex in range(offsetMonths, len(df)-1):
  weight = 1/ int(industryCount * 0.1)
  monthYield = 0
  for acceptedIndustry in acceptedIndustriesList[monthIndex-offsetMonths]:
      monthYield += weight* df["Industry"+str(acceptedIndustry)][monthIndex+1]
  strategyYield *= 1+monthYield
  monthlyYieldHistory["month"].append(df["month"][monthIndex+1])
  monthlyYieldHistory["yield"].append(monthYield)
  monthlyYieldHistory["strategyYield"].append(strategyYield)

print(strategyYield)

saveDFToExcel(monthlyYieldHistory,"IndustryMomentumStrategy","Monthly yields")


