# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
import pandas
from sklearn import linear_model
from sklearn.metrics import *


def saveDFToExcel(dataFrame,fileName,sheetName):
  regressionDF = pandas.DataFrame(dataFrame)
  regressionDF.to_excel(fileName + '.xlsx', sheet_name=sheetName,index=False)

df = pandas.read_excel("Factors (with market & industries).xlsx", "Factors")
offsetMonths = 36
industryCount = 50

strategyYield = 1
monthlyYieldHistory = {"month":[],"yield":[],"strategyYield":[]}

for monthIndex in range(offsetMonths,len(df)-1):
  weight = 1/ industryCount
  monthYield = 0
  for i in range(1,industryCount+1):
      monthYield += weight* df["Industry"+str(i)][monthIndex+1]
  strategyYield *= 1+monthYield
  monthlyYieldHistory["month"].append(df["month"][monthIndex+1])
  monthlyYieldHistory["yield"].append(monthYield)
  monthlyYieldHistory["strategyYield"].append(strategyYield)


print(strategyYield)

saveDFToExcel(monthlyYieldHistory,"NaiveStrategy","Monthly yields")


