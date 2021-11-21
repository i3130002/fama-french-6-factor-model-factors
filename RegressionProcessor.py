# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
import pandas
from sklearn import linear_model
from sklearn.metrics import *

axis = ["SMBFactor", "HMLFactor", "RMWFactor",
        "CMAFactor", "MoMFactor", "Market"]
df = pandas.read_excel("Factors (with market & industries).xlsx", "Factors")

def calculateRegression(x, y):
    regr = linear_model.LinearRegression()
    regr.fit(x, y)
    yPredict = regr.predict(x)
    return regr, yPredict


def prepareRegressionDF():
    regressionDF = {
        "month": []
    }
    for i in axis:
        regressionDF[i] = []
    regressionDF["IndustryYield"]= []
    regressionDF["R2"]= []
    regressionDF["intercept"]= []
    for i in axis:
        regressionDF["coefficient_" + i] = []
    return regressionDF


def fillRegressionDF(month,IndustryYield,factors,regressionDF,r2, intercept, coef_):
    regressionDF["month"].append(month)
    regressionDF["IndustryYield"].append(IndustryYield)
    for i in range(len(factors)):
      regressionDF[axis[i]].append(factors[i])
    regressionDF["R2"].append(r2)
    regressionDF["intercept"].append(intercept)
    for i in axis:
        regressionDF["coefficient_" +
                     i].append(coef_[axis.index(i)])


def produceRegressionDF(x,y):
    regressionDF = prepareRegressionDF()
    for i in range(len(y)):
        factors = []
        for factor in axis:
          factors.append(df[factor][i])
        if i < rollingSize-1:
            fillRegressionDF(df["month"][i],y.iloc[i],factors,regressionDF,None, None, [None]*6)
            continue

        start = i - rollingSize + 1
        end = i
        regr, yPredict = calculateRegression(x[start:end], y[start:end])
        fillRegressionDF(df["month"][i],y.iloc[i],factors,regressionDF,
            r2_score(y[start:end], yPredict), regr.intercept_, regr.coef_)
    return regressionDF


def saveDFToExcel(dataFrame,fileName,sheetName):
  regressionDF = pandas.DataFrame(dataFrame)
  regressionDF.to_excel(fileName + '.xlsx', sheet_name=sheetName,index=False)

def saveRegressionDF(isDisplay,regressionDF):
    regressionDF = pandas.DataFrame(regressionDF)
    regressionDF.to_excel('regressionOutput.xlsx', sheet_name='regression')
    if isDisplay:
        print(regressionDF[rollingSize-1-1:rollingSize-1+4])




rollingSize = 36
x = df[axis]
# y = df['IndustryExcessReturn']
industryCount = 50
industryColumns = ["Industry" + str(x) for x in range(1,industryCount+1)]

# regressionDF = produceRegressionDF(x,y)
# saveRegressionDF(True,regressionDF)


# %%
print(y.iloc[1])


# %%

industryCount = 50
industryColumns = ["Industry" + str(x) for x in range(1,industryCount+1)]
print(industryCount)
print(industryColumns)
regressionOutput = {}
for i in industryColumns:
  y = df[i]
  regressionOutput[i] = produceRegressionDF(x,y)



# %%
import time
writer = pandas.ExcelWriter("regressionOutput.xlsx", engine = 'xlsxwriter')
for i in industryColumns:
  regressionDF = pandas.DataFrame(regressionOutput[i])
  regressionDF.to_excel(writer, sheet_name = i)
  
writer.save()
time.sleep(3)
writer.close()


# %%
acceptedIndustriesList = []
for monthIndex in range(rollingSize, len(df)):
  allIntercepts = []
  for i in industryColumns:
    allIntercepts.append(regressionOutput[i]["intercept"][monthIndex])
  allIntercepts.sort()
  lastAcceptable = allIntercepts[int(industryCount * -0.1)]
  acceptedIndustries = []
  for i in industryColumns:
    if regressionOutput[i]["intercept"][monthIndex]>= lastAcceptable and regressionOutput[i]["intercept"][monthIndex]> 0:
      acceptedIndustries.append(i)
  acceptedIndustriesList.append(acceptedIndustries)

print(acceptedIndustriesList[:2])


# %%
strategyYield = 1
monthlyYieldHistory = {"month":[],"yield":[],"strategyYield":[]}
for monthIndex in range(rollingSize, len(df)-1):
  weight = 1/ int(industryCount * 0.1)
  monthYield = 0
  for acceptedIndustry in acceptedIndustriesList[monthIndex-rollingSize]:
      monthYield += weight* regressionOutput[acceptedIndustry]["IndustryYield"][monthIndex+1]
  strategyYield *= 1+monthYield
  monthlyYieldHistory["month"].append(df["month"][monthIndex+1])
  monthlyYieldHistory["yield"].append(monthYield)
  monthlyYieldHistory["strategyYield"].append(strategyYield)

print(strategyYield)

saveDFToExcel(monthlyYieldHistory,"SectorRotation6AxisNaive","Monthly yields")


