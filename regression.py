# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
import pandas
from sklearn import linear_model

df = pandas.read_csv("Factors (with market).csv")
# df=df[:-1]
# df["IndustryValue"] = list(range(1, df.shape[0]+1))
x = df[["HMLFactor","RMWFactor","MoMFactor","aa"]]
y = df['IndustryExcessReturn']
regr = linear_model.LinearRegression()
regr.fit(x, y)
# aa","asdsad" "MoMFactor","HMLFactor","RMWFactor",SMBFactor ,"CMAFactor","MoMFactor","MarketExcessReturn"
# print(regr.intercept_) 
# print(regr.coef_) 


# %%
print(x.shape)
print(y.shape)


# %%
from sklearn.metrics import *
from scipy import stats
yPredict = regr.predict(x)
slope, intercept, r, p, se = stats.linregress(yPredict, y)
print(yPredict)
print(r2_score(y, yPredict))
print(r**2)


# %%
import pandas
import numpy as np
from sklearn import linear_model
from sklearn.metrics import *

x = np.array([3,8,10,17,24,27])
y = np.array([2, 8, 10, 13, 18, 20])
x = x.reshape(6, 1)
print(x.shape)
print(x)
# y.reshape(-1, 1)
regr = linear_model.LinearRegression()
regr.fit(x, y)
yPredict = regr.predict(x)

print(yPredict)

print(r2_score(y,yPredict))


# %%
import pandas
import numpy as np
from sklearn import linear_model
from sklearn.metrics import *

df = pandas.read_excel("testRegression.xlsx")
# df["IndustryValue"] = list(range(1, df.shape[0]+1))
x = df[['x']]
print(x.shape)
y = df['y']
regr = linear_model.LinearRegression()
regr.fit(x, y)
yPredict = regr.predict(x)

print(yPredict)

print(r2_score(y,yPredict))


