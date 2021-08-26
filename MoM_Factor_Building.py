import pandas as pd
import operator

df_MoM = pd.read_excel(r'Momentum.xlsx', sheet_name='Yield')
df_MarketCap = pd.read_excel(r'Momentum.xlsx', sheet_name='MarketCap')

# print(df_MoM)
# print(df_MarketCap)

Window_Size = 12
Section_Count = 5

RollingAverage_Frame = {}

for key in list(df_MoM.keys())[1:]:
    Corporation = df_MoM[key]
    RollingAverage_Frame[key] = []
    for i in range(len(Corporation) - Window_Size + 1):
        RollingAverage_Frame[key].append(sum(Corporation[i: i + Window_Size]) / Window_Size)


# print(RollingAverage_Frame)


class Company:
    def __init__(self, name, month, rolling_average, market_cap, next_month_yield):
        self.name = name
        self.month = month
        self.rolling_average = rolling_average
        self.market_cap = market_cap
        self.next_month_yield = next_month_yield


Company_Frame = []
# اونایی که ندارنو بعدا میذاریم -99
# و توی حقله دوم اونهارو اضافه نمیکنیم
for i in range(len(df_MoM['a']) - Window_Size):

    Companies_At_Month = []

    for key in list(df_MoM.keys())[1:]:
        Companies_At_Month.append(
            Company(key, Window_Size + i, RollingAverage_Frame[key][i], df_MarketCap[key][Window_Size + i - 1],
                    df_MoM[key][Window_Size + i]))

    Company_Frame.append(Companies_At_Month)

# Set size
for Date in Company_Frame:
    Date.sort(key=operator.attrgetter("market_cap"), reverse=True)
    for company in Date[0:len(Date) // 2]:
        company.size = "Big"
    for company in Date[len(Date) // 2:]:
        company.size = "Small"

# Set momentum

for Date in Company_Frame:
    Date.sort(key=operator.attrgetter("rolling_average"), reverse=True)
    for company in Date[0:len(Date) // 2]:
        company.momentum = "High"
    for company in Date[len(Date) // 2:]:
        company.momentum = "Low"


# print(Company_Frame)


def average(lst):
    return sum(lst) / len(lst)


MoM = []

for Date in Company_Frame:
    small = [company for company in Date if company.size == "Small"]
    SL = [company for company in small if company.momentum == "Low"]
    SH = [company for company in small if company.momentum == "High"]
    big = [company for company in Date if company.size == "Big"]
    BL = [company for company in big if company.momentum == "Low"]
    BH = [company for company in big if company.momentum == "High"]

    BH_average = average([company.next_month_yield for company in BH])
    BL_average = average([company.next_month_yield for company in BL])
    SH_average = average([company.next_month_yield for company in SH])
    SL_average = average([company.next_month_yield for company in SL])

    # print(BH_average)

    MoM.append(0.5 * (SH_average + BH_average) - 0.5 * (SL_average + BL_average))

print(MoM)
