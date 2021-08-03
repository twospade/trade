from numpy.core.fromnumeric import transpose
import pandas as pd
import warnings
import numpy as np
import xlwings as xw
from pandas.core.frame import DataFrame
warnings.filterwarnings("ignore")

# coupon_rate = 'Coupon Rate'
coupon_rate = 'COUPON RATE'
# maturity ='Maturity'
maturity = 'MATURITY'
# file = r'D:/vscode/trade/Trades_Summary_5_9Jul.xlsx'
file = r'D:/vscode/trade/trade_weekly.xlsx'
# file = r'D:/vscode/trade/Trades Summary 12-16 July.xlsx'
# sheet = 'CEBIC'
sheet1 = '1. STD TRADE SUMMARY'
sheet2 = '2. STD POSITION CHANGE'
# sheet = 'CEBIC-UPDATED'
xls = pd.ExcelFile(file)
df1 = pd.read_excel(xls, sheet1)
df1 = df1.loc[df1['Counterparty'].str.contains('CEBI') == False]
df1['Trader'] = df1['Trader'].apply(lambda x: x.strip())

df2 = pd.read_excel(xls, sheet2)
df2[['BondIssuer', 'Other']] = df2['BondName'].str.split(n=1).apply(pd.Series)

pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)

# pd.set_option('display.float_format',  '{:,.0f}'.format)


def check_section():
    # soe_df = df2.loc[df2['Section'] == '国企']
    sum_df = df2.groupby(['Section']).sum()
    sum_df['%'] = sum_df.Balance/sum_df.sum().Balance
    sum_df['Change%'] = sum_df.Change/sum_df.Balance
    res = sum_df[['Balance', '%', 'Change', 'Change%']]
    # res['%'] = res['%'].map('{:.2f}%'.format)
    res['Change'] = res['Change'].map('{:,.0f}'.format)
    # res['Change%'] = res['Change%'].map('{:.2f}%'.format)
    return res


def check_value():


    issuer_group = df2.groupby('BondIssuer').agg(
        {'Balance': 'sum', 'Change': 'sum'})


    issuer_group['Long'] = issuer_group['Balance'].apply('lambda x:  ')

    issuer_group['%Balance'] = issuer_group['Balance'] / \
        issuer_group.sum()['Balance']
    issuer_group['%Change'] = issuer_group.Change/issuer_group.sum().Change
    # issuer_group.style.format({'%': "{:.2f}", 'Balance': "{:,.0f}"})
    issuer_df = df2[['Issuer', 'BondIssuer', 'Balance']].groupby(
        ['Issuer', 'BondIssuer']).count()
    issuer_list = []
    for i in issuer_df.index.to_list():
        if i and not '#' in i[0]:
            issuer_list.append(i)
    issuer_dict = dict((y, x) for x, y in issuer_list)
    output = issuer_group.sort_values('Balance', ascending=[False])
    # output.head(10).style.format({'%': "{:.2f}", 'Balance': "{:,.0f}"})
    output['Issuer_c'] = output.index.map(issuer_dict)
    return output[['Issuer_c', 'Balance', '%Balance', 'Change', '%Change']]


def f(x):
    d = {}
    d['Price Mean'] = x['Execution Price'].mean()
    d['Price Std'] = std(x['Execution Price'])
    d['Count'] = x['Amount'].count()
    d['Amount Sum'] = (x['Amount']).sum()
    d['Amount Mean'] = (x['Amount']).mean()
    d['Amount Std'] = std(x['Amount'])
    return pd.Series(d, index=['Price Mean', 'Price Std', 'Amount Sum', 'Amount Std', 'Amount Mean', 'Count'])
def b(x):
    d = {}
    d['Balance Sum'] = x['Balance'].sum()
    d['Change Sum'] = x['Change'].sum()
    d['Count'] = x['Balance'].count()


def std(n):
    return np.std(n)


wb = xw.Book('./position_template.xlsx')

sheet_position = wb.sheets['Position']
sheet_counterparty = wb.sheets['CounterParty']
sheet_issuer = wb.sheets['Issuer']
sheet_section = wb.sheets['Section']
sheet_trader = wb.sheets['Trader']
sheet_ighy = wb.sheets['Others']

res = check_value()
sheet_position['A3'].value = res[['Issuer_c', 'Balance', '%']].head(10)

res = check_section()
sheet_position['A31'].value = res.head(10)
sheet_section['A10'].value = res.head(10)

df1['Amount'] = df1['Execution Price']*df1['Execution Qty']/100000

# sum_df = df1.groupby(['Counterparty']).count().sort_values(
#     'Amount', ascending=[False])['Amount']
# print(sum_df.head(10))
# sheet_counterparty['A1'].value= sum_df.head(10)
# # wb.sheets[0]['A40'].value = sum_df.head(10)

# sum_df = df1.groupby(['Counterparty']).sum().sort_values(
#     'Amount', ascending=[False])['Amount']
# print(sum_df.head(10))
# sheet_counterparty['A13'].value= sum_df.head(10)

# sum_df = df1.groupby(['Counterparty']).mean().sort_values(
#     'Amount', ascending=[False])['Amount']
# print(sum_df.head(10))
# sheet_counterparty['A26'].value= sum_df.head(10)

sum_df = df1.groupby(['Counterparty']).apply(
    f).sort_values('Amount Sum', ascending=[False])
sum_df['Amount%'] = sum_df['Amount Sum']/sum_df['Amount Sum'].sum()
print(sum_df.head(20))
sheet_counterparty['A1'].value = sum_df.head(20)

sum_df = df1.groupby(['Trader']).apply(f)
sum_df['Amount%'] = sum_df['Amount Sum']/sum_df['Amount Sum'].sum()
# sum_df['%'] = sum_df['Amount']/sum_df['Amount'].sum()
sum_df.insert(0, '/', '')
print(sum_df.head(10))
sheet_trader['A1'].value = sum_df.head(10)


sum_df = df1.groupby(['Trader', 'IG/HY']).apply(f)
sum_df['Amount%'] = sum_df['Amount Sum']/sum_df['Amount Sum'].sum()
print(sum_df)
sheet_trader['A10'].value = sum_df

sum_df = df1.groupby(['Trader', 'B/S']).apply(f)
sum_df['Amount%'] = sum_df['Amount Sum']/sum_df['Amount Sum'].sum()
print(sum_df)
sheet_trader['A20'].value = sum_df


sum_df = df1.groupby(['NAME_CHINESE_SIMPLIFIED']).apply(
    f).sort_values('Amount Sum', ascending=False)
sum_df['Amount%'] = sum_df['Amount Sum']/sum_df['Amount Sum'].sum()
# sum_df = df1.groupby(['NAME_CHINESE_SIMPLIFIED']).sum(
# ).sort_values('Amount', ascending=[False])
print(sum_df)
sheet_issuer['A1'].value = sum_df

df1[coupon_rate] = df1[coupon_rate].apply(lambda x: round(x))
# sum_df = df1.groupby([coupon_rate]).sum(
# ).sort_values('Amount', ascending=[False])['Amount']
sum_df = df1.groupby([coupon_rate]).sum(
).sort_values('Amount', ascending=[False])
sum_df['%'] = sum_df['Amount']/sum_df.sum().Amount
print(sum_df)
sheet_ighy['A1'].value = sum_df[['Amount', '%']]

sum_df = df1.groupby(['IG/HY']).sum(
).sort_values('Amount', ascending=[False])
sum_df['%'] = sum_df['Amount']/sum_df.sum().Amount
print(sum_df[['Amount', '%']])
sheet_ighy['A20'].value = sum_df[['Amount', '%']]

sum_df = df1.groupby(['B/S']).sum().sort_values('Amount',
                                                ascending=[False])
sum_df['%'] = sum_df['Amount']/sum_df['Amount'].sum()
print(sum_df.head(10))
sheet_ighy['A30'].value = sum_df[['Amount', '%']].head(10)

df1['Perpetual'] = df1[maturity].apply(lambda x: 'Perpetual' in x)
sum_df = df1.groupby(['Perpetual']).apply(f)
sum_df['%'] = sum_df['Amount Sum']/sum_df['Amount Sum'].sum()
print(sum_df)
sheet_ighy['A40'].value = sum_df[['Amount Sum', '%']]

sum_df = df1.groupby(['板块']).apply(f)
sum_df['Amount%'] = sum_df['Amount Sum']/sum_df['Amount Sum'].sum()
# sum_df.loc['total'] = sum_df.select_dtypes(pd.np.number).sum()
# sum_df.loc['total',['Amount Sum','%','Count']] = sum_df.sum()
# sum_df.loc['mean',['Price Mean']] = sum_df.mean()
# sum_df.loc['std',['Price Mean','Amount Sum','Count']] = sum_df.agg(np.std)
print(sum_df)
sheet_section['A1'].value = sum_df
wb.save()
wb.save('myreport.xlsx')
