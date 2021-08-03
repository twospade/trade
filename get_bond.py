import xlwings as xw
import pandas as pd
from datetime import datetime


wb = xw.Book('./note.xlsx')
# sheet1 = '地产'
# sheet1 = '金融机构'
# sheet1 = '央企'
sheet1 = '债券信息'
df = pd.read_excel(r'./债券数据_20210727.xlsx', sheet1)
df['年度'] = df['发行日'].apply(lambda x: x.year)
col = '省'
# col = '市'
# col = '发行日'
# col = '中文名'
item = '中文名'
# df1 = df.groupby([col,item]).count()
# df1.sort_values([col,'ISIN'],ascending=[True,False]).head(50)
# idx = df1.groupby(col)['ISIN'].transform(max) == df1['ISIN']
# df1[idx]
# wb.sheets[0]['A1'].value = df1[idx].sort_values('ISIN',ascending=False)

df2 = df.groupby([col, '年度']).count()

wb.sheets[0]['A1'].value = df2.sort_values(
    [col, '年度'], ascending=[True, False])
wb.save()
