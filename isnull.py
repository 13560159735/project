import pandas as pd
import numpy as np
# pd.set_option('display.float_format',lambda x : '%.2f' % x)
source_file='数据.xlsx'
df=pd.read_excel(
    source_file,
    sheet_name='缺失数据',
    dtype={'凭证号':str,'科目编号':str}
)
# print(df)
# print(df.isnull())
# print('0000000000000000000000000')
# print(df.isnull().any())
# print('1111111111111111111111111')
# print(df.isnull().sum())
# print(np.sum(df.isnull(),axis=1))
# print(np.sum(df.isnull(),axis=0))
# print(df)

# df ,这种数据结构我们没学过，而且是这个库自创的数据结构
# 需要喂给for的东西一般是列表
# for用
# newList=df.values.tolist()
# print(newList)
# sum=0
# for i in newList:
#     print(i)
#     for j in i:
#         # print(pd.isna(j))
#         if pd.isna(j):
#             sum=sum+1

# print(sum)
print(df)
print('11111111111111')
df1=df.dropna(how='all')
print(df1)



