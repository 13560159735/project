import pandas as pd
import numpy as np
# pd.set_option('display.float_format',lambda x : '%.2f' % x)
source_file='数据.xlsx'
df=pd.read_excel(
    source_file,
    sheet_name='缺失数据',
    dtype={'凭证号':str,'科目编号':str}
)
print(df)
print(df.isnull())
print('0000000000000000000000000')
# print(df.isnull().any())
# print('1111111111111111111111111')
# print(df.isnull().sum())
print(np.sum(df.isnull(),axis=1))
print(np.sum(df.isnull(),axis=0))
