import pandas as pd
source_file='兆丰序20250930.xls'
df=pd.read_excel(
    source_file,
    sheet_name='123',
    usecols=['日期','凭证字号','科目代码','科目名称','借方金额','贷方金额']
    )
print(df)
df.isnull().sum()
print(df)
# df.dtypes
# print(df.dtypes)
# df.set_index(['日期','凭证字号'])
# print(df)
print("123")