import pandas as pd
# pd.set_option('display.float_format',lambda x : '%.2f' % x)
source_file='数据.xlsx'
# df=pd.read_excel(
#     source_file,
#     sheet_name=None
    # dtype={'凭证号':str,'科目编号':str}
# )
df=pd.read_excel(
    source_file,
    sheet_name='序时账',
    # header=0,                                        #取第一行第数据作为列索引
    # index_col=0,                                     #取第一列第数据作为行索引
    # usecols=['日期','凭证号','科目编号','科目名称'],   #只取这四列数据
    # dtype={'凭证号':str,'科目编号':str},              #将数据类型修改为字符串
    # skipfooter=4                                     #跳过最后4行，从倒数第5行开始向上取数据
)
# print(df)
# print(df.dtypes)
# print(df.head(3))
# print(df.tail(3))

# print(df.index) # 查看行索引
# print(df.columns) # 查看列索引
# print(df.shape) # 查看行数与列数
# print(df.size) # 查看元素数，即行列数乘积
# print(df.describe())
# print(df.describe(include='O'))
print(df)
print('000000000000000000000000')
# df.set_index('日期',inplace=True)
# df.set_index('凭证号',inplace=True,append=True)
df.set_index(['日期','凭证号'],inplace=True)
print(df)
print('1111111111111111111111111111111')
# df.reset_index(level=['日期','凭证号'],inplace=True)
# print(df)
# print('222222222222222222222222')
df.reset_index(level='凭证号',inplace=True)
print(df)
print('222222222222222222222222')
df.reset_index(level='日期',inplace=True)
print(df)

