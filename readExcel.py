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
    header=0,
    index_col=0,
    usecols=['日期','凭证号','科目编号','科目名称'],
    dtype={'凭证号':str,'科目编号':str},
    skipfooter=4
)
print(df)