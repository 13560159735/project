import pandas as pd
filePath='员工信息.xlsx'
df=pd.read_excel(filePath)
print(df)
df.to_excel('./员工信息1.xlsx',index=False)
df.to_excel('./员工信息2.xlsx')