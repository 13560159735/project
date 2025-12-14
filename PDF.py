import pdfplumber
import pandas as pd


import pdfplumber
import pandas as pd

p=pdfplumber.open('./22222222.pdf')
# print(p.pages[0])

for i in range(len(p.pages)):
    talbes=p.pages[i].extract_tables()
    print(f'第{i+1}页一共有{len(talbes)}个表格')
    for j in range(len(talbes)):
        df=pd.DataFrame(talbes[j])
        df.to_excel(f'./22222222的pdf转excel.xlsx')

