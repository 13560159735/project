import pandas as pd
from datetime import datetime
import json
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import xlwings as xw
import shutil

dfReadLeft=pd.read_excel(
    '123.xlsx',
    sheet_name='南京-2025-销售',
    header=[0,1,2,3]
)

dfReadLeftList=dfReadLeft.values.tolist()

dfReadRight=pd.read_excel(
    '123.xlsx',
    sheet_name='25年收入凭证',
    header=[0,1,2,3]
)

dfReadRightList=dfReadRight.values.tolist()

typeToColumnOfLeft={
    '单据编号':3,
    '金额':19,
    '凭证唯一码(未开票)':39,
    '凭证合并(未开票)':40,
    '凭证金额(未开票)':41,
    '凭证唯一码(已开票)':44,
    '凭证合并(已开票)':45,
    '凭证金额(已开票)':46,
}
leftResult={}
for index,item in enumerate(dfReadLeftList):
    currentOrderNumber=item[typeToColumnOfLeft['单据编号']]
    if currentOrderNumber not in leftResult:
        leftResult[currentOrderNumber]={
            'amount':0,
            'isna':True,
            'indexAll':[]
        }
    if (not pd.isna(item[typeToColumnOfLeft['金额']])):
        leftResult[currentOrderNumber]['amount']=leftResult[currentOrderNumber]['amount']+item[typeToColumnOfLeft['金额']]
        leftResult[currentOrderNumber]['isna']=False
    leftResult[currentOrderNumber]['indexAll'].append(index)

with open(f'leftResult.json','w',encoding='utf-8') as f:
    json.dump(leftResult,f,ensure_ascii=False, indent=4)

def find_first_containing(lst,substring='XSCKD'):
    for item in lst:
        if substring in str(item):    #包含XSCKD的单号用  in
            return item
    return ''
typeToColumnOfRight={
    '摘要':3,
    '贷方金额':8,
    '科目全名':5,
    '凭证唯一码':0,
    '凭证合并':2,
    '凭证金额':8,
}
rightResult={}
for index,item in enumerate(dfReadRightList):
    splitArray=str(item[typeToColumnOfRight['摘要']]).split('/')
    currentOrderNumber=find_first_containing(splitArray)
    if currentOrderNumber !='':
        if currentOrderNumber not in rightResult:
            rightResult[currentOrderNumber]={
                'amount':0,
                'isna':True,
                'indexAll':[]
            }
        if (not pd.isna(item[typeToColumnOfRight['贷方金额']])):
            rightResult[currentOrderNumber]['amount']=rightResult[currentOrderNumber]['amount']+item[typeToColumnOfRight['贷方金额']]
            rightResult[currentOrderNumber]['isna']=False
        rightResult[currentOrderNumber]['indexAll'].append(index)

with open(f'rightResult.json','w',encoding='utf-8') as f:
    json.dump(rightResult,f,ensure_ascii=False, indent=4)