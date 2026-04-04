import pandas as pd
from datetime import datetime
import json
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import xlwings as xw
import shutil
import requests
import os
import openpyxl
from docx import Document
from collections import deque
from docx.shared import Pt
from docx.oxml.ns import qn
from pyecharts.charts import Bar
from pyecharts import options as opts

data='2026-01-11'  #修改日期即可
position='兆丰壹号'
dateTimestamp=pd.Timestamp(data)

dfReadMain=pd.read_excel(
    '1.xlsx',
    sheet_name=position,
    header=[0]
)

appMain=xw.App(visible=False)
wbMain=appMain.books.open('1.xlsx')
wsMain=wbMain.sheets[0]

mainColumn={
    '合同出租面积':3,
    '合同起租日':10,
    '合同到期日':11
}

hasRentArea=0
notHasRentArea=0

mainList=dfReadMain.values.tolist()

for mainIndex,mainItem in enumerate(mainList):
    if(not pd.isna(mainItem[mainColumn['合同出租面积']])):
        if(not pd.isna(mainItem[mainColumn['合同起租日']])) and (not pd.isna(mainItem[mainColumn['合同到期日']])) and (mainItem[mainColumn['合同起租日']]<=dateTimestamp) and (dateTimestamp<=mainItem[mainColumn['合同到期日']]):
            print('正在租',mainIndex,mainItem)
            hasRentArea=hasRentArea+mainItem[mainColumn['合同出租面积']]
        else:
            print('未在租',mainIndex,mainItem)
            notHasRentArea=notHasRentArea+mainItem[mainColumn['合同出租面积']]
    else:
        print('这一行面积为空',mainIndex,mainItem)

allRentArea=hasRentArea+notHasRentArea
allRentArea=round(allRentArea,2)
hasRentArea=round(hasRentArea,2)
notHasRentArea=round(notHasRentArea,2)
rate=round(hasRentArea/allRentArea*100,2)
print(allRentArea,hasRentArea,notHasRentArea,rate)

title=f"{data},{position}可出租面积总数为{allRentArea},已出租面积为{hasRentArea},未出租面积为{notHasRentArea},出租率为{rate}%"
bar=Bar(init_opts=opts.InitOpts(width="100%",height="calc(100vh - 40px)"))
bar.add_xaxis(['可出租面积总数','已出租面积','未出租面积'])
bar.add_yaxis(data,[allRentArea,hasRentArea,notHasRentArea])
bar.set_global_opts(
    title_opts=opts.TitleOpts(
        title=title
    ),
    tooltip_opts=opts.TooltipOpts(trigger='axis'),
    legend_opts=opts.LegendOpts(),
    xaxis_opts=opts.AxisOpts(type_='category'),
    yaxis_opts=opts.AxisOpts(type_='value')
)
bar.render(f"{title}.html")

wsMain[7,1].value=allRentArea
wsMain[7,3].value=notHasRentArea
wsMain[7,5].value=hasRentArea
wsMain[7,7].value=f"{rate}%"

wbMain.save('1.xlsx')
wbMain.close()
appMain.quit()