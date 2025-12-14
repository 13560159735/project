# a = 1000
# b=200
# print(a+b)
# print(a)
# print(b)
# c='姓名'
# print(c)
# d='2025-10-25 21:08:01'
# print(d)
# class1=1
# print(type(class1))
# my_str = "应收账款的期末余额为2000元" # 定义一个字符串

# print(my_str.replace('应收账款','货币资金'))  # 把'应收账款'替换为'货币资金'
# my_list = [100, 200, 300, 400, 500] # 定义一个列表对象my_list
# print(my_list[0]) #查看列表对象my_list索引值为0的元素
# print(my_list[3]) #查看列表对象my_list索引值为3的元素
# print(my_list[1:3:2]) #查看列表对象my_list索引值在1-2之间的元素
# print(len(my_list))
# a=[1,2,3]
# b=[4,5,6]
# c=[7,8,9]
# d=[a,b,c]
# print(a)
# print(d)
# e=[[1, 2, 3],
#    [4, 5, 6],
#    [7, 8, 9]]
# print(len(e))
# e.insert(1,[10,11,12])
# print(e)
# e.append([13,14,15])
# print(e)
# del e[2]
# print(e)
# d={
# '货币资金':1000,
# '应收账款':2000,
# 'a的值':a
# }
# print(d)
# print(list(d.keys()))
# print(d.keys())
# account_dict = {
#     '银行存款': 100000,
#     '应收账款': 50000,
#     '固定资产': 200000,
#     '短期借款': 80000,
#     '实收资本': 270000
# }

# # 获取keys并转换为list
# keys_list = list(account_dict.keys())
# print(keys_list)
# print(list(d.values()))
# print(d.values())
# d['a的值']
# print(d['a的值'])
# d['a的值']=b
# print(d)
# d['c的值']=c
# print(d)
# del d['c的值']
# print(d)
# aa=(100,200,'您好',a,d)
# print(aa)
# print(aa[2:4])
# test=20
# if test in a:
#     print(test,'test在a中')
# else:
#     print(test,'test不在a中')
# test1=15
# if test1<=10:
#     print('差')
# elif test1>10 and test1<=20:
#     print('中')
# elif test1>20 and test1<=30:
#     print('良')
# else:
    # print('优')
# my_list=[100,200,300,400,500]
# new_my_list=[]
# for i in my_list:
#     if i <=300:
#         new_my_list.append(i+10)
#     else:
#         new_my_list.append(i)
# print(new_my_list)
# my_company=[{
#     'name':'阿里巴巴',
#     'profit':30,
#     'stocks':True,
# },{
#     'name':'百度',
#     'profit':5,
#     'stocks':True,  
# },{
#     'name':'腾讯',
#     'profit':50,
#     'stocks':True,
# },{
#     'name':'华为',
#     'profit':100,
#     'stocks':False,
# }]
# new_my_company=[]
# for i in my_company:
#     if i['name']=='阿里巴巴':
#         i['CEO']='马云'
#     elif i['name']=='百度':
#         i['CEO']='李彦宏'
#     elif i['name']=='腾讯':
#         i['CEO']='马化腾'
#     else:
#         i['CEO']='任正非'
#     if i['profit'] >= 30 and i['stocks']==True:
#         new_my_company.append(i['name'])
# print(new_my_company)
# print(my_company)
# e=[[1, 2, 3],
#    [4, 5, 6],
#    [7, 8, 9]]
# new_e=[]
# for i in e:
#     for j in i:
#         new_e.append(j)
#         print(j)
# print(new_e)
# x=0
# while x<7:
#     x=x+2
#     print(x)
# a=[1,2,3,4,5]
# new_a=[]
# for i in a:
#     new_a.append(i+1)
# print(new_a)
# b=[10,20,30,40,50]
# new_b=[]
# for i in b:
#     new_b.append(i+1)
# print(new_b)
# c=[100,200,300,400,500]
# new_c=[]
# for i in c:
#     new_c.append(i+1)
# print(new_c)
# def listPlusOne(originlist):
#     new_list=[]
#     for i in originlist:
#         new_list.append(i+1)
#     print(new_list,'我在函数内')
#     return new_list
# listPlusOne(a)
# listPlusOne(b)
# listPlusOne(c)
# d=[1000,2000,3000,4000,5000]
# new_d=listPlusOne(d)
# print(new_d,'我这函数外')

# a=[1,2,3,4,5]
# a.insert(3,6)
# print(a)
# b=[1,2,3,4,5]
# b.append(6)
# print(b)
# new_b=[]
# for i in b:
#     new_b.append(i+1)
# print(new_b)
# b=[1,2,3,4,5]
# b.insert(2,30)
# print(b)
# new_b=[]
# for i in b:
#     new_b.append(i+1)
# print(new_b)
# a=[
#     [1001,'货币资金',1000],
#     [1122,'应收账款',2000],
#     [1501,'固定资产',5000],
# ]
# new_a=[]
# for i in a:
#     for j in i:
#         new_a.append(j)
# print(new_a)
# b=[
#     [1001,'货币资金',1000],
#     [1122,'应收账款',2000],
#     [1501,'固定资产',5000],
# ]
# new_b=[]
# for i in b:
#     new_b.append(i)
# print(new_b)

# c=[
#     [1001,'货币资金',1000],
#     [1122,'应收账款',2000],
#     [1501,'固定资产',5000],
#     [1502,'累计折旧',3000],
# ]
# new_c=[]
# for i in c:
#     for j in i:
#         new_c.append(j)
# print(new_c)

# def demo1 (num1,num2,num3):
#     num= num1+num2-num3
#     print(num,"我在里面")
#     # return num
#     # return [num,num1,num2,num3]
#     return {
#         "num":num,
#         "num1":num1,
#         "num2":num2,
#         "num3":num3
#     }
# result= demo1(10,20,5)
# print(result,"我在外面")
# 循环遍历列表中的元素
# for k in [1000,'应收账款',2000,'3000',4000]:
# 尝试执行的代码
    # try:
    #     result = k / 2
# 出现异常才会执行的代码
    # except Exception as e:
    #     print('出现错误')
    #     print('错误信息：',e)
# 没有异常才会执行的代码
    # else:
    #     print(result)
    #     print('完成')
# 无论是否有异常，都会执行的代码
    # finally:
    #     print('*' * 10)



# 定义一个方法，参数为公司列表，计算出这个公司列表里面利润最小的公司
#并返回公司名
# def findMinCompany(companyList):
#     min=companyList[0]
#     for i in companyList:
#         if i["profit"]<min["profit"]:
#             min=i
#     return min["name"]

# minCompany=findMinCompany(my_company) 
# print(minCompany)

# 定义一个方法，参数为公司列表，计算出所有公司的利润总和，并返回利润总额打印
# def sumProfit(companyList):
#     profit=0
#     for i in companyList:
#         profit=profit+i["profit"]
#     return profit
# sum=sumProfit(my_company)
# print(sum)


#定义一个方法，参数为公司列表，计算出这个列表里所有公司（含母子公司）所有公司的利润，并返回这个利润
#定义一个方法，参数为公司列表，计算出这个列表里每个母公司及其所有子公司的利润总和，并存放在列表中，并返回这个列表
'''
定义一个方法，参数为公司列表，计算出这个列表里每个母公司及其子
公司的利润总和，并存放在列表中，计算出列表中最大的利润，最小的利润，
输出并返回一个字典，字典中包括利润列表，最大利润，最小利润
'''
# {
#     listProfit:[6, 15, 24, 33],
#     max:33,
#     min:6
# }
my_company=[{
    'name':'阿里巴巴',
    'profit':1,
    'stocks':True,
    'children':[{
        'name':'阿里子公司A',
        'profit':2,
        'stocks':True,
    },{
        'name':'阿里子公司B',
        'profit':3,
        'stocks':False,
    }]
},{
    'name':'百度',
    'profit':4,
    'stocks':True, 
    'children':[{
        'name':'百度子公司A',
        'profit':5,
        'stocks':True,
    },{
        'name':'百度子公司B',
        'profit':6,
        'stocks':False,
    }] 
},{
    'name':'华为',
    'profit':7,
    'stocks':False,
    'children':[{
        'name':'华为子公司A',
        'profit':8,
        'stocks':True,
    },{
        'name':'华为子公司B',
        'profit':9,
        'stocks':False,
    }] 
},{
    'name':'腾讯',
    'profit':10,
    'stocks':True,
     'children':[{
        'name':'腾讯子公司A',
        'profit':11,
        'stocks':True,
    },{
        'name':'腾讯子公司B',
        'profit':12,
        'stocks':False,
    }] 
}]

def totalProfit(companyList):
    profit=0
    for i in companyList:
        profit=profit+i['profit']
        for j in i['children']:
            profit=profit+j['profit']
    return profit

profit=totalProfit(my_company) 
print(profit)

def CompanyProfit(companyList):   
    profit=[]
    for i in companyList:
        # profitCurrent=0
        # profitCurrent=profitCurrent+i['profit']

        profitCurrent=i['profit']
        for j in i['children']:
            profitCurrent=profitCurrent+j['profit']
        profit.append(profitCurrent)
    return profit

profit=CompanyProfit(my_company) 
print(profit)


# def companyProfitDic(companyList):
#     listProfit=[]
#     for i in companyList:
#         profitDic=i['profit']
#         for j in i['children']:
#             profitDic=profitDic+j['profit']
#         listProfit.append(profitDic)
    
#     maxProfit=listProfit[0]
#     minProfit=listProfit[0]
#     for i in listProfit:
#         if i>=maxProfit:
#             maxProfit=i
#         if i<minProfit:
#             minProfit=i 

    # minProfit=listProfit[0]
    # for i in listProfit:
    #     if i<minProfit:
    #         minProfit=i       
    
#     return {
#         'listProfit':listProfit,
#         'max':maxProfit,
#         'min':minProfit
#     }

# profitDic=companyProfitDic(my_company) 
# print(profitDic)

# 定义一个方法，第一个参数为公司列表，列表包括母公司及其子公司，
# 第二个参数为母公司或者子公司名，这个方法找到该名称对应的母公司及子公司集团的所有利润
# 并返回这个利润
# def groupProfit(companyList,companyName):
#     companyGroup={}
#     for i in companyList:
#         if companyName==i['name']:
#             companyGroup=i
#         else:
#             for j in i['children']:
#                 if companyName==j["name"]:
#                     companyGroup=i
#     profit=companyGroup['profit']
#     for i in companyGroup['children']:
#         profit=profit+i['profit']
#     return profit

# profitResult=groupProfit(my_company,'腾讯子公司A')
# print(profitResult)

