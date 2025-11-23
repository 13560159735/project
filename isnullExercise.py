# 请帮我用程序统计处a中有多少个缺失值,以下面的数据为例，统计出每一横行的缺失值[2,1,1,2]
a=[
    ['张三','李四',None,'赵六',None],
    [25,None,55,87,44],
    [3,7,1,None,3],
    [None,'滨河','渤海',None,'黄埔']
]
# len(a)
# print(len(a))

# len(a[0])
# print(len(a[0]))

# sum=0
# for i in range(len(a)):
#     for j in a[i]:
#         if j==None:
#             sum=sum+1
    
# print(sum)


numList=[]
for i in a:
    sum=0
    for j in i:
        if j==None:
            sum=sum+1
    numList.append(sum)
    
print(numList)


