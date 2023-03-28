import pandas
import pandas as pd
import numpy as np

#整个表为df1
file_path1 = "drp_example.xlsx"
file_path2 = "保税单1.xlsx"
file_path3 = "报价单1.xlsx"
df1 = pd.DataFrame(pandas.read_excel(io=file_path1,header=None,skiprows=13))

#保税单，作为df3
df3 = pd.DataFrame(pandas.read_excel(io=file_path2,header=None,usecols="E,G,I"))

#报价单 ，作为df4,df5(remark)
df4 = pd.DataFrame(pandas.read_excel(io=file_path3,header=None,usecols="C,G"))

df5 = pd.DataFrame(pandas.read_excel(io=file_path3,header=None,usecols="C,J"))
#取A到O列 作为df2
df2 = pd.DataFrame(pandas.read_excel(io=file_path1,header=None,skiprows=13,usecols="A:H,M,N,O,U,AH,AI,AJ"))


#删除df1,df2空行
df1 = df1[~(df1[0].isnull())]
df2 = df2[~(df2[0].isnull())]


#保存到新的excel文件
# df2.to_excel("new_Data1.xlsx",index=False,header=False)


#P栏位的计算方式，先做判断U栏位含有相关字符串，就把AH栏(34)相应位置为0
# "U"列 = 21 列 OEE DRI 列

str1 = "SMT"
str2 = "apple to consign"
str3 = "FOC"
str4 = "N/A"

row_range = range(1,len(df1))

for row in row_range:
    str_U = df2.iloc[row,11]
    str_AH = df2.iloc[row,12]
    if type(str_U) == type(str1) and type(str_AH) == type(str1):
        if str_U.find(str1)!=-1:
            df2.iat[row,12] = 0
        elif str.lower(str_AH).find(str2)!=-1 or str_AH.find(str3)!=-1 or str_AH.find(str4)!=-1:
            df2.iat[row,12] = 0

#设置表头为第一行
c_list = df2.values.tolist()[0]
df2.columns = c_list
df2.drop([0],inplace=True)

#计算P栏

df2["Materials Amount"] = df2["Req"].astype(float)*df2["U/P"].astype(float)


#取税率，I（8）列大于0 取I 否则取 G（6）列


#先将保税清册 APN （第五列） 分成group
df3.columns = ['Apple PN', 'count_tax', 'temp_tax']
df3.drop([0],inplace=True)


# group_BS = df3.groupby('Apple PN',group_keys=False)



#计算duty%列

df3['Duty%'] = np.where(df3['temp_tax']>0,df3['temp_tax'],df3['count_tax'])

#只要需要的两列
df3 = df3[['Apple PN','Duty%']]

# #匹配df2表的apn，将duty%列添加进去
# df2['Duty%'] = np.where(df3['Apple PN']==df2['Apple PN'],df3['Duty%'])

#以drp为基准，匹配保税单的apn，并加上duty%
df2 = df2.merge(df3, how='left', on='Apple PN')

#匹配不到的是空值，方便计算 先将其赋为0
df2['Duty%'] = df2['Duty%'].fillna(0)



#计算R栏 ，

df2['Duty'] = df2["Materials Amount"].astype(float)*df2["Duty%"].astype(float)

#计算S栏
df2['VAT Tax(13%)'] = df2['Materials Amount'].astype(float)*1.13

#计算T栏
df2['Amount'] = df2['Materials Amount']+df2['Duty']+df2['VAT Tax(13%)']

#计算U栏，先设定df4，为OEN PN 和price
df4.columns = ['OEM PN', '報價檔單價']
df4.drop([0],inplace=True)

df2 = df2.merge(df4, how='left', on='OEM PN')
#匹配不到的是空值，方便计算 先将其赋为0
df2['報價檔單價'] = df2['報價檔單價'].fillna(0)

#判斷material amount >0,將單價做一遍篩選
df2['報價檔單價'] = np.where(df2['Materials Amount']>0,df2['報價檔單價'],0)

#将U/P栏与price栏做比较，相同为Y 不同为N


df2['物料價格對比'] = np.where(df2['U/P'] == df2['報價檔單價'],'Y','N')

#df5 将remark抓取过来

df5.columns = ['OEM PN', 'Remark']
df5.drop([0],inplace=True)

df2 = df2.merge(df5,how='left',on='OEM PN')

#去重
df2 =df2.drop_duplicates()

# df4.to_excel("new_Data2.xlsx",index=False,header=False)
#保存到新的excel文件
df2.to_excel("new_Data1.xlsx",index=False,header=True)


