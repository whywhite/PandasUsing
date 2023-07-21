import pandas
import pandas as pd
import numpy as np
import re


def data_Insert(df, i, df_add):
    df1 = df.iloc[:i]
    df2 = df.iloc[i:]
    # df_new = pd.concat([df1, df_add, df2], ignore_index=True)
    df_new = df1.append(df_add, ignore_index=True).append(df2, ignore_index=True)
    return df_new


def get_Drpttl(df):
    # drp最后一行增加TTL 加总(P , R , S ,T 栏)

    ttl_P = df['Materials Amount'].astype(float).sum()
    ttl_R = df['Duty'].astype(float).sum()
    ttl_S = df['VAT Tax(13%)'].astype(float).sum()
    ttl_T = df['Amount'].astype(float).sum()

    ttl_dic = {'Active': "TTL", "Materials Amount": ttl_P, "Duty": ttl_R, "VAT Tax(13%)": ttl_S, "Amount": ttl_T}
    df = df.append(ttl_dic, ignore_index=True)
    return df


if __name__ == '__main__':
    # 整个表为df1
    file_path1 = "drp(8.7).xlsx"
    file_path4 = 'drp_example.xlsx'
    file_path5 = 'FXCD FB12 NPI Watch FATP P2 Build Final Cost 0426 ver 2.0.xlsx'
    file_path2 = "保税单1.xlsx"
    file_path3 = "报价单1.xlsx"

    file_path6 = 'FB1 SB FB2 SB DVTP3 FATP FXCD DRP (8.7).xlsx'

    # 保税单，作为df3
    # 修改： 先全部取，然后设置第一行为表头，再通过列名取需要的列
    # df3 = pd.DataFrame(pd.read_excel(io=file_path2, header=None, usecols=['APN', '優惠稅率', '暫定稅率']))
    df3 = pd.DataFrame(pd.read_excel(io=file_path2, header=0))
    df3 = df3[['APN', '優惠稅率', '暫定稅率']]

    # 报价单 ，作为df4,df5(remark)
    # 修改：同dF3
    df4 = pd.DataFrame(pandas.read_excel(io=file_path3, header=0))
    df4 = df4[['OEM PN', 'Price(USD)']]
    df5 = pd.DataFrame(pandas.read_excel(io=file_path3, header=0))
    df5 = df5[['OEM PN', ' Remark']]

    # 取A到O列 作为df2
    # 修改

    df2 = pd.DataFrame(pd.read_excel(io=file_path5, sheet_name='DRP 9.2', header=0, skiprows=13,
                                     usecols=['Active', 'Fill Color', 'Section', 'Component', 'Apple PN', 'OEM PN',
                                              'Vendor', 'Config', 'Rev', 'EEE/EEEE', 'Req',
                                              'OEM DRI', 'U/P', 'PO #', 'GSM Approval']))
    # df2 = df2[[
    #     'Active', 'Fill Color', 'Section', 'Component',
    #     'Apple PN', 'OEM PN', 'Vendor', 'Config',
    #     'Rev', 'EEE/EEEE', 'Req', 'OEM DRI', 'U/P', 'PO #', 'GSM Approval'
    #            ]]
    # 删除df2空行
    df2 = df2[~(df2['Active'].isnull())]


    df2 = df2.fillna('')
    smt_mask = df2['OEM DRI'].str.contains('smt', case=False, na=False)
    apple_mask = df2['U/P'].str.contains('apple to consign', case=False, na=False)
    foc_mask = df2['U/P'].str.contains('foc', case=False, na=False)
    na_mask = df2['U/P'].str.contains('n/a', case=False, na=False)
    applinter_mask = df2['U/P'].str.contains('apple internal', case=False, na=False)
    fconsign_mask = df2['U/P'].str.contains('fX to consign', case=False, na=False)
    mprice_mask = df2['U/P'].str.contains('mp price|mp pricing', case=False, na=False)

    df2.loc[smt_mask, 'U/P'] = 0
    df2.loc[~smt_mask & (apple_mask | foc_mask | na_mask | applinter_mask | fconsign_mask | mprice_mask), 'U/P'] = 0

    df2['U/P'] = pd.to_numeric(df2['U/P'], errors='coerce')
    df2 = df2.fillna(1)
    df2["Materials Amount"] = df2["Req"].astype(float) * df2["U/P"]
    # 取税率，I（8）列大于0 取I 否则取 G（6）列

    # 先将保税清册 APN （第五列） 分成group
    df3.columns = ['Apple PN', 'count_tax', 'temp_tax']
    # df3.drop([0], inplace=True)

    # group_BS = df3.groupby('Apple PN',group_keys=False)

    # 计算duty%列

    df3['Duty%'] = np.where(df3['temp_tax'] > 0, df3['temp_tax'], df3['count_tax'])

    # 只要需要的两列
    df3 = df3[['Apple PN', 'Duty%']]

    # #匹配df2表的apn，将duty%列添加进去
    # df2['Duty%'] = np.where(df3['Apple PN']==df2['Apple PN'],df3['Duty%'])

    # 以drp为基准，匹配保税单的apn，并加上duty%
    df2 = df2.merge(df3, how='left', on='Apple PN')

    # 匹配不到的是空值，方便计算 先将其赋为0
    df2['Duty%'] = df2['Duty%'].fillna(0)

    # 计算R栏 ，

    df2['Duty'] = df2["Materials Amount"].astype(float) * df2["Duty%"].astype(float)

    # 计算S栏
    df2['VAT Tax(13%)'] = df2['Materials Amount'].astype(float) * 1.13

    # 计算T栏
    df2['Amount'] = df2['Materials Amount'] + df2['Duty'] + df2['VAT Tax(13%)']

    # 计算U栏，先设定df4，为OEN PN 和price
    df4.columns = ['OEM PN', '報價檔單價']
    # df4.drop([0], inplace=True)

    df2 = df2.merge(df4, how='left', on='OEM PN')
    # 匹配不到的是空值，方便计算 先将其赋为0
    df2['報價檔單價'] = df2['報價檔單價'].fillna(0)

    # 判斷material amount >0,將單價做一遍篩選
    df2['報價檔單價'] = np.where(df2['Materials Amount'] > 0, df2['報價檔單價'], 0)

    # 将U/P栏与price栏做比较，相同为Y 不同为N

    df2['物料價格對比'] = np.where(df2['U/P'] == df2['報價檔單價'], 'Y', 'N')

    # df5 将remark抓取过来

    df5.columns = ['OEM PN', 'Remark']
    # df5.drop([0], inplace=True)

    df2 = df2.merge(df5, how='left', on='OEM PN')

    # 去重
    df2 = df2.drop_duplicates()

    # 保留前三十行数据
    # df2 = df2.head(30)

    # 添加TTL
    df2 = get_Drpttl(df2)
    # print(df2)

    # 插入DRP版本 和TTL AMOUNT

    df_version = pd.DataFrame(pandas.read_excel(io=file_path1, header=None))
    # print(df_version.head())
    drp_Version = df_version.iloc[2, 0]
    p1 = re.compile(r'[(](.*?)[)]', re.S)
    p2 = ''.join(re.findall(p1, drp_Version))
    ttl_T_2 = df2['Amount'].astype(float).sum() / 2
    df_add = pd.DataFrame({'Active': 'Based on DRP v' + p2, "VAT Tax(13%)": 'TTL Amount', "Amount": ttl_T_2}, index=[0])
    df2 = data_Insert(df2, 0, df_add)



    # 表头插入

    df_col1 = pd.DataFrame(df2.columns)
    df_col2 = df_col1.T
    df_col2.columns = df2.columns
    # print(df_col2)
    df_new = data_Insert(df2, 1, df_col2)

    # df4.to_excel("new_Data2.xlsx",index=False,header=False)
    # 保存到新的excel文件

    # df2.to_excel("new_Data1.xlsx", index=False, header=True)
    df_new.to_excel("new_Data3.xlsx", index=False, header=False)
