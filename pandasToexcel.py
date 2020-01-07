import pandas as pd
import openpyxl
import numpy
import datetime


def read():
    # ecl = pd.ExcelFile('C:\\Users\\Administrator\\Desktop\\123.xlsx')
    # for j in ecl.sheet_names:
    #     for i in pd.read_excel(ecl,j):
    #         print(i)
    ecl = pd.read_excel('C:\\Users\\Administrator\\Desktop\\123.xlsx', sheet_name='123')
    ecl2 = pd.read_excel('C:\\Users\\Administrator\\Desktop\\061階工令報表12-18.xlsx', sheet_name='SMT (570)')
    df1 = pd.DataFrame(ecl)
    df2 = pd.DataFrame(ecl2)
    # result = pd.merge(df1,df2.loc[:,'C2','B4'],how='left',on='C2')
    result = df1.merge(df2, left_on='Order', right_on='工令編碼', how='left')
    result.to_excel('C:\\Users\\Administrator\\Desktop\\1234.xlsx')
    for i in result.values:
        print(i)
    # print(result.values)
    # ecl = pd.read_excel('C:\\Users\\Administrator\\Desktop\\123.xlsx')
    # ecl2 = pd.read_excel('C:\\Users\\Administrator\\Desktop\\067階工令報表12-18.xlsx')
    # for i in ecl2.values:
    #     print(i)
    # ecl = pd.ExcelFile('C:\\Users\\Administrator\\Desktop\\067階工令報表12-18.xlsx')
    # for j in ecl.sheet_names:
    #     print(j)


def get_mac(item):
    # ecl = pd.read_excel('C:\\Users\\Administrator\\Desktop\\123.xlsx', sheet_name='123')
    # ecl2 = pd.read_excel('C:\\Users\\Administrator\\Desktop\\061階工令報表12-18.xlsx', sheet_name='SMT (570)')
    df = pd.read_excel('C:\\Users\\Administrator\\Desktop\\123.xlsx', sheet_name='123')  # 需要匹配的数据表
    data = pd.read_excel('C:\\Users\\Administrator\\Desktop\\061階工令報表12-18.xlsx', sheet_name='SMT (570)')  # 被匹配的数据表
    sn = [x for x in data['工令編碼']]  # 被匹配的数据表中的字段
    df['L2'] = get_mac(df['Order'])
    a = []
    for x in item:
        try:
            mac_2 = list(data.loc[[sn.index(x)]]['Order'])[0]  # 定位到bb表的行，并取出要匹配的字段
            print(x, mac_2)
            a.append(mac_2)
        except:
            print(x, 0)
            a.append(0)


def vlookup():
    # ecl = pd.read_excel('C:\\Users\\Administrator\\Desktop\\123.xlsx', sheet_name='123')
    # ecl2 = pd.read_excel('C:\\Users\\Administrator\\Desktop\\061階工令報表12-18.xlsx', sheet_name='SMT (570)')
    # table_a_name = input("请输入A表文件名：")
    table_a_path = 'C:\\Users\\Administrator\\Desktop\\123.xlsx'
    # sheet_a_name = input("请输入A表中的sheet名称：")
    table_a = pd.read_excel(table_a_path, sheet_name='123', converters={'Order': str}).dropna(axis=1, how='all')
    # table_b_name = input("请输入B表文件名：")
    table_b_path = 'C:\\Users\\Administrator\\Desktop\\061階工令報表12-18.xlsx'
    # sheet_b_name = input("请输入B表中的sheet名称：")
    table_b = pd.read_excel(table_b_path, sheet_name='SMT (570)', converters={'工令編碼': str})
    table_b_2 = table_b.groupby("Order").L1.sum().reset_index()
    table_c = table_a.merge(right=table_b_2, how='left', left_on='Order', right_on='工令編碼')
    table_c.to_excel('c.xlsx', index=False)


def vlookups():
    df = pd.read_excel("C:\\Users\\Administrator\\Desktop\\123.xlsx", sheet_name='123')
    # 提取Name列
    s = df["Order"]
    # 转为list
    listName = s.tolist()  # list
    # 在list中修改字符串
    # for i, v in enumerate(listName):
    #     print(v)
    #     listName[i] = str(v).strip()[str(v).index(']') + 2:str(v).index(']') + 11]
    # print(listName)
    # list转为dataframe
    data = pd.DataFrame(listName, columns=['Chapter'])
    # print(data)
    # 按列拼接dataframe
    dfA = pd.concat([df, data], axis=1)
    # print(dfA)
    # 合并dataframe
    dfB = pd.read_excel("C:\\Users\\Administrator\\Desktop\\061階工令報表12-18.xlsx", sheet_name='SMT (570)')
    print(dfB['工令編碼'])
    # 对关键字Chapter列向左连接（左边dfA为全部）
    dfC = pd.merge(dfA, dfB, how='left', left_on='Chapter', right_on='工令編碼')  # on=['Chapter']
    # print(dfC)
    # 保存到csv中
    dfC.to_excel('genSum.xlsx', encoding="utf_8_sig")


def myvlookup():
    df1 = pd.read_excel("feng.xlsx", sheet_name='工作表1')
    df2 = pd.read_excel("liu.xlsx", sheet_name='工作表1')
    # 获取你需要vlookup的列
    s = df1["姓名"]
    mo = df2['姓名']
    listName = s.tolist()
    moList = mo.tolist()
    # 定义一个列表用来存储vlookup结果
    result = []
    flag = False
    # 判断是否有相等的值
    for i in range(len(listName)):
        # for j in range(len(moList)):
        #     # print(i)
        #     # print(listName[i] == moList[j])
        #     if (str(listName[i]).strip() == str(moList[j]).strip()):
        #         flag = True
        #         break
        # #保证result的数据与取出来的数据位置一致
        # if flag:
        #     result.append(listName[i])
        #     flag = False
        # else:
        #     result.append('N/A')
        result.append("=VLOOKUP(A"+str(2+i)+",[liu.xlsx]工作表1!$A$2:$A$7,1,0)")
    print(result)
    # 准备数据
    data = pd.DataFrame(result, columns=['result'])
    # 将结果添加到1表的最后一列
    resultdf = pd.concat([df1, data], axis=1)
    #导出结果
    resultdf.to_excel('result.xlsx', index=False)


if __name__ == '__main__':
    # read()
    # vlookups()
    print('program start')
    start = datetime.datetime.now()
    print(datetime.datetime.now())
    myvlookup()
    print('program stop')
    print(datetime.datetime.now())
    end = datetime.datetime.now()
    print('cost')
    print(end - start)
