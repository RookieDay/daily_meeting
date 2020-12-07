# -*- coding: utf-8 -*-
"""
Created on Mon Jun 10 15:38:38 2019

@author: Xinchun
根据上市公司进行分类
"""
import pandas as pd
import numpy as np

def tmzn_replace(tmZn):
    if tmZn[1] == '3':
        time = '6月13日'+ tmZn[3:]
    else:
        time = '6月14日'+ tmZn[3:]
    return time 

data = pd.read_excel('详细排会安排表-June.xlsx')

# 找到所有的房间编号 特别注意最后一个编号要录入啊
roomListIdx = []
for idx in data.index:
    if type(data.loc[idx][0]) == str:
        roomListIdx.append(idx)
roomListIdx.append(idx+1)

# 有效的时间列表
timeZone = data.columns[1:]

# 汇集上市公司的信息
tmp = pd.DataFrame([])
count = 0
for i in range(0, len(roomListIdx)-1):
    iRaw =  roomListIdx[i]
    nextRaw = roomListIdx[i+1]
    roomName = data.loc[iRaw][0]
    for col in data.columns[1:]:
        cmpName = data.loc[iRaw][col]
        if type(cmpName) == str and len(cmpName) > 1: #这是一个有效的上市公司名称
            pos = cmpName.find("-")
            cmpName = cmpName[:pos]
            print(cmpName)
            #如果是1v1情况要特殊处理
            clientInfo = data.loc[iRaw+1:nextRaw-1][col]
            clientInfo = clientInfo[clientInfo.isna() == False]
            aLine = [cmpName, col, roomName, "    ".join(list(clientInfo))]
            tmp[count] = aLine
            count += 1
tmp = tmp.T
tmp.columns = ['上市公司', '时间', '房间号', '客户清单']

cmpList = list(set(tmp['上市公司']))
result = pd.DataFrame([])
count = 0
for cmp in cmpList:
    aLine = [cmp, '时间', '房间号', '客户清单']
    result[count] = aLine
    count += 1
    for tmZn in timeZone:
        part = tmp[tmp['上市公司'] == cmp]
        part = part[part['时间'] == tmZn]
        #时间进行替换
        fulltmZn = tmzn_replace(tmZn)
        if len(part) > 0:
            aLine = [cmp, fulltmZn, part.iloc[0]['房间号'], part.iloc[0]['客户清单']]
        else:
            aLine = [cmp, fulltmZn, np.nan, np.nan]
        result[count] = aLine
        count += 1

result = result.T
result.columns = ['上市公司', '时间', '房间号', '客户清单']
result.to_excel('上市公司提取结果.xlsx', index=False)
        

