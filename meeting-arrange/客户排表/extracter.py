# -*- coding: utf-8 -*-
"""
Created on Mon Jun  4 09:19:52 2018

@author: ZTFE
"""

import numpy as np
import pandas as pd

def find_room(idx):
    for i in range(0, len(roomListIdx)):
        if roomListIdx[i] > idx:
            break
    return i-1

def assemble(agent, ins, name, roomIdx, col):
    idx = roomListIdx[roomIdx]
    # 提取房间名称和股票代码
    roomStr = data.loc[idx]['Time']
    stock = data.loc[idx][col]
#    pos = len(stock) -1
#    while (stock[pos] >='0' and stock[pos] <= '9'):
#        pos = pos-1
#    stock = stock[:pos+1]
    if col[1] == '3':
        time = '6月13日'+ col[3:]
    else:
        time = '6月14日'+ col[3:]
    aLine = [agent, ins, name,  time, stock, roomStr]
    return aLine

def virtual_assemble(agent, ins, name, col):
    roomStr = np.nan
    stock = np.nan
    if col[1] == '3':
        time = '6月13日'+ col[3:]
    else:
        time = '6月14日'+ col[3:]
    aLine = [agent, ins, name,  time, stock, roomStr]
    return aLine

def merge(preLine, aLine):
    preStock = preLine[4]
    preRoom = preLine[5]
    stock = preStock + '\n' + aLine[4]
    room = preRoom + '\n' + aLine[5]    
    aLine[4] = stock
    aLine[5] = room
    return aLine

data = pd.read_excel('详细排会安排表-June.xlsx')
clientData = pd.read_excel('客户报名情况反馈表.xlsx')

# 找到所有的房间编号 特别注意最后一个编号要录入啊
roomListIdx = []
for idx in data.index:
    if type(data.loc[idx][0]) == str:
        roomListIdx.append(idx)
roomListIdx.append(idx+1)

# 根据客户信息，找到他们的房间编号和时间
result = pd.DataFrame([])
count = 0
# 没有匹配的做下记录
unmatchedList = []

for i, clnt in enumerate(clientData['客户信息']):
    nameID = clnt
    nameID = nameID.strip()
    print(nameID)
    ins = clientData.loc[i]['机构名称']
    agent = clientData.loc[i]['对口销售']
    name = clientData.loc[i]['姓名']
    # 加入空行作为分割
    aLine = [agent, ins, name,  '时间', '上市公司', '房间号']
    result[count] = aLine
    count += 1
    totalfind = 0
    for col in data.columns[1:]:
        flag = 0
        for idx in data.index:    
            context = data.loc[idx][col]
            if type(context) == str and context.find(nameID) > -1:
                # 判定 房间号和日期
                roomIdx = find_room(idx)
                if flag < 1:                
                    aLine = assemble(agent, ins, name, roomIdx, col)
                    result[count] = aLine
                    count += 1
                    flag =1
                    totalfind = 1
                else:
                    preLine = aLine
                    aLine = assemble(agent, ins, name, roomIdx, col)                    
                    bLine = merge(preLine, aLine)
                    count -= 1
                    result[count] = bLine
                    count += 1
        if flag < 1:
            unmatchedList += [nameID]
            aLine = virtual_assemble(agent, ins, name, col)
            result[count] = aLine
            count += 1
    if totalfind <1:
        print(nameID)
result = result.T
result.columns = ['对口销售', '机构', '姓名','时间', '上市公司','房间号']
result.to_excel('客户-提取结果.xlsx', index = False)