#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jun 25 11:42:12 2021

@author: jiangping
"""



import pandas as pd
import numpy as np
from time import *
import os.path
from tkinter import *
from tkinter.filedialog import *


#路径选择1
def selectPath1():
    global path_1
    path_ = askopenfilename()
    path1.set(path_)
    path_1 = path_

#路径选择2
def selectPath2():
    global path_2
    path_ = askopenfilename()
    path2.set(path_)
    path_2 = path_
    
#路径选择3
def selectPath3():
    global path_3
    path_ = askopenfilename()
    path3.set(path_)
    path_3 = path_
    
#路径选择4
def selectPath4():
    global path_4
    path_ = askdirectory()
    path4.set(path_)
    path_4 = path_


def fun():
    start =time()
    #老系统读数据
    t1 = pd.read_excel(path_1)
    t1.drop(t1.index, inplace=True)
    
    path = path_1[0:path_1.rfind('/')]
    file_name=os.listdir(path)
  #  file_name.remove('.DS_Store')
    
    for i in file_name: 
        print("正在读取 " + i)
        file_path=path+'\\'+i           
        t = pd.read_excel(file_path)
        t1= pd.concat([t1, t], ignore_index=True)
    
    # t1.rename(columns = {'设备ID(设备唯一码)':'设备ID'},inplace=True)
    t1.columns=t1.columns.str.replace(' ','')
    t1['设施类型']=t1['设施类型'].astype(str).astype(object)
    t1['电压等级']=t1['电压等级'].astype(str).astype(object)
    t1['局厂名称']=t1['局厂名称'].astype(str).astype(object)
    # t1['变电站代码']=t1['变电站代码'].astype(str).astype(object)
    # t1['线路代码/安装位置码']=t1['线路代码/安装位置码'].astype(str).astype(object)
    # bs1=t1['局厂代码']+t1['下属代码']+t1['变电站代码']+t1['线路代码/安装位置码']+t1['设备ID']
    bs1=t1['局厂名称']+t1['设施类型']+t1['电压等级']
    t1.insert(0,'设备标识',bs1)
    
    
    
    #新系统读数据
    t2= pd.read_excel(path_2,sheet_name='sheet1')
    t2.drop(t2.index, inplace=True)
    
    table = pd.read_excel(path_2,sheet_name=None)
    sheet_name=list(table.keys())
      
    for j in sheet_name:  
        print("正在读取 " + j)
        #sheetdata=table[j]
        sheetdata=pd.read_excel(path_2,sheet_name=j)
        t2= pd.concat([t2, sheetdata], ignore_index=True)
    
    # t2['变电站代码']=t2['变电站代码'].astype(str).astype(object)
    # t2['单位代码'] = t1['单位代码'].astype(str).astype(object)
    t2['设备类别'] = t2['设备类别'].astype(str).astype(object)
    t2['单位名称'] = t2['单位名称'].astype(str).astype(object)
    t2['电压等级(kV)'] = t2['电压等级(kV)'].astype(str).astype(object)
    bs2=t2['单位名称']+t2['设备类别']+t2['电压等级(kV)']
    t2.insert(0,'设备标识',bs2)
    
    
    
    #---开始比对前检查的分割线----------------------------------------------------------------------------------------
    
    #匹配失败检查
    mgi=pd.merge(t1['设备标识'].drop_duplicates(),t2['设备标识'].drop_duplicates(),how='inner')
    t1_sh3=t1[~(t1['设备标识'].isin(mgi['设备标识']))]
    t2_sh3=t2[~(t2['设备标识'].isin(mgi['设备标识']))]
    
    
    #重复检查(剔除匹配失败记录后)
    v1=t1[~(t1['设备标识'].isin(t1_sh3['设备标识']))]['设备标识'].value_counts()
    v2=t2[~(t2['设备标识'].isin(t2_sh3['设备标识']))]['设备标识'].value_counts()
    
    mgo=pd.merge(t1[t1['设备标识'].isin(v1[v1>1].index)]['设备标识'],t2[t2['设备标识'].isin(v2[v2>1].index)]['设备标识'],how='outer')
    mgo.drop_duplicates(inplace=True)
    
    t1_sh2=t1[t1['设备标识'].isin(mgo['设备标识'])]
    t2_sh2=t2[t2['设备标识'].isin(mgo['设备标识'])]
    
    
    #剩余一对一匹配清单
    t1_sh1=t1[~(t1['设备标识'].isin(mgo['设备标识']) | t1['设备标识'].isin(t1_sh3['设备标识']))]
    t2_sh1=t2[~(t2['设备标识'].isin(mgo['设备标识']) | t2['设备标识'].isin(t2_sh3['设备标识']))]
    
    
    
    #疑似重复及关联失败清单
    # file_name1=path_4+'\\'+par1
    # writer1 = pd.ExcelWriter(file_name1, engine='xlsxwriter')
    #
    # t1_sh2.to_excel(writer1,sheet_name='疑似重复-老系统',index=False)
    # t2_sh2.to_excel(writer1,sheet_name='疑似重复-新系统',index=False)
    # t1_sh3.to_excel(writer1,sheet_name='关联失败-老系统',index=False)
    # t2_sh3.to_excel(writer1,sheet_name='关联失败-新系统',index=False)
    #
    # writer1.save()
    # print('已生成 疑似重复及关联失败清单')
    
    
    #---开始指标比对的分割线----------------------------------------------------------------------------------------

    ls1=['设备标识','设施类型'  ,'单位代码'  ,'单位代码'  ,'局厂名称'  ,'下属代码'  ,'下属单位名称'  ,'变电站代码'  ,'变电站名称'  ,'电压等级'  ,'元件类型'  ,'设备台数/线路条数'  ,'线路长度/设备容量'  ,'设备百台年数'  ,'统计期间小时'  ,'可用小时'  ,'运行小时'  ,'备用小时'  ,'调度停运备用小时'  ,'作业前受累停运备用时间'  ,'作业后受累停运备用时间'  ,'受累停运备用小时'  ,'不可用小时'  ,'计划停运小时'  ,'大修停运小时'  ,'小修停运小时'  ,'试验停运小时'  ,'清扫停运小时'  ,'改造施工停运小时'  ,'非计划停运小时'  ,'第1类非计划停运小时'  ,'第2类非计划停运小时'  ,'第3类非计划停运小时'  ,'第4类非计划停运小时'  ,'强迫停运小时'  ,'备用次数'  ,'调度停运备用次数'  ,'作业前受累停运备用次数'  ,'作业后受累停运备用次数'  ,'受累停运备用次数'  ,'计划停运次数'  ,'大修次数'  ,'小修次数'  ,'试验次数'  ,'清扫次数'  ,'改造施工次数'  ,'非计划停运次数'  ,'第1类非计划停运次数'  ,'第1类非停运(重合成功)次数'  ,'第2类非计划停运次数'  ,'第3类非计划停运次数'  ,'第4类非计划停运次数'  ,'强迫停运次数'  ,'计划停运系数'  ,'非计划停运系数'  ,'强迫停运系数'  ,'可用系数'  ,'运行系数'  ,'计划停运率'  ,'非计划停运率'  ,'强迫停运率'  ,'暴露率'  ,'平均无故障操作次数'  ,'连续可用小时'  ,'正确动作率'  ,'总操作次数'  ,'正常操作次数'  ,'调试操作次数'  ,'切除故障次数'  ,'非正确动作次数'  ,'带电作业次数'  ,'带电作业小时数'  ,'非计划停运条次比'  ,'线路跳闸率'  ,'线路跳闸重合成功率'  ,'GIS内部元件数'  ,'GIS内部断路器个数'  ,'GIS内部电流互感器个数'  ,'GIS内部电压互感器个数'  ,'GIS内部隔离开关个数'  ,'GIS内部避雷器个数'  ,'GIS内部母线段数'  ,'合成绝缘子串数'  ,'玻璃绝缘子串数'  ,'瓷质绝缘子串数'  ,'其它绝缘子串数'  ,'平均退役寿命'  ,'最大退役寿命'  ,'最小退役寿命'  ,'平均投运年限'  ,'最大停运次数'  ,'在投设备数'  ,'退出设备数'  ,'线路计算方法'  ,'统计方式'  ,'起始时间'  ,'终止时间'  ,'设计制造单位类型'  ,'统计分类'  ,'统计任务ID'  ,'内部电流互感器数'  ,'内部隔离开关数'  ,'内部断路器数'  ,'退出原因'  ,'杆塔总数'  ,'杆塔总数'  ,'资产性质'  ]
    # ls1=['设备标识','元件在用小时数','可用系数','可用小时','运行系数','运行小时','暴露率','计划停运系数','计划停运小时','非计划停运系数','非计划停运小时',
    #'强迫停运系数','强迫停运小时','设备台年数','计划停运率','计划停运次数','非计划停运率','非计划停运次数','强迫停运率','强迫停运次数','连续可用小时']
    ls2=['设备标识','统计任务ID',	'设备类别',	'单位代码',	'单位名称',	'局厂代码',	'局厂名称',	'下属代码',	'下属单位名称',	'变电站代码',	'变电站名称',	'电压等级(kV)',	'设备台数/线路条数',	'线路长度/设备容量',	'统计设备数',	'起始时间',	'统计期间小时',	'设备百台年数',	'可用小时',	'运行小时',	'备用小时',	'调度停运备用小时',	'作业前受累停运备用小时',	'作业后受累停运备用小时',	'受累停运备用小时',	'不可用小时',	'计划停运小时',	'大修停运小时',	'小修停运小时',	'试验停运小时',	'清扫停运小时',	'改造施工停运小时',	'非计划停运小时',	'第一类非计划停运小时',	'第二类非计划停运小时',	'第三类非计划停运小时',	'第四类非计划停运小时',	'强迫停运小时',	'备用次数',	'调度停运备用次数',	'作业前受累停运备用次数',	'作业后受累停运备用次数',	'受累停运备用次数',	'计划停运次数',	'大修次数',	'小修次数',	'试验次数',	'清扫次数',	'改造施工次数',	'非计划停运次数',	'第一类非计划停运次数',	'第一类非停运(重合成功)次数',	'第二类非计划停运次数',	'第三类非计划停运次数',	'第四类非计划停运次数',	'强迫停运次数',	'最大停运次数',	'计划停运系数',	'非计划停运系数',	'强迫停运系数',	'可用系数',	'运行系数',	'计划停运率',	'非计划停运率',	'强迫停运率',	'暴露率',	'连续可用小时',	'计划停运影响可用系数占比(%)',	'非计划停运影响可用系数占比(%)',	'终止时间',	'元件类型',	'创建时间']
    # ls2=['设备标识','设备ID','局厂名称','下属单位名称','变电站名称','安装位置代码','安装位置名称','统计期间小时','可用系数','可用小时','运行系数','运行小时','暴露率','计划停运系数','计划停运小时','非计划停运系数','非计划停运小时',
    #'强迫停运系数','强迫停运小时','设备百台年数','计划停运率','计划停运次数','非计划停运率','非计划停运次数','强迫停运率','强迫停运次数','连续可用小时']
    
    # ls=['设备标识','设备ID','局厂名称','下属单位名称','变电站名称','安装位置代码','安装位置名称',
    #'old统计期间小时','new统计期间小时','old可用系数','new可用系数','可用系数是否一致',
    #'old可用小时','new可用小时','old运行系数','new运行系数','运行系数是否一致',
    #'old运行小时','new运行小时','old暴露率','new暴露率','暴露率是否一致',
    #'old计划停运系数','new计划停运系数','计划停运系数是否一致','old计划停运小时','new计划停运小时',
    #'old非计划停运系数','new非计划停运系数','非计划停运系数是否一致','old非计划停运小时','new非计划停运小时',
    #'old强迫停运系数','new强迫停运系数','强迫停运系数是否一致','old强迫停运小时','new强迫停运小时',
    #'old设备台年数','new设备台年数',
    #'old计划停运率','new计划停运率','计划停运率是否一致','old计划停运次数','new计划停运次数',
    #'old非计划停运率','new非计划停运率','非计划停运率是否一致','old非计划停运次数','new非计划停运次数',
    #'old强迫停运率','new强迫停运率','强迫停运率是否一致','old强迫停运次数','new强迫停运次数',
    #'old连续可用小时','new连续可用小时','连续可用小时是否一致']

    ls = ['设备标识',
         'old可用系数','new可用系数','可用系数是否一致',
         'old可用小时','new可用小时','可用小时是否一致','old运行系数','new运行系数','运行系数是否一致',
         'old运行小时','new运行小时','运行小时是否一致','old暴露率','new暴露率','暴露率是否一致',
         'old计划停运系数','new计划停运系数','计划停运系数是否一致','old计划停运小时','new计划停运小时','计划停运小时是否一致',
         'old非计划停运系数','new非计划停运系数','非计划停运系数是否一致','old非计划停运小时','new非计划停运小时','非计划停运小时是否一致',
         'old强迫停运系数','new强迫停运系数','强迫停运系数是否一致','old强迫停运小时','new强迫停运小时','强迫停运小时是否一致',

         'old计划停运率','new计划停运率','计划停运率是否一致','old计划停运次数','new计划停运次数',
         'old非计划停运率','new非计划停运率','非计划停运率是否一致','old非计划停运次数','new非计划停运次数',
         'old强迫停运率','new强迫停运率','强迫停运率是否一致','old强迫停运次数','new强迫停运次数'
        ,
         'old连续可用小时','new连续可用小时'
        ,'连续可用小时是否一致'
          ]
    dic1={'设备标识':'设备标识',
         '可用系数':'old可用系数',
         '可用小时':'old可用小时',
         '运行系数':'old运行系数',
         '运行小时':'old运行小时',
         '暴露率':'old暴露率',
         '计划停运系数':'old计划停运系数',
         '计划停运小时':'old计划停运小时',
         '非计划停运系数':'old非计划停运系数',
         '非计划停运小时':'old非计划停运小时',
         '强迫停运系数':'old强迫停运系数',
         '强迫停运小时':'old强迫停运小时',
         '设备台年数':'old设备台年数',
         '计划停运率':'old计划停运率',
         '计划停运次数':'old计划停运次数',
         '非计划停运率':'old非计划停运率',
         '非计划停运次数':'old非计划停运次数',
         '强迫停运率':'old强迫停运率',
         '强迫停运次数':'old强迫停运次数'
        ,
         '连续可用小时':'old连续可用小时'
          }
    
    dic2={'设备标识':'设备标识',

         '统计期间小时':'new统计期间小时',
         '可用系数':'new可用系数',
         '可用小时':'new可用小时',
         '运行系数':'new运行系数',
         '运行小时':'new运行小时',
         '暴露率':'new暴露率',
         '计划停运系数':'new计划停运系数',
         '计划停运小时':'new计划停运小时',
         '非计划停运系数':'new非计划停运系数',
         '非计划停运小时':'new非计划停运小时',
         '强迫停运系数':'new强迫停运系数',
         '强迫停运小时':'new强迫停运小时',
         '设备百台年数':'new设备台年数',
         '计划停运率':'new计划停运率',
         '计划停运次数':'new计划停运次数',
         '非计划停运率':'new非计划停运率',
         '非计划停运次数':'new非计划停运次数',
         '强迫停运率':'new强迫停运率',
         '强迫停运次数':'new强迫停运次数'
        ,
         '连续可用小时':'new连续可用小时'
          }
    
    
    #一致性校验函数    
    def Diff(df,col1,col2,diff):
        df[diff]= np.where((round(df[col2],3)-round(df[col1],3))==0,1,0)
        return df
    
    
    #老系统字段筛选
    df1=t1_sh1.loc[:,ls1]
    df1.rename(columns=dic1, inplace=True)
    
    #新系统字段筛选
    df2=t2_sh1.loc[:,ls2]
    df2.rename(columns=dic2, inplace=True)
    
    #union后创建校验字段
    df0 = pd.merge(df1,df2,how='inner',on='设备标识')
    df0=Diff(df0,'new可用系数','old可用系数','可用系数是否一致')
    df0=Diff(df0,'new运行系数','old运行系数','运行系数是否一致')
    df0=Diff(df0,'new暴露率','old暴露率','暴露率是否一致')
    df0=Diff(df0,'new计划停运系数','old计划停运系数','计划停运系数是否一致')
    df0=Diff(df0,'new非计划停运系数','old非计划停运系数','非计划停运系数是否一致')
    df0=Diff(df0,'new强迫停运系数','old强迫停运系数','强迫停运系数是否一致')
    df0=Diff(df0,'new计划停运率','old计划停运率','计划停运率是否一致')
    df0=Diff(df0,'new非计划停运率','old非计划停运率','非计划停运率是否一致')
    df0=Diff(df0,'new强迫停运率','old强迫停运率','强迫停运率是否一致')
    df0=Diff(df0,'new连续可用小时','old连续可用小时','连续可用小时是否一致')
    df0=Diff(df0,'new可用小时','old可用小时','可用小时是否一致')
    df0=Diff(df0,'new运行小时','old运行小时','运行小时是否一致')
    df0=Diff(df0,'new计划停运小时','old计划停运小时','计划停运小时是否一致')
    df0=Diff(df0,'new强迫停运小时','old强迫停运小时','强迫停运小时是否一致')
    df0=Diff(df0,'new非计划停运小时','old非计划停运小时','非计划停运小时是否一致')

    #调整字段顺序
    sh1=df0.loc[:,ls]
    
    
    #汇总
    sh2 = pd.DataFrame(columns=['合计','一致','不一致','一致占比'],index=['可用系数','运行系数','暴露率','计划停运系数','非计划停运系数','强迫停运系数','计划停运率','非计划停运率','强迫停运率'
        ,'连续可用小时'
                                                               ])
    sh2['合计']=sh1['设备标识'].shape[0]
    
    for idx in sh2.index:
        if idx == '连续可用小时':
            sh2['一致'][idx] = 0
        else:
            col = idx+'是否一致'
            sh2['一致'][idx] = sh1[col].value_counts()[1]
            sh2['不一致'][idx] = sh2['合计'][idx]-sh1[col].value_counts()[1]
            sh2['一致占比'][idx] = sh1[col].value_counts()[1]/sh2['合计'][idx]

    sh2.insert(0,'指标',sh2.index)
    
    #筛选问题清单
    q1=sh1
    q1['设备标识']=q1.apply(lambda x: str(x['设备标识']), axis=1)
            
    v1=q1[q1['可用系数是否一致']==0]

    v2=q1[q1['运行系数是否一致']==0]
    v3=q1[q1['暴露率是否一致']==0]
    v4=q1[q1['计划停运系数是否一致']==0]
    v5=q1[q1['非计划停运系数是否一致']==0]
    v6=q1[q1['强迫停运系数是否一致']==0]
    v7=q1[q1['计划停运率是否一致']==0]
    v8=q1[q1['非计划停运率是否一致']==0]
    v9=q1[q1['强迫停运率是否一致']==0]
    v10=q1[q1['连续可用小时是否一致']==0]
    v11=q1[q1['可用小时是否一致'] == 0]
    v12=q1[q1['运行小时是否一致'] == 0]
    v13=q1[q1['计划停运小时是否一致'] == 0]
    v14=q1[q1['强迫停运小时是否一致'] == 0]

    q=q1[q1['设备标识'].isin(v1['设备标识'])| q1['设备标识'].isin(v11['设备标识'])  | q1['设备标识'].isin(v2['设备标识'])| q1['设备标识'].isin(v12['设备标识']) | q1['设备标识'].isin(v3['设备标识']) | q1['设备标识'].isin(v4['设备标识'])  | q1['设备标识'].isin(v13['设备标识']) | q1['设备标识'].isin(v5['设备标识']) | q1['设备标识'].isin(v6['设备标识'])  | q1['设备标识'].isin(v14['设备标识']) | q1['设备标识'].isin(v7['设备标识']) | q1['设备标识'].isin(v8['设备标识']) | q1['设备标识'].isin(v9['设备标识']) | q1['设备标识'].isin(v10['设备标识'])]
     
    #读取运行事件   
    y = pd.read_excel(path_3)
        
    
    #保存问题清单excel
    file_name2=path_4+'\\'+par2
    writer2 = pd.ExcelWriter(file_name2, engine='xlsxwriter')
    sh2.to_excel(writer2,sheet_name='统计',index=False)
    y.to_excel(writer2,sheet_name='运行事件2019',index=False)
    q.to_excel(writer2,sheet_name='不一致问题筛选',index=False)
    
    workbook = writer2.book
    
    #设置excel格式并导出  
    fmt = workbook.add_format({'font_size': 10, 
                              'font_name': u'宋体',
                              'valign':'vcenter',
                              'align':'center'})
    
    color_fmt = workbook.add_format({'bg_color':'#FFC7CE'})
    percent_fmt = workbook.add_format({'num_format':'0.00%'})
    border_fmt = workbook.add_format({'border': 1})

    l=[3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36, 41, 46, 51]
    
    s1=writer2.sheets['不一致问题筛选']       
    for col_num, value in enumerate(q.columns.values.take(l)):
            s1.write(0, l[col_num], value, color_fmt)
    
    
    s2 = writer2.sheets['统计']
    
    s2.set_column('A:A', 15, fmt)
    s2.set_column('B:E', 10, fmt)
    
    s2.conditional_format('E2:E11',
                          {'type':'cell',
                          'criteria':'<',
                          'value': 1,
                          'format': color_fmt})
    
    
    s2.conditional_format('E2:E11',                  
                          {'type':'no_blanks',
                          'format': percent_fmt})
    
    
    s2.conditional_format('A1:E11',
                          {'type':'no_blanks', 
                          'format': border_fmt})
    
    
    writer2.save()
    print('已生成 问题清单')
    end = time()
    print('Running time: %s Seconds'%(end-start))


#开始执行
def discriminate():
    global result,par1,par2
    par1 = str(parameter1.get())
    par2 = str(parameter2.get())
    fun()    
    result.set('执行结束，结果已输出到指定路径！')
    
   
root = Tk()
#标题
root.title('输变电可靠性指标校验程序')
#不允许改变窗口大小
root.resizable(False, False)
root.focusmodel()
#定义字符串
path1 = StringVar()
path2 = StringVar()
path3 = StringVar()
path4 = StringVar()
parameter1 = StringVar()
parameter2 = StringVar()
result = StringVar()

#第零行
Label(root,text='新老系统单台指标计算结果比对',height = 2,font='Helvetic 15').grid(row = 0,columnspan=4)

#第一行
Label(root,text='老系统单台计算结果表格',height = 2,
      width=25,justify='right').grid(row = 1,column = 0)
Entry(root,textvariable = path1,width =35).grid(row = 1,column = 1)
Label(root,text='').grid(row = 1,column = 2)
Button(root,text ='文件选择',command = selectPath1).grid(row = 1,column = 3)
a = Label(root,text='').grid(row = 1,column = 4)

#第二行
Label(root,text='新系统单台计算结果表格',height = 2,
      width=25,justify='right').grid(row = 2,column = 0)
Entry(root,textvariable = path2,width =35).grid(row = 2,column = 1)
Label(root,text='').grid(row = 2,column = 2)
Button(root,text ='文件选择',command = selectPath2).grid(row = 2,column = 3)
Label(root,text='').grid(row = 2,column = 4)

#第三行
Label(root,text='运行事件表格',height = 2,
      width=25,justify='right').grid(row = 3,column = 0)
Entry(root,textvariable = path3,width =35).grid(row = 3,column = 1)
Label(root,text='').grid(row = 3,column = 2)
Button(root,text ='文件选择',command = selectPath3).grid(row = 3,column = 3)
Label(root,text='').grid(row = 3,column = 4)

#第四行
Label(root,text='比对结果输出路径：',height = 2,
      width=25,justify='right').grid(row = 4,column = 0)
Entry(root,textvariable = path4,width =35).grid(row = 4,column = 1)
Label(root,text='').grid(row = 4,column = 2)
Button(root,text ='路径选择',command = selectPath4).grid(row = 4,column = 3)
Label(root,text='').grid(row = 4,column = 4)

#第五行
Label(root,text='重复及关联失败表格命名：',height = 2,
      width=25,justify='right').grid(row = 5,column = 0)
Entry(root,textvariable = parameter1,width =35).grid(row = 5,column = 1)
parameter1.set('疑似重复及关联失败清单.xlsx')

#第六行
Label(root,text='不一致表格命名：',height = 2,
      width=25,justify='right').grid(row = 6,column = 0)
Entry(root,textvariable = parameter2,width =35).grid(row = 6,column = 1)
parameter2.set('问题清单.xlsx')

#第七行
result_str = Label(root,textvariable=result,justify='left').grid(row = 7,column = 1)
Button(root,text ='开始执行',command = discriminate, justify ='center').grid(row = 7,column = 3)

#第八行
Label(root,text='',height = 1).grid(row = 8,column = 0)

root.mainloop()



