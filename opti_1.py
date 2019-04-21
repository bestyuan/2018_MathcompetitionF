# -*- coding: utf-8 -*-

import xlrd
import datetime
import pandas as pd
from xlrd import xldate_as_tuple
import numpy as np
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif']=['SimHei']
Wide_body = [332, 333, '33E', '33H', '33L', 773]
Narrow_body = [319, 320, 321, 323, 325, 738, '73A', '73E', '73H', '73L']
#  DTDT 国内到达 航站楼 国内出发 航站楼
dict_time = {"DTDT":15,"DTDS":20,"DSDT":20,"DSDS":15,
            "DTIT":35,"DTIS":40,"DSIT":40,"DSIS":35,
            "ITDT":35,"ITDS":40,"ISDT":40,"ISDS":45,
            "ITIT":20,"ITIS":30,"ISIT":30,"ISIS":20}

def GetInfo_1(puck,gate):
    test1 = pd.DataFrame()
    x=[]
    y=[]
    for i in gate:
        for j in puck:
            if j["label"]==i["登机口"]:
                i["航班数量"]+=1
                i["航班"].append("到达航班:")
                i["航班"].append(j["到达航班"])
                i["航班"].append("出发航班:")
                i["航班"].append(j["出发航班"])
    test_gate = pd.DataFrame(data=gate)
    test1["登机口"]=test_gate['登机口']
    test1["航班数量"]=test_gate['航班数量']
    test1['航班']=test_gate['航班']
    # test1.to_csv("hangbandengjikou.csv", encoding="gb2312", index=None)
    # test1.plot(x='登机口',y='航班数量',kind="line")
    #登机口 与每个登机口数量折线图
    test1["平均使用率"]=test_gate["平均使用率"]
    for i in test1["登机口"]:
        x.append(i)
    x1 = range(len(x))
    for i in test1["航班数量"]:
        y.append(i)
    names=[]
    for name in range(len(x)):
        names.append(x[name])
    plt.plot(x1,y,marker="*",mec='b',mfc="w",label=u"每个登机口分配的航班数")
    plt.xlabel(u"登机口")
    plt.ylabel(u"航班数量")
    plt.legend()
    plt.xlim(1,69)
    plt.ylim(0,16)
    plt.xticks(x1,names,rotation=45)
    plt.show()
    #使用率的折线图
    plt.figure()
    x=[]
    y=[]
    for i in test1["登机口"]:
        x.append(i)
    for i in test1["平均使用率"]:
        y.append(i)
    names=[]
    for name in range(len(x)):
        names.append(x[name])
    plt.plot(x1,y,marker="*",mec='b',mfc="w",label=u"每个登机口的平均使用率")
    plt.xlabel(u"登机口")
    plt.ylabel(u"登机口平均使用率")
    plt.legend()
    plt.xlim(1,69)
    plt.ylim(0,1)
    plt.xticks(x1,names,rotation=45)
    plt.show()



    return
def GetInfo_kuanzhai(puck):
    w_success=0
    n_success=0
    w_failar=0
    n_failar=0
    list_kuan=[]
    list_zhai=[]
    l_k=[]
    l_z=[]
    for i in puck:
        if i["label"]!="0":
            if i["机体类别"]=="W":
                list_kuan.append(i["label"])
                w_success+=1
            if i["机体类别"]=="N":
                list_zhai.append(i["label"])
                n_success+=1
        else:
            if i["机体类别"]=="W":
                w_failar+=1
            if i["机体类别"]=="N":
                n_failar+=1
    l_k=set(list_kuan)
    l_z=set(list_zhai)
    print(len(l_k))
    print(len(l_z))
    print("w_success",w_success)
    print("n_success",n_success)
    print("w_failar",w_failar)
    print("n_failar",n_failar)

def GetInfo_ave_rate(puck,gate):
    for i in puck:
      t1 = i["到达登机口时刻"]
      t2 = i["驶离登机口时刻"]
      if t2<=287 or t1>576:
          continue
      elif t1<287 and t2>=576:
          i["登机口占用时间"]=575-287
      elif t1>=287 and t2<576:
          i["登机口占用时间"]=t2-t1
      elif t1>=287 and t1<576 and t2>576:
          i["登机口占用时间"]=575-t1
      elif t1<=287 and t2>287 and t2<576:
          i["登机口占用时间"]=t2-287
      for j in gate:
          if i["label"]==j["登机口"]:
              j["登机口使用时间"]+=i["登机口占用时间"]*5
    for k in gate:
        if k["登机口使用时间"]!=0:
            k["平均使用率"]=k["登机口使用时间"]/1440.0
    rate_csv = pd.DataFrame(data=gate)
    rate = pd.DataFrame()
    rate["登机口"]=rate_csv["登机口"]
    rate["登机口使用时间"]=rate_csv["登机口使用时间"]
    rate["平均使用率"]=rate_csv["平均使用率"]
    rate.to_csv("use_rate.csv", encoding="gb2312", index=None)


def GetFinal_result(puck):
    on_date = []
    on_time =[]
    out_date=[]
    out_time=[]
    for i in puck:
        on_date.append(i["到达时间"].strftime("%d-%b-%Y"))
        on_time.append(i["到达时间"].strftime("%H:%M"))
        out_date.append(i["出发时间"].strftime("%d-%b-%Y"))
        out_time.append(i["出发时间"].strftime("%H:%M"))
    test1 = pd.DataFrame()
    test = pd.DataFrame(data=puck)
    test1["飞机转场记录号"]=test["飞机转场记录号"]
    test1["到达日期"]=on_date
    test1["到达时刻"]=on_time
    test1["到达航班"]=test["到达航班"]
    test1["到达类型"]=test["到达类型"]
    test1["飞机型号"]=test["飞机型号"]
    test1["出发日期"]=out_date
    test1["出发时刻"]=out_time
    test1["出发航班"]=test["出发航班"]
    test1["出发类型"]=test["出发类型"]
    test1["上线机场"]=test["上线机场"]
    test1["下线机场"]=test["下线机场"]
    test1["问题一登机口"]=test["label"]
    test1.to_csv("problem_1.csv",index=None,encoding="gb2312")

def readHeaders(row):
    headers = []
    for header in row:
        headers.append(''.join(header.split('\n')))
    return headers


def readAllHeaders(workbook):
    sheet_names = workbook.sheet_names()
    allheaders = {}
    for sheet_name in sheet_names:
        sheet = workbook.sheet_by_name(sheet_name)
        allheaders[sheet_name] = readHeaders(sheet.row_values(0))
    return allheaders


def readAllSheets(workbook):
    sheet_names = workbook.sheet_names()
    allsheets = {}
    for sheet_name in sheet_names:
        sheet = workbook.sheet_by_name(sheet_name)
        header = readHeaders(sheet.row_values(0))
        allsheets[sheet_name] = []
        for i in range(1, sheet.nrows):
            new_row = {}
            for j in range(0, sheet.ncols):
                cell = sheet.cell(i, j).value
                ctype = sheet.cell(i, j).ctype
                if ctype == 3 and (header[j] == '到达日期' or header[j] == '出发日期'):
                    val = xldate_as_tuple(cell, workbook.datemode)
                elif ctype == 3 and (header[j] == '到达时刻' or header[j] == '出发时刻'):
                    val = xldate_as_tuple(cell, workbook.datemode)
                else:
                    val = cell
                new_row[header[j]] = val
            allsheets[sheet_name].append(new_row)
    return allsheets


def ReadFixedData():
    # _fixed版本修复了时间中的空格
    workbook = xlrd.open_workbook(u'InputData_fixed.xlsx')
    allheaders = readAllHeaders(workbook)
    allsheets = readAllSheets(workbook)

    PucksSheet = []
    for row in allsheets['Pucks']:
        new_tuple = row['到达日期'][0:3] + row['到达时刻'][3:6]
        row.pop('到达日期')
        row.pop('到达时刻')
        row['到达时间'] = datetime.datetime(*new_tuple)
        new_tuple = row['出发日期'][0:3] + row['出发时刻'][3:6]
        row.pop('出发日期')
        row.pop('出发时刻')
        row['出发时间'] = datetime.datetime(*new_tuple)
        row["label"] = "0"
        row["diff_time"] = row["出发时间"] - row["到达时间"]
        row["到达登机口时刻"]=0
        row["驶离登机口时刻"]=0
        row["登机口占用时间"]=0
        if row['飞机型号'] in Wide_body:
            row['机体类别'] = 'W'
        else:
            row['机体类别'] = 'N'
        date20 = datetime.datetime.strptime('2018-01-20', '%Y-%m-%d')
        date21 = datetime.datetime.strptime('2018-01-21', '%Y-%m-%d')
        if (row['到达时间'] < date21 and row['到达时间'] > date20) or \
                (row['出发时间'] < date21 and row['出发时间'] > date20):
            PucksSheet.append(row)

    TicketsSheet = []
    for row in allsheets['Tickets']:
        new_tuple = row['到达日期']
        row['到达时间'] = datetime.datetime(*new_tuple)
        row.pop('到达日期')
        new_tuple = row['出发日期']
        row['出发时间'] = datetime.datetime(*new_tuple)
        row.pop('出发日期')
        row["时间"]=0
        row["到达航班登机口"]="0"
        row["出发航班登机口"] = "0"
        row["换乘成功"]="0"
        for i in PucksSheet:
            if (i["到达航班"]==row["到达航班"]and i["到达时间"].day==row['到达时间'].day):
                row["到达类型"]=i["到达类型"]
            if (i["出发航班"]==row["出发航班"]and i["出发时间"].day==row['出发时间'].day):
                row["出发类型"]=i["出发类型"]

        date20 = datetime.datetime.strptime('2018-01-20', '%Y-%m-%d')
        if (row['到达时间'] == date20) or \
                (row['出发时间'] == date20):
            TicketsSheet.append(row)
    #    GatesSheet = allsheets['Gates']
    GatesSheet = []
    for row in allsheets['Gates']:
        row["label"] = "0"
        ##get data
        row["航班数量"] = 0
        row["窄体机"] = 0
        row["宽体机"] = 0
        row["航班数"]=0
        row["航班"]=[]
        row["登机口使用时间"]=0
        row["平均使用率"]=0
        ##get data end
        GatesSheet.append(row)

    return (PucksSheet, TicketsSheet, GatesSheet, allheaders)


def diff_on(list_a, inx):
    a = list_a["到达类型"]
    for i in a:
        if i == inx:
            return 1
    return 0


def diff_off(list_a, inx):
    a = list_a["出发类型"]
    for i in a:
        if i == inx:
            return 1
    return 0


def Shiyan_Sort(sheet):
    new_sheet = []
    for x in range(len(sheet)):
        min_ele = sheet[0]["到达时间"]
        min_num = 0
        for i in range(len(sheet)):
            if sheet[i]["到达时间"] < min_ele:
                min_ele = sheet[i]["到达时间"]
                min_num = i

        new_sheet.append(sheet[min_num])
        sheet.remove(sheet[min_num])
    return new_sheet
def Time_index():
    time_idx = []
    idx_time = []
    last_time = datetime.datetime.strptime('2018-01-19', '%Y-%m-%d')
    add_time = datetime.timedelta(minutes=5)
    count = 0
    for i in range(864):
        time_str = (last_time + add_time).strftime("%Y-%m-%d %H:%M:%S")
        temp_ele_ti = {time_str: count}
        temp_ele_it = {str(count): time_str}
        time_idx.append(temp_ele_ti)
        idx_time.append(temp_ele_it)
        count = count + 1
        last_time = last_time + add_time
    return (time_idx, idx_time)


def used_port1(one_list):
    return list(set(one_list))

if __name__=="__main__":
    Array=[]
    count =0
    count1=0
    for loopF in range(9,10):
        PucksSheet, TicketsSheet, GatesSheet, allheaders = ReadFixedData()
        Sorted_PucksSheet = Shiyan_Sort(PucksSheet)
        Time_idx,Idx_time = Time_index()

        arrive_plane = [] #每次到达的飞机
        temp_plane_in_airport = [] #临时机位中进入机场的飞机
        plane_in_airport = [] #正常进入机场中的飞机
        temp_postion=[] #临时机位中的飞机
        over_plane= [] #已经处理完的飞机
        temp_over_plane= [] #临时机位中进入机场的已经处理完的飞机
        used_port = [] #记录用过的登机口
        temp_arrive_plane= [] #临时机位中当前时刻要处理的飞机
        All_arrive_count = 0
        AA_count = 0
        fial_count = 0
        S_count =0
        T_count=0
        T_port =[]
        S_port =[]
        for T_port_num in range(29):
            a = []
            T_port.append(a)
        for S_port_num in range(42):
            a = []
            S_port.append(a)

        T_port_record = []
        S_port_record = []
        for T_port_num in range(29):
            a=[]
            T_port_record.append(a)
        for S_port_num in range(42):
            a = []
            S_port_record.append(a)

        for i in range(len(Idx_time)):   #计算占用每个登机口的时间 i的范围应该为287<=i<576
            # 检查临时机位上是否有需要起飞的飞机，安排失败

            fial_mark = []
            for ci in range(len(temp_postion)):
                fly_time_str = temp_postion[ci]["出发时间"].strftime("%Y-%m-%d %H:%M:%S")
                fly_time = 0
                for it in range(len(Time_idx)):
                    temp_dict_keys =str(Time_idx[it].keys())
                    temp_dict_keys = temp_dict_keys[12:31]
                    if temp_dict_keys == fly_time_str:
                        fly_time = Time_idx[it][fly_time_str]
                if fly_time ==0:
                    print("fly_time -Something wrong")
                if i>fly_time:
                    fial_mark.append(ci)
                    fial_count+=1
            for fm in range (len(fial_mark)):
                del temp_postion[fial_mark[fm] - fm]

            mark_del = []  # 记录飞走的飞机
            # 检查是否有登机口被释放(飞机起飞后45分钟)
            for xip in range(len(plane_in_airport)):
                # 前十个时间点不用看，不在考虑范围内
                if i > 10:
                    if Idx_time[i - 9][str(i - 9)] == plane_in_airport[xip]["出发时间"].strftime("%Y-%m-%d %H:%M:%S"):
                        # 找当前飞机对应的那个登机口
                        for j in range(len(GatesSheet)):
                            if GatesSheet[j]["登机口"]==plane_in_airport[xip]["label"]:
                                GatesSheet[j]["label"]="0"

                                plane_in_airport[xip]["驶离登机口时刻"]=i

                                over_plane.append(plane_in_airport[xip])
                                mark_del.append(xip)
                                # print(plane_in_airport[xip]["飞机转场记录号"]+"起飞")
            for md in range(len(mark_del)):
                del plane_in_airport[mark_del[md]-md]


            temp_mark_del = [] # 记录从临时位进入登机口后飞走的飞机
            #检查是否有登机口被释放（飞机起飞后45分钟）
            for xitp in range (len(temp_plane_in_airport)):

                if i >10:
                    if Idx_time[i-9][str(i-9)] == temp_plane_in_airport[xitp]["出发时间"].strftime("%Y-%m-%d %H:%M:%S"):
                    #找当前飞机对应的那个登机口
                        for j in  range(len(GatesSheet)):
                            if GatesSheet[j]["登机口"]==temp_plane_in_airport[xitp]["label"]:
                                GatesSheet[j]["label"]="0"

                                temp_plane_in_airport[xitp]['驶离登机口时刻']=i

                                temp_over_plane.append(temp_plane_in_airport[xitp])
                                temp_mark_del.append(xitp)
            for mdt in range(len(temp_mark_del)):
                del temp_plane_in_airport[temp_mark_del[mdt]-mdt]

            # 检查是否有飞机到达
            for xi in range(len(Sorted_PucksSheet)):
                if Idx_time[i][str(i)] == Sorted_PucksSheet[xi]["到达时间"].strftime("%Y-%m-%d %H:%M:%S"):
                    arrive_plane.append(Sorted_PucksSheet[xi])
                    All_arrive_count += 1
                pass
            pass
            already_temp_del=[]
        #检查当前时刻 临时机位中有没有飞机要出发

            for ict in range(len(temp_postion)):
                end_time_str = temp_postion[ict]["出发时间"].strftime("%Y-%m-%d %H:%M:%S")
                end_time =0
                for it in range(len(Time_idx)):
                    temp_dict_keys = str(Time_idx[it].keys())
                    temp_dict_keys=temp_dict_keys[12:31]
                    if temp_dict_keys ==end_time_str:
                        end_time = Time_idx[it][end_time_str]
                if end_time==0:
                    print("Something wrong")
                if (end_time-i)<loopF:
                    for j in range(len(GatesSheet)):
                        if GatesSheet[j]["label"]=="0": # 判断当前登机口是否空余
                            if GatesSheet[j]["到达类型"] == temp_postion[ict]["到达类型"] and GatesSheet[j]["出发类型"] ==temp_postion[ict]["出发类型"] \
                                and temp_postion[ict]["机体类别"]==GatesSheet[j]["机体类别"]:
                                temp_postion[ict]["label"] = GatesSheet[j]["登机口"]


                                temp_postion[ict]["到达登机口时刻"]=i


                                if GatesSheet[j]["登机口"][0]=="T":
                                    T_port_num = int(GatesSheet[j]["登机口"][1:])
                                    T_dic = {"进入时间":i,"班次":temp_postion[ict]["飞机转场记录号"]}
                                    T_port[T_port_num].append(T_dic)
                                elif GatesSheet[j]["登机口"][0]=="S":
                                    S_port_num = int(GatesSheet[j]["登机口"][1:])
                                    S_dic = {"进入时间":i,"班次":temp_postion[ict]["飞机转场记录号"]}
                                    S_port[S_port_num].append(S_dic)
                                temp_plane_in_airport.append(temp_postion[ict])
                                already_temp_del.append(ict)
                                #标记登机口被占用
                                GatesSheet[j]["label"]="1"
                                used_port.append(GatesSheet[j]["登机口"])
                                break
                            pass
                        pass
                    pass
                    if temp_postion[ict]["label"]!="0": #说明找到了
                        continue
                    for j in range(len(GatesSheet)):
                        if GatesSheet[j]["label"]=="0": #判断当前登机口是否空余
                            if diff_on(GatesSheet[j],temp_postion[ict]["到达类型"]) and diff_off(GatesSheet[j],temp_postion[ict]["出发类型"]) and temp_postion[ict]["机体类别"]==GatesSheet[j]["机体类别"]:
                                temp_postion[ict]["label"]=GatesSheet[j]["登机口"]


                                temp_postion[ict]["到达登机口时刻"]=i


                                if GatesSheet[j]["登机口"][0]=="T":
                                    T_port_num = int (GatesSheet[j]["登机口"][1:])
                                    T_dic = {"进入时间":i,"班次":temp_postion[ict]["飞机转场记录号"]}
                                    T_port[T_port_num].append(T_dic)
                                elif GatesSheet[j]["登机口"][0]=="S":
                                    S_port_num = int (GatesSheet[j]["登机口"][1:])
                                    S_dic = {"进入时间":i,"班次":temp_postion[ict]["飞机转场记录号"]}
                                    S_port[S_port_num].append(S_dic)
                                temp_plane_in_airport.append(temp_postion[ict])
                                already_temp_del.append(ict)
                                #标记登机口被占用
                                GatesSheet[j]["label"]="1"
                                used_port.append(GatesSheet[j]["登机口"])
                                break
                            pass
                        pass
                    pass
                pass
            pass

            for atd in range(len(already_temp_del)):
                del temp_postion[already_temp_del[atd]-atd]
            #处理到达的飞机
            for xa in range(len(arrive_plane)):
                count1+=1
                if arrive_plane[xa]["label"]=="0" or arrive_plane[xa]["label"]=="temp_position":
                    AA_count+=1
                    for j in range(len(GatesSheet)):
                        if GatesSheet[j]["label"]=="0": #判断当前登机口是否空余
                            #先找完全匹配的
                            if GatesSheet[j]["到达类型"] == arrive_plane[xa]["到达类型"] and GatesSheet[j]["出发类型"] ==arrive_plane[xa]["出发类型"] \
                                and arrive_plane[xa]["机体类别"]==GatesSheet[j]["机体类别"]:
                                arrive_plane[xa]["label"]=GatesSheet[j]["登机口"]


                                arrive_plane[xa]["到达登机口时刻"]=i

                                count+=1
                                if GatesSheet[j]["登机口"][0]=="T":
                                    T_port_num = int(GatesSheet[j]["登机口"][1:])
                                    T_dic = {"进入时间":i,"班次":arrive_plane[xa]["飞机转场记录号"]}
                                    T_port[T_port_num].append(T_dic)
                                elif GatesSheet[j]["登机口"][0]=="S":
                                    S_port_num = int(GatesSheet[j]["登机口"][1:])
                                    S_dic = {"进入时间":i,"班次":arrive_plane[xa]["飞机转场记录号"]}
                                    S_port[S_port_num].append(S_dic)

                                plane_in_airport.append(arrive_plane[xa])
                                #标记登机口被占用
                                GatesSheet[j]["label"]="1"
                                #记录用过的登机口
                                used_port.append(GatesSheet[j]["登机口"])
                                break
                            pass
                        pass
                    pass
                    if arrive_plane[xa]["label"]!="0" and arrive_plane[xa]["label"]!="temp_position": # 说明找到了
                        continue
                    for j in range(len(GatesSheet)):
                        if GatesSheet[j]["label"]=="0": #判断当前登机口是否空余
                            #非完全匹配
                            if diff_on(GatesSheet[j],arrive_plane[xa]["到达类型"]) and diff_off(GatesSheet[j],arrive_plane[xa]["出发类型"]) and arrive_plane[xa]["机体类别"]==GatesSheet[j]["机体类别"]:
                                arrive_plane[xa]["label"]=GatesSheet[j]["登机口"]
                                count+=1


                                arrive_plane[xa]["到达登机口时刻"]=i

                                if GatesSheet[j]["登机口"][0]=="T":
                                    T_port_num = int (GatesSheet[j]["登机口"][1:])
                                    T_dic = {"进入时间":i,"班次":arrive_plane[xa]["飞机转场记录号"]}
                                    T_port[T_port_num].append(T_dic)
                                elif GatesSheet[j]["登机口"][0]=="S":
                                    S_port_num = int (GatesSheet[j]["登机口"][1:])
                                    S_dic = {"进入时间":i,"班次":arrive_plane[xa]["飞机转场记录号"]}
                                    S_port[S_port_num].append(S_dic)
                                plane_in_airport.append(arrive_plane[xa])
                                #标记登机口被占用
                                GatesSheet[j]["label"]="1"
                                used_port.append(GatesSheet[j]["登机口"])
                                break
                            pass
                        pass
                    pass
                    #没找到可用的登机口，进入临时机位
                    if arrive_plane[xa]["label"]=="0":
                        for ti in range(len(temp_postion)):
                            if temp_postion[ti]["飞机转场记录号"]==arrive_plane[xa]["飞机转场记录号"]:
                                break
                        temp_postion.append(arrive_plane[xa])

                pass
            pass
            arrive_plane=[]
        pass
        dict_t =[len(over_plane),len(temp_over_plane),fial_count]
        Array.append(dict_t)

    pass

    countlable=0
    countll=0
    for i in Sorted_PucksSheet:
        if i["label"]!="0"and i["label"]!="temp_position":
            countlable+=1
        if i["label"]=="0":
            countll+=1

GetInfo_ave_rate(Sorted_PucksSheet,GatesSheet)
GetInfo_1(Sorted_PucksSheet,GatesSheet)
GetInfo_kuanzhai(Sorted_PucksSheet)
GetFinal_result(Sorted_PucksSheet)
print(countlable)
print(countll)

print(count1)
print(AA_count)
print(Array)
print(T_port)
print(S_port)
print(len(T_port))
print(len(S_port))
all_port = set(used_port)













