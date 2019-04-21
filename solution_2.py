# -*- coding: utf-8 -*-
import time
import xlrd
import datetime
from xlrd import xldate_as_tuple
from datetime import timedelta
import pandas as pd
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif']=['SimHei']

Wide_body = [332, 333, '33E', '33H', '33L', 773]
Narrow_body = [319, 320, 321, 323, 325, 738, '73A', '73E', '73H', '73L']
#  DTDT 国内到达 T航站楼 国内出发 T航站楼
dict_time = {"DTDT":[15,0],"DTDS":[20,1],"DSDT":[20,1],"DSDS":[15,0],
            "DTIT":[35,0],"DTIS":[40,1],"DSIT":[40,1],"DSIS":[35,0],
            "ITDT":[35,0],"ITDS":[40,1],"ISDT":[40,1],"ISDS":[45,2],
            "ITIT":[20,0],"ITIS":[30,1],"ISIT":[30,1],"ISIS":[20,0]}


def GetInfo_1(puck,gate):
    test1 = pd.DataFrame()
    x=[]
    y=[]
    kuan=0
    zhai=0
    for i in gate:
        for j in puck:
            if j["label"]==i["登机口"]:
                i["航班数量"]+=1
                i["航班"].append("到达航班:")
                i["航班"].append(j["到达航班"])
                i["航班"].append("出发航班:")
                i["航班"].append(j["出发航班"])

    for i in gate:
        if i["航班数量"]!=0:
            if i["机体类别"]=="W":
                kuan+=1
            if i["机体类别"]=="N":
                zhai+=1
    print("被占用登机口宽体机数量",kuan)
    print("被占用登机口窄体机数量",zhai)


    test_gate = pd.DataFrame(data=gate)
    test1["登机口"]=test_gate['登机口']
    test1["航班数量"]=test_gate['航班数量']
    test1['航班']=test_gate['航班']
    test1.to_csv("hangbandengjikou_problem2.csv", encoding="gb2312", index=None)
    test1.plot(x='登机口',y='航班数量',kind="line")
    # 登机口 与每个登机口数量折线图
    test1["平均使用率"]=test_gate["平均使用率"]
    for i in test1["登机口"]:
        x.append(i)
    x1 = range(len(x))
    for i in test1["航班数量"]:
        y.append(i)
    names=[]
    for name in range(len(x)):
        names.append(x[name])
    plt.plot(x1,y,marker="*",mec='r',mfc="w",color="g",label=u"每个登机口分配的航班数")
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
    plt.plot(x1,y,marker="o",mec='r',mfc="w",color="g",label=u"每个登机口的平均使用率")
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
    W_T=0
    W_S=0
    N_T=0
    N_S=0
    list_kuan=[]
    list_zhai=[]
    l_k=[]
    l_z=[]
    for i in puck:
        if i["label"]!="0" and i["label"]!="temp_position":
            if i["机体类别"]=="W" and i["label"][0]=="T":
                W_T+=1
            if i["机体类别"] == "W" and i["label"][0]== "S":
                W_S += 1
            if i["机体类别"] == "N" and i["label"][0] == "T":
                N_T += 1
            if i["机体类别"]=="N"and i["label"][0] == "S":
                N_S+=1
        else:
            if i["机体类别"]=="W" and i["label"][0]=="T":
                w_failar+=1
            # if i["机体类别"] == "W" and i["终端厅"] == "S":
            #     w_failar += 1
            # if i["机体类别"] == "N" and i["终端厅"] == "T":
            #     w_failar += 1
            if i["机体类别"]=="N"and i["label"][0] == "S":
                n_failar+=1
    l_k=set(list_kuan)
    l_z=set(list_zhai)
    print(len(l_k))
    print(len(l_z))
    print("w_T",W_T)
    print("W_S",W_S)
    print("N_T",N_T)
    print("N_S",N_S)
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
      elif t1<287 and t2>=287 and t2<576:
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
    rate.to_csv("use_rate_problem2.csv", encoding="gb2312", index=None)



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
    test1["问题二登机口"]=test["label"]
    test1.to_csv("problem_2.csv",index=None,encoding="gb2312")


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
        count=0
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
                count+=1
            if (i["出发航班"]==row["出发航班"]and i["出发时间"].day==row['出发时间'].day):
                row["出发类型"]=i["出发类型"]
                count+=1
        if count!=2:
            continue
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
        row["航班"]=[]
        row["登机口使用时间"]=0
        row["平均使用率"]=0
        if row["终端厅"]=="S":
            row["sort"]=1
        else:
            row["sort"]=0
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

##add a class
class new (object):

    def __init__(self,plane,ticksheet,pucksSheet):  #Plane为一个列表
        self.plane=plane
        self.ticksheet=ticksheet
        self.pucksSheet=pucksSheet
    def count_time(self):
        list_a=[]
        t_consume1=0
        t_consume2 = 0
        land_leave = " "
        for i in self.ticksheet:
            if i["到达航班"]==self.plane[0]["到达航班"]and i['到达时间'].day==self.plane[0]["到达时间"].day:
                 list_a.append(i)
        for k in range(0,len(list_a)):
            people = list_a[k]["乘客数"]
            arr_type = list_a[k]["到达类型"]
            out_type = list_a[k]["出发类型"]
            for j in self.pucksSheet:
                if list_a[k]["出发航班"] ==j["出发航班"] and j['到达时间'].day == list_a[k]["到达时间"].day:
                    land_leave=j["label"]

            land_arr_T = "T"
            t_consume1+= self.time_consume(people,land_arr_T,land_leave,arr_type,out_type)
            land_arr_S = "S"
            t_consume2+= self.time_consume(people,land_arr_S,land_leave,arr_type,out_type)
        if t_consume1<t_consume2:
            # return land_arr_T,t_consume1
            return "T",t_consume1
        else:
            # return land_arr_S,t_consume2
            return "S",t_consume2

    def time_consume(self,people,land_arr,land_leave,arr_type,out_type):

        # if land_leave=="0" or land_leave=="temp_position":
        if land_arr=="T":
            return people*0.5*(dict_time[arr_type+"T"+out_type+"T"][0]+dict_time[arr_type+"T"+out_type+"S"][0])
        if land_arr=="S":
            return people*0.5*(dict_time[arr_type + "S" + out_type + "T"][0] + dict_time[arr_type + "S" + out_type + "S"][0])
        # else:
        #     if land_arr=="T":
        #         return people*(dict_time[arr_type+"T"+out_type+land_leave[0]][0])
        #     if land_arr=="S":
        #         return people*(dict_time[arr_type + "S" + out_type + land_leave[0]][0])


##add a class
# class new (object):
#
#     def __init__(self,plane,ticksheet,pucksSheet):  #Plane为一个列表
#         self.plane=plane
#         self.ticksheet=ticksheet
#         self.pucksSheet=pucksSheet
#     def count_time(self):
#         list_a=[]
#         t_consume1=0
#         t_consume2 = 0
#         for i in self.ticksheet:
#             if i["到达航班"]==self.plane[0]["到达航班"]and i['到达时间'].day==self.plane[0]["到达时间"].day:
#                  list_a.append(i)
#         for i in range(0,len(list_a)):
#             people = list_a[i]["乘客数"]
#             arr_type = list_a[i]["到达类型"]
#             out_type = list_a[i]["出发类型"]
#             for j in self.pucksSheet:
#                 if list_a[i]["出发航班"] ==j["出发航班"] and j['到达时间'].day == list_a[i]["到达时间"].day:
#                     land_leave=j["label"]
#             land_arr_T = "T"
#             t_consume1+= self.time_consume(people,land_arr_T,land_leave,arr_type,out_type)
#             land_arr_S = "S"
#             t_consume2+= self.time_consume(people,land_arr_S,land_leave,arr_type,out_type)
#         if t_consume1<t_consume2:
#             # return land_arr_T,t_consume1
#             return "T",t_consume1
#         else:
#             # return land_arr_S,t_consume2
#             return "S",t_consume2
#
#     def time_consume(self,people,land_arr,land_leave,arr_type,out_type):
#
#         if land_leave=="0" or land_leave=="temp_position":
#             if land_arr=="T":
#                 return people*0.5*(dict_time[arr_type+"T"+out_type+"T"][0]+dict_time[arr_type+"T"+out_type+"S"][0])
#             if land_arr=="S":
#                 return people*0.5*(dict_time[arr_type + "S" + out_type + "T"][0] + dict_time[arr_type + "S" + out_type + "S"][0])
#         else:
#             if land_arr=="T":
#                 return people*(dict_time[arr_type+"T"+out_type+land_leave[0]][0])
#             if land_arr=="S":
#                 return people*(dict_time[arr_type + "S" + out_type + land_leave[0]][0])



#判断换乘是否失败 时间不够则失败
#获取换乘成功人数
def GetTransferSuccess(tick,puck):

    for i in tick:
        for j in Sorted_PucksSheet:
            if j["到达航班"] == i["到达航班"] and j["到达时间"].day == i["到达时间"].day:
                i["到达航班登机口"] = j["label"]
                i["到达时间"] = j["到达时间"]
            if j["出发航班"] == i["出发航班"] and j["出发时间"].day == i["出发时间"].day:
                i["出发航班登机口"] = j['label']
                i["出发时间"] = j["出发时间"]
        # 换乘   换乘失败是指没有足够的换乘时间。停靠临时机位的旅客忽略不计
        # 有temp_position
        if i["到达航班登机口"] != "0" and i["出发航班登机口"] != "0" and i["到达航班登机口"] != "temp_position" and i[
            "出发航班登机口"] != "temp_position":

            Interval = i["出发时间"] - i["到达时间"]
            a_time = Interval.total_seconds() / 60
            # 乘客花费时间
            if dict_time[i["到达类型"] + i["到达航班登机口"][0] + i["出发类型"] + i["出发航班登机口"][0]][0] <= a_time:
                i["换乘成功"] = "yes"
            else:
                i["换乘成功"] = "no"
    # write time
    sum_time = 0
    peo = 0
    sum_people = 0
    sum_people_failar = 0
    sum_people += i["乘客数"]
    for i in TicketsSheet:
        peo += i["乘客数"]
        if i["换乘成功"] == "yes":
            a = Judge(i)
            i["时间"] = i["乘客数"] * a
            sum_time += i["时间"]
            sum_people += i["乘客数"]
        else:
            sum_people_failar += i["乘客数"]




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


if __name__ == "__main__":
    PucksSheet, TicketsSheet, GatesSheet, allheaders = ReadFixedData()
    Sorted_PucksSheet = Shiyan_Sort(PucksSheet)
    Time_idx, Idx_time = Time_index()
    count_WW = 0
    count_NN = 0
    start=time.time()
    c_l=0
    # add
    sum_time=0
    arrive_plane = []  # 每次到达的飞机
    plane_in_airport = []  # 机场中的飞机
    temp_postion = []  # 临时机位中的飞机
    over_plane = []  # 已经处理完的飞机
    used_port = []  # 记录用过的登机口

    All_arrive_count = 0
    All_temp_count = 0
    AA_count = 0
    fial_count = 0

    count_W = 0
    count_N_out = 0
    count_N_out2 = 0
    all_success=0
    for i in range(len(Idx_time)):
        # 检查临时机位上是否有需要起飞的飞机，安排失败
        fial_mark = []
        for ci in range(len(temp_postion)):
            if Idx_time[i][str(i)] == temp_postion[ci]["出发时间"].strftime("%Y-%m-%d %H:%M:%S"):
                # print(temp_postion[ci]["飞机转场记录号"] + " 安排失败")
                fial_mark.append(ci)
                fial_count += 1
        for fm in range(len(fial_mark)):
            del temp_postion[fial_mark[fm] - fm]

        mark_del = []  # 记录飞走的飞机
        # 检查是否有登机口被释放(飞机起飞后45分钟)
        for xip in range(len(plane_in_airport)):
            # 前十个时间点不用看，不在考虑范围内
            if i > 10:
                if Idx_time[i - 10][str(i - 10)] == plane_in_airport[xip]["出发时间"].strftime("%Y-%m-%d %H:%M:%S"):
                    # 找当前飞机对应的那个登机口

                    for j in range(len(GatesSheet)):
                        if GatesSheet[j]["登机口"] == plane_in_airport[xip]["label"]:
                            GatesSheet[j]["label"] = "0"
                            # print(GatesSheet[j]["登机口"] + " 释放")
                            plane_in_airport[xip]["驶离登机口时刻"] = i
                            # print(plane_in_airport[xip]["飞机转场记录号"] + " 起飞")
                            over_plane.append(plane_in_airport[xip])
                            mark_del.append(xip)
        for md in range(len(mark_del)):
            del plane_in_airport[mark_del[md] - md]
            # 检查是否有飞机到达
        for xi in range(len(Sorted_PucksSheet)):
            # 把临时机位中的飞机也看作是刚刚到达的飞机
            if len(temp_postion) != 0:
                for ti in range(len(temp_postion)):
                    arrive_plane.append(temp_postion[ti])
                    All_temp_count += 1
                temp_postion = []
            if Idx_time[i][str(i)] == Sorted_PucksSheet[xi]["到达时间"].strftime("%Y-%m-%d %H:%M:%S"):
                arrive_plane.append(Sorted_PucksSheet[xi])
                All_arrive_count += 1
                # print(Sorted_PucksSheet[xi]["飞机转场记录号"] + " 到达")
            pass
        pass
        # 处理到达的飞机
        for xa in range(len(arrive_plane)):

            if arrive_plane[xa]["label"] == "0" or arrive_plane[xa]["label"] == "temp_position":
                AA_count += 1

                # add time_consume:plane,ticksheet
                new_object = new(arrive_plane,TicketsSheet,Sorted_PucksSheet)
                _str,consumetime=new_object.count_time()    #_str is T or S
                if _str=="S":
                    GatesSheet = sorted(GatesSheet, key=lambda puck: puck["终端厅"])
                    c_l=1

                # sum_time+=consumetime
                # arrive_plane[xa]['label']=label
                for j in range(len(GatesSheet)):
                    if GatesSheet[j]["label"] == "0":  # 判断当前登机口是否空余
                        # 先找完全匹配的
                        if GatesSheet[j]["到达类型"] == arrive_plane[xa]["到达类型"] and GatesSheet[j]["出发类型"] == \
                                arrive_plane[xa]["出发类型"] and arrive_plane[xa]["机体类别"] == GatesSheet[j]["机体类别"]:

                            arrive_plane[xa]["label"] = GatesSheet[j]["登机口"]
                            arrive_plane[xa]["到达登机口时刻"] = i
                            plane_in_airport.append(arrive_plane[xa])
                            GatesSheet[j]["label"] = "1"
                            used_port.append(GatesSheet[j]["登机口"])
                            # print(GatesSheet[j]["登机口"] + " 占用")
                            break
                        pass
                pass
                if arrive_plane[xa]["label"] != "0" and arrive_plane[xa]["label"] != "temp_position":  # 说明找到了
                    continue
                for j in range(len(GatesSheet)):
                    if GatesSheet[j]["label"] == "0":  # 判断当前登机口是否空余
                        # 非完全匹配
                        if diff_on(GatesSheet[j], arrive_plane[xa]["到达类型"]) and diff_off(GatesSheet[j],
                                                                                         arrive_plane[xa]["出发类型"]) and \
                                arrive_plane[xa]["机体类别"] == GatesSheet[j]["机体类别"]:
                            arrive_plane[xa]["label"] = GatesSheet[j]["登机口"]
                            arrive_plane[xa]["到达登机口时刻"] = i
                            plane_in_airport.append(arrive_plane[xa])
                            GatesSheet[j]["label"] = "1"
                            used_port.append(GatesSheet[j]["登机口"])
                            # print(GatesSheet[j]["登机口"] + " 占用")
                            break
                        pass
                    pass
                pass
                # 没找到可用的登机口，进入临时机位
                if arrive_plane[xa]["label"] == "0" or arrive_plane[xa]["label"] == "temp_position":
                    for ti in range(len(temp_postion)):
                        if temp_postion[ti]["飞机转场记录号"] == arrive_plane[xa]["飞机转场记录号"]:
                            break
                    arrive_plane[xa]["label"] = "temp_position"
                    temp_postion.append(arrive_plane[xa])
                if c_l == 1:
                    GatesSheet = sorted(GatesSheet, key=lambda puck: puck["sort"])
                    check = 0
            pass
        pass
        arrive_plane = []
    pass
    all_port = used_port1(used_port)

    # write  gate in  tick
print("问题二运行时间",time.time()-start)



# GetInfo_ave_rate(Sorted_PucksSheet, GatesSheet)
# GetInfo_1(Sorted_PucksSheet, GatesSheet)
GetInfo_kuanzhai(Sorted_PucksSheet)
# GetFinal_result(Sorted_PucksSheet)




