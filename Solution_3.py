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
dict_gate ={"T-North":0,"T-Center":1,"T-South":2,"S-North":3,"S-Center":4,"S-South":5,"S-East":6}
dict_gate_1 ={"T-North":1,"T-Center":2,"T-South":3,"S-North":4,"S-Center":5,"S-South":6,"S-East":0}
dict_gate_2 ={"T-North":2,"T-Center":3,"T-South":4,"S-North":5,"S-Center":6,"S-South":0,"S-East":1}
dict_gate_3 ={"T-North":3,"T-Center":4,"T-South":5,"S-North":6,"S-Center":0,"S-South":1,"S-East":2}
dict_gate_4 ={"T-North":4,"T-Center":5,"T-South":6,"S-North":0,"S-Center":1,"S-South":2,"S-East":3}
dict_gate_5 ={"T-North":5,"T-Center":6,"T-South":0,"S-North":1,"S-Center":2,"S-South":3,"S-East":4}
dict_gate_6 ={"T-North":6,"T-Center":0,"T-South":1,"S-North":2,"S-Center":3,"S-South":4,"S-East":5}
dict_sort ={"T-North":"sort","S-East":"sort1","S-South":"sort2","S-Center":"sort3","S-North":"sort4","T-South":"sort5","T-Center":"sort6",}
#行走时间
tr_time= { "T-North_T-North":10,"T-North_T-Center":15,"T-North_T-South":20,"T-North_S-North":25,
           "T-North_S-Center":20,"T-North_S-South":25,"T-North_S-East":25,
            "T-Center_T-Center":10,"T-Center_T-South":15,"T-Center_S-North":20,"T-Center_S-Center":15,
           "T-Center_S-South":20,"T-Center_S-East":20,
           "T-South_T-South":10,"T-South_S-North":25,"T-South_S-Center":20,"T-South_S-South":25,"T-South_S-East":25,
           "S-North_S-North":10,"S-North_S-Center":15,"S-North_S-South":20,"S-North_S-East":20,
           "S-Center_S-Center":10,"S-Center_S-South":15,"S-Center_S-East":15,
           "S-South_S-South":10,"S-South_S-East":20,
           "S-East_S-East":10,
            "T-Center_T-North":15,"T-South_T-North":20,"S-North_T-North":25,"S-Center_T-North":20,
            "S-South_T-North":25,"S-East_T-North":25,
            "T-South_T-Center":15,"S-North_T-Center":20,"S-Center_T-Center":15,
           "S-South_T-Center":20,"S-East_T-Center":20,
            "S-North_T-South":25,"S-Center_T-South":20,"S-South_T-South":25,"S-East_T-South":25,
            "S-Center_S-North":15,"S-South_S-North":20,"S-East_S-North":20,
            "S-South_S-Center":15,"S-East_S-Center":15,
            "S-East_S-South":20,
           }
#登机口区域
Gate_area=["T-North","T-Center","T-South","S-North","S-Center","S-South","S-East"]


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
    test1.to_csv("hangbandengjikou_problem3.csv", encoding="gb2312", index=None)
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
    plt.plot(x1,y,marker="o",mec='r',mfc="w",color="b",label=u"每个登机口分配的航班数")
    plt.xlabel(u"登机口")
    plt.ylabel(u"航班数量")
    plt.legend()
    plt.xlim(1,69)
    plt.ylim(0,25)
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
    plt.plot(x1,y,marker="o",mec='r',mfc="w",color="b",label=u"每个登机口的平均使用率")
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
        if i["label"] != "0" and i["label"] != "temp_position":
            if i["机体类别"] == "W" and i["label"][0] == "T":
                W_T += 1
            if i["机体类别"] == "W" and i["label"][0] == "S":
                W_S += 1
            if i["机体类别"] == "N" and i["label"][0] == "T":
                N_T += 1
            if i["机体类别"] == "N" and i["label"][0] == "S":
                N_S += 1
        else:
            if i["机体类别"] == "W" and i["label"][0] == "T":
                w_failar += 1
            # if i["机体类别"] == "W" and i["终端厅"] == "S":
            #     w_failar += 1
            # if i["机体类别"] == "N" and i["终端厅"] == "T":
            #     w_failar += 1
            if i["机体类别"] == "N" and i["label"][0] == "S":
                n_failar += 1
    l_k = set(list_kuan)
    l_z = set(list_zhai)
    print(len(l_k))
    print(len(l_z))
    print("w_T", W_T)
    print("W_S", W_S)
    print("N_T", N_T)
    print("N_S", N_S)

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
    rate.to_csv("use_rate_problem3.csv", encoding="gb2312", index=None)


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
    test1["问题三登机口"]=test["label"]
    test1.to_csv("problem_3.csv",index=None,encoding="gb2312")

def GetTensity(tick,puck,gate):
    people=0
    # allpeople=0
    for list_a in tick:
        for j in puck:
            if j["到达航班"] == list_a["到达航班"] and j["到达时间"].day == list_a["到达时间"].day:
                list_a["到达航班登机口"] = j["label"]
            if j["出发航班"] == list_a["出发航班"] and j["出发时间"].day == list_a["出发时间"].day:
                list_a["出发航班登机口"] = j['label']
        # 换乘   换乘失败是指没有足够的换乘时间。停靠临时机位的旅客忽略不计
        # 有temp_position
        if list_a["到达航班登机口"] != "0" and list_a["出发航班登机口"] != "0" and list_a["到达航班登机口"] != "temp_position" and list_a[
            "出发航班登机口"] != "temp_position":

            flow_time = dict_time[list_a["到达类型"] + list_a["到达航班登机口"][0] + list_a["出发类型"] + list_a["出发航班登机口"][0]][0]
            mrt_time = 16 * dict_time[list_a["到达类型"] + list_a["到达航班登机口"][0] + list_a["出发类型"] + list_a["出发航班登机口"][0]][1]
            gate_1 = "0"
            gate_2 = "0"
            for i in gate:
                if i["登机口"] == list_a["到达航班登机口"]:
                    gate_1 = i["区域"]
                if i["登机口"] == list_a["出发航班登机口"]:
                    gate_2 = i["区域"]

            if gate_1 != "0" and gate_2 != "0":
                travel_time = tr_time[list_a["到达航班登机口"][0]+"-" + gate_1 + "_" +
                                      list_a["出发航班登机口"][0] +"-" +gate_2]
            people+=list_a["乘客数"]
            list_a["换乘时间"] = flow_time + mrt_time + travel_time
        list_a['航班连接时间']=list_a['航班连接时间']/60
        list_a['紧张度']=list_a["换乘时间"]/list_a["航班连接时间"]
    ti = pd.DataFrame(data=tick)
    # result = pd.DataFrame()
    print("人数",people)
    ti.to_csv("transfer_rate.csv",encoding="gb2312")
    tens_rate = {0.1: 0, 0.2: 0, 0.3: 0, 0.4: 0, 0.5: 0, 0.6: 0, 0.7: 0, 0.8: 0, 0.9: 0, 1.0: 0}
    a = [0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9,1.0]
    for i in tick:
        if i["紧张度"]!=0:
            t=int( i['紧张度']/0.1)
            if t>=10:
                t=9
            tens_rate[a[t]]+=i["乘客数"]

    for i in range(0,10):
        if i ==0:
            tens_rate[a[1]]=tens_rate[a[1]]+tens_rate[a[0]]
        elif i ==9:
            tens_rate[a[i]]=tens_rate[a[i]]+tens_rate[a[i-1]]
        else:
            tens_rate[a[i+1]]=tens_rate[a[i+1]]+tens_rate[a[i]]
    for i in range(len(a)):
        tens_rate[a[i]]=tens_rate[a[i]]/people
    y = []
    for i in range(len(a)):
        y.append(tens_rate[a[i]])
    print("中转旅客比率",y)
    plt.bar(a, y, width=0.03,color="b", label=u"中转旅客比率分布图")
    plt.xlabel(u"紧张度")
    plt.ylabel(u"中转旅客比率")
    plt.legend()
    plt.xlim(0,1)
    plt.ylim(0,1)
    # plt.xticks(x1,names,rotation=45)
    plt.show()


    #统计换乘时间分布图
    zhongzhuan ={}
    for i in range(22):
        key=i*5
        zhongzhuan[key]=0
    tem =[]
    for j in range(22):
        tem.append(j*5)

    for i in tick:
        if i["换乘时间"]!=0:
            t=int( i['换乘时间']/5)
            if t>=21:
                t=21
            zhongzhuan[tem[t]]+=i["乘客数"]
    for i in range(0,22):
        if i ==0:
            zhongzhuan[tem[1]]=zhongzhuan[tem[1]]+zhongzhuan[tem[0]]
        elif i ==21:
            zhongzhuan[tem[i]]=zhongzhuan[tem[i]]+zhongzhuan[tem[i-1]]
        else:
            zhongzhuan[tem[i+1]]=zhongzhuan[tem[i+1]]+zhongzhuan[tem[i]]


    for i in range(len(tem)):
        zhongzhuan[tem[i]]=zhongzhuan[tem[i]]/people



    # test = pd.DataFrame(data=zhongzhuan,index=[0])
    # test.to_csv("data.csv",index=None)
    y = []
    for i in range(len(tem)):
        y.append(zhongzhuan[tem[i]])
    print("中转旅客比率",y)
    plt.bar(tem,y,width=3,color="b",label=u"旅客换乘时间分布图")
    plt.xlabel(u"换乘时间")
    plt.ylabel(u"中转旅客比率")
    plt.legend()
    plt.xlim(0,110)
    plt.ylim(0,1)
    # plt.xticks(x1,names,rotation=45)
    plt.show()







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
        count = 0
        new_tuple = row['到达日期']
        row['到达时间'] = datetime.datetime(*new_tuple)
        row.pop('到达日期')
        new_tuple = row['出发日期']
        row['出发时间'] = datetime.datetime(*new_tuple)
        row.pop('出发日期')
        row["时间"] = 0
        row["到达航班登机口"] = "0"
        row["出发航班登机口"] = "0"
        row["换乘成功"] = "0"
        row["换乘时间"]=0
        row["紧张度"]=0
        for i in PucksSheet:
            if (i["到达航班"] == row["到达航班"] and i["到达时间"].day == row['到达时间'].day):
                row["到达类型"] = i["到达类型"]
                span1 = i["到达时间"]
                count += 1
            if (i["出发航班"] == row["出发航班"] and i["出发时间"].day == row['出发时间'].day):
                row["出发类型"] = i["出发类型"]
                span2 = i["出发时间"]
                count += 1
        if count != 2:
            continue
        span=span2-span1
        _time = span.total_seconds()
        row["航班连接时间"]=_time
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
        row["sort"]=dict_gate[row["终端厅"]+'-'+row["区域"]]
        row["sort1"]=dict_gate_1[row["终端厅"]+'-'+row["区域"]]
        row["sort2"]=dict_gate_2[row["终端厅"]+'-'+row["区域"]]
        row["sort3"]=dict_gate_3[row["终端厅"]+'-'+row["区域"]]
        row["sort4"]=dict_gate_4[row["终端厅"]+'-'+row["区域"]]
        row["sort5"]=dict_gate_5[row["终端厅"]+'-'+row["区域"]]
        row["sort6"]=dict_gate_6[row["终端厅"]+'-'+row["区域"]]

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




#判断换乘是否失败 时间不够则失败
def transfer_time(list_a,gate):
    flow_time=dict_time[list_a["到达类型"]+list_a["到达航班登机口"][0]+list_a["出发类型"]+list_a["出发航班登机口"][0]][0]
    mrt_time=16*dict_time[list_a["到达类型"]+list_a["到达航班登机口"][0]+list_a["出发类型"]+list_a["出发航班登机口"][0]][1]
    gate_1="0"
    gate_2="0"
    for i in gate:
        if i["登机口"]==list_a["到达航班登机口"]:
            gate_1= i["区域"]
        if i["登机口"] == list_a["出发航班登机口"]:
            gate_2= i["区域"]

    if gate_1!="0" and gate_2!="0":
        travel_time = tr_time[list_a["到达航班登机口"][0]+gate_1+"-"+
                            list_a["出发航班登机口"][0]+gate_2]

    list_a["换乘时间"]=flow_time+mrt_time+travel_time

##add a class
class new (object):

    def __init__(self,plane,ticksheet,pucksSheet):  #Plane为一个列表
        self.plane=plane
        self.ticksheet=ticksheet
        self.pucksSheet=pucksSheet
    def count_time(self):
        list_a=[]

        Flight_connection_time=0
        time_dict={}
        for i in self.ticksheet:
            if i["到达航班"]==self.plane[0]["到达航班"]and i['到达时间'].day==self.plane[0]["到达时间"].day:
                 list_a.append(i)
        if len(list_a)==0:
            return 1
        for gate_area in Gate_area:
            count_people=0
            t_consume1=0
            for i in range(0,len(list_a)):  #debug 为空的话
                people = list_a[i]["乘客数"]
                arr_type = list_a[i]["到达类型"]
                out_type = list_a[i]["出发类型"]
                count_people+=1
                t_consume1+= self.time_consume(people,gate_area,arr_type,out_type)
                Flight_connection_time += list_a[i]["航班连接时间"]
            t_consume1=t_consume1/count_people
            Flight_connection_time= Flight_connection_time/len(list_a)
            Flight_connection_time = Flight_connection_time/60.0
            tensity = t_consume1/Flight_connection_time
            time_dict.update({gate_area:tensity})
        ten=min(time_dict, key=time_dict.get)
        return ten

    def time_consume(self,people,land_arr,arr_type,out_type):

        temp_arr = land_arr.split("-")
        Flow_time = 0.5*(
                dict_time[arr_type + temp_arr[0] + out_type + "T"][0] + dict_time[arr_type + temp_arr[0] + out_type + "S"][0])
        Mrt_time = 0.5*16*(
                dict_time[arr_type + temp_arr[0] + out_type + "T"][1] + dict_time[arr_type + temp_arr[0] + out_type + "S"][1])
        travel_time = 1/7*(tr_time[land_arr+"_"+Gate_area[0]]+tr_time[land_arr+"_"+Gate_area[1]]+tr_time[land_arr+"_"+Gate_area[2]]
        +tr_time[land_arr+"_"+Gate_area[3]]+ tr_time[land_arr + "_" + Gate_area[4]]+tr_time[land_arr+"_"+Gate_area[5]]+tr_time[land_arr+"_"+Gate_area[6]])

        return people *(Flow_time+Mrt_time+travel_time)





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

                            plane_in_airport[xip]["驶离登机口时刻"]=i
                            # print(GatesSheet[j]["登机口"] + " 释放")
                            # print(plane_in_airport[xip]["飞机转场记录号"] + " 起飞")
                            # get data
                            all_success+=1
                            GatesSheet[j]["航班数量"] += 1
                            if (plane_in_airport[xip]["机体类别"] == "W"):
                                count_W += 1
                                GatesSheet[j]["宽体机"] += 1
                            if (plane_in_airport[xip]["机体类别"] == "N"):
                                count_N_out += 1
                                GatesSheet[j]["窄体机"] += 1
                            # get data end
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
                _str=new_object.count_time()    #_str is T-North,T-Center etc...;

                if _str!=1:
                    # temp_str = _str.split("-")
                    GatesSheet= sorted(GatesSheet,key=lambda la:la[dict_sort[_str]])
                # sum_time+=consumetime
                # arrive_plane[xa]['label']=label

                for j in range(len(GatesSheet)):
                    if GatesSheet[j]["label"] == "0" :  # 判断当前登机口是否空余
                        # 先找完全匹配的
                        if GatesSheet[j]["到达类型"] == arrive_plane[xa]["到达类型"] and GatesSheet[j]["出发类型"] == \
                                arrive_plane[xa]["出发类型"] and arrive_plane[xa]["机体类别"] == GatesSheet[j]["机体类别"]:
                            arrive_plane[xa]["到达登机口时刻"]=i
                                # and GatesSheet[j]["终端厅"]==temp_str[0] and GatesSheet[j]["区域"]==temp_str[1]:
                            arrive_plane[xa]["label"] = GatesSheet[j]["登机口"]
                            plane_in_airport.append(arrive_plane[xa])
                            GatesSheet[j]["label"] = "1"
                            used_port.append(GatesSheet[j]["登机口"])
                            # print(GatesSheet[j]["登机口"] + " 占用")
                            break
                        pass
                    elif GatesSheet[j]["label"] == "0" :
                        if GatesSheet[j]["到达类型"] == arrive_plane[xa]["到达类型"] and GatesSheet[j]["出发类型"] == \
                                arrive_plane[xa]["出发类型"] and arrive_plane[xa]["机体类别"] == GatesSheet[j]["机体类别"]:
                            arrive_plane[xa]["到达登机口时刻"]=i
                            arrive_plane[xa]["label"] = GatesSheet[j]["登机口"]
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
                    if GatesSheet[j]["label"] == "0" :  # 判断当前登机口是否空余
                        # 非完全匹配
                        if diff_on(GatesSheet[j], arrive_plane[xa]["到达类型"]) and diff_off(GatesSheet[j],arrive_plane[xa]["出发类型"]) and \
                            arrive_plane[xa]["机体类别"] == GatesSheet[j]["机体类别"]:
                            arrive_plane[xa]["label"] = GatesSheet[j]["登机口"]
                            arrive_plane[xa]["到达登机口时刻"] = i
                            plane_in_airport.append(arrive_plane[xa])
                            GatesSheet[j]["label"] = "1"
                            used_port.append(GatesSheet[j]["登机口"])
                            # print(GatesSheet[j]["登机口"] + " 占用")
                            break
                        pass
                    elif  GatesSheet[j]["label"] == "0" :
                        if diff_on(GatesSheet[j], arrive_plane[xa]["到达类型"]) and diff_off(GatesSheet[j],arrive_plane[xa]["出发类型"]) and \
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
                    if GatesSheet[j]["label"] == "0" :  # 判断当前登机口是否空余
                        # 非完全匹配
                        if diff_on(GatesSheet[j], arrive_plane[xa]["到达类型"]) and diff_off(GatesSheet[j],
                                                                                         arrive_plane[xa]["出发类型"]) and \
                                arrive_plane[xa]["机体类别"] == GatesSheet[j]["机体类别"] :
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

            pass
        pass
        arrive_plane = []
    pass
    all_port = used_port1(used_port)

    # write  gate in  tick
    print("问题三运行时间:",time.time()-start)









# GetInfo_ave_rate(Sorted_PucksSheet, GatesSheet)
# GetInfo_1(Sorted_PucksSheet, GatesSheet)
GetInfo_kuanzhai(Sorted_PucksSheet)
# GetTensity(TicketsSheet,Sorted_PucksSheet,GatesSheet)
# GetFinal_result(Sorted_PucksSheet)