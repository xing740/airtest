# -*- encoding=utf8 -*-
__author__ = "Administrator"

#excel
import xlrd
import xlwt
from xlutils.copy import copy

from airtest.core.api import *

auto_setup(__file__)



from poco.drivers.android.uiautomation import AndroidUiautomationPoco
poco = AndroidUiautomationPoco(use_airtest_input=True, screenshot_each_action=False)

excelTitle = [["QQ号", "姓名", "账号", "性别", "生日", "组织"]]
excelName = "爬虫数据.xls" 

def createExcel(path, sheet_name, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.write(i, j, value[i][j])  # 像表格中写入数据（对应的行和列）
    workbook.save(path)  # 保存工作簿
    print("xls格式表格写入数据成功！")
    
def excelAppend(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i+rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
    print("xls格式表格【追加】写入数据成功！")

def getExcelAllName(path, arr):
    workbook = xlrd.open_workbook(path)
    worksheet = workbook.sheet_by_index(0)
    rows = worksheet.nrows
    for row in range(1, rows):
        arr.append(worksheet.cell_value(row, 1))

def dealAcount(path):
    w = xlrd.open_workbook(path)
    ws = w.sheet_by_index(0)
    rows = ws.nrows
    nw = copy(w)
    wws = nw.get_sheet(0)
    for row in range(1, rows):
        acc = ws.cell_value(row, 2)
        newAcc = ""
        for it in acc:
            if it.isdigit():
                newAcc+=it
        wws.write(row, 10, newAcc)
    nw.save(path)  
                

def openPeopleInfo(name):
    poco(text=name).click()
    poco("com.tencent.eim:id/right_icon").click()  
    poco("com.tencent.eim:id/bmqq_profile_card_title_rightview").click()
    if poco(text="更多资料").exists():
        poco("com.tencent.eim:id/formitem_right_textview").click()
     
def returnMainUI():
    while poco(text="返回").exists():
        poco(text="返回").click()
    poco("com.tencent.eim:id/ivTitleBtnLeft").click()

def getPeopleInfo():
    arr = []
    for it in poco("android.widget.GridView").offspring("android.view.View"):
        childs = it.children()
        if len(childs) < 2:
            continue
        arr.append(childs[1].get_name())
    arr.append(poco("好友资料").children()[2].get_name())
    allInfo.append(arr)
                                                       
    


allName = ["邢征亮"]
allInfo = []
#getExcelAllName(excelName, allName)   

while True:
    for title in poco("com.tencent.eim:id/troop_member_list").child("android.widget.FrameLayout").offspring("com.tencent.eim:id/tv_name"):       
        name = title.get_text();
        if name in allName:
            continue
        allName.append(name)
        openPeopleInfo(name)
        sleep(3)
        getPeopleInfo()
        returnMainUI()
    if len(allInfo) > 0:
        excelAppend(excelName, allInfo)
        allInfo.clear()
    poco("com.tencent.eim:id/troop_member_list").swipe([0.0185, -0.3558])
    sleep(5)

        
