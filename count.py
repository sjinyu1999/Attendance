import os
import time
import win32com.client as win32
#print(os.getcwd())
path=os.getcwd()
path1="\\09月汇总表"
fname = path+path1
#print(fname)
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)
# FileFormat = 51 is for .xlsx extension
wb.SaveAs(fname, FileFormat = 51)
# FileFormat = 56 is for .xls extension
# wb.SaveAs(fname[:-1], FileFormat = 56)
wb.Close()
excel.Application.Quit()
print("转换成功！！！！")
# 删除原格式的表“09月汇总表.xls”
# os.remove("09月汇总表.xls")
from openpyxl import load_workbook
wb = load_workbook(filename='09月汇总表.xlsx')
# print(wb.sheetnames)
ws = wb['刷卡记录']
# print(ws)
print('记录格式：姓名，日期，上班时间，下班时间，备注，工号，部门')
print('[格式："戚厚之", "2021-09-07", "09:27", "18:54", "", "64", ""]')
minrow = ws.min_row  # 最小行
maxrow = ws.max_row  # 最大行
mincol = ws.min_column  # 最小列
maxcol = ws.max_column  # 最大列
m = 3
n = 4
maxro=int(maxrow/2)
for i in range(minrow, maxro-1):
    # 录入姓名
    m += 2
    c1 = str(ws.cell(row=m, column=11).value)
    c1 ='"'+c1+'"'
    #print(c1)
    # print(c1)
    # 录入工号
    c5 = str(ws.cell(row=m, column=3).value)
    c5 = '"'+c5+'"'
    # print('"'+c5+'"')
    # 录入部门
    c6 = str(ws.cell(row=m, column=20).value)
    if(c6=="None"):
        c6='" "'
    else:
        c6='"'+c6+'"'
    # 录入上下班时间
    n += 2
    for j in range(mincol-1, maxcol):
        j = j+1
        # 录入日期
        c2 = ws.cell(row=4, column=j).value
        # 数值保留两位
        c2 = "{0:02d}".format(c2)
        c2 = "2021-09-"+str(c2)
        c2 = '"'+c2+'"'
        #print('"'+c2+'"')
        #录入时间
        c3 = str(ws.cell(row=n, column=j).value)
        #print(len(c3))
        if(len(c3)==4):
            #print("空")
            c3='" "'+', " "'
        elif(len(c3)==6):
            #print("上班时间："+c3[0:5])
            c3='"'+c3[0:5]+'"'+', " "'
        elif(len(c3)==12):
            c3='"'+c3[0:5]+'"'+', '+c3[6:11]
            #print("上班时间："+c3[0:5])
            #print("下班时间："+c3[6:11])
        elif(len(c3)==18):
            c3='"'+c3[0:5]+'"'+', '+c3[12:17]
            #print("上班时间："+c3[0:5])
            #print("下班时间："+c3[12:17])
        #print(c3)
        #print("姓名："+c1+",日期："+c2+","+goin+',备注：" "'+",工号："+c5+",部门："+c6)
        print(c1+", "+c2+", "+c3+', " "'+", "+c5+", "+c6)

