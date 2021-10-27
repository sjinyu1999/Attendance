import openpyxl
import random
from openpyxl import Workbook
from openpyxl import load_workbook
#生成一个 Workbook 的实例化对象，wb即代表一个工作簿（一个 Excel 文件）
wb = openpyxl.Workbook()
# 获取活跃的工作表，ws代表wb(工作簿)的一个工作表
ws = wb.active
#更改工作表ws的title
ws.title = 'output'
#对ws的单个单元格传入数据
ws['B2'] = random.random()
from openpyxl.styles import Alignment
#horizontal代表水平方向,vertical代表垂直方向
align = Alignment(horizontal='center',vertical='center')
ws['B2'].alignment = align
from openpyxl.styles import Font
ws['B2'].font = Font(bold=True,color="FF0000")
#ws.append([1,2,3])
# 获取名为'output'的sheet页
sheet1=wb['output']
#写入数据
j=2
for i in range(0,100):
    j+=i
    m=j+i
    #print(j,m)
    #合并单元格
    #print("正在合并单元格。。。")
    ws.merge_cells(start_row=j,start_column=2,end_row=m,end_column=2)
    ws.merge_cells(start_row=2,start_column=j,end_row=2,end_column=m)
    ws.cell(2,j).alignment = align
    ws.cell(j,2).alignment = align
    sheet1.cell(row=2,column=j,value=random.random())
    sheet1.cell(row=j,column=2,value=random.random())
#获取按行所有的数据
#按行读取
minrow=sheet1.min_row #最小行
maxrow=sheet1.max_row #最大行
mincol=sheet1.min_column #最小列
maxcol=sheet1.max_column #最大列
print('打印结果:')
for i in range(minrow,maxrow+1):
    for j in range(mincol,maxcol+1):
        cell=sheet1.cell(i,j).value
        #列
        column=sheet1.cell(i,j).column
        #行
        row=sheet1.cell(i,j).row
        #坐标
        coordinate=sheet1.cell(i,j).coordinate
        #print(i,j)
        #cell=str(cell)
        if (cell==None):
            continue
        else:
            print(cell)
    print()
wb.save('output.xlsx')
