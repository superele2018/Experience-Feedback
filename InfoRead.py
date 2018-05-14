import xlrd
import xlwt
from xlwt import Workbook, easyxf
import math
import matplotlib.pyplot as plt
import os
import time
from pylab import *
from matplotlib.ticker import MultipleLocator
from matplotlib.ticker import FormatStrFormatter

data=xlrd.open_workbook('test.xls')
table=data.sheets()[0]

wbk = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = wbk.add_sheet('sheet 1',cell_overwrite_ok=True)


f = open("output.txt","w")



default = easyxf('font: name Arial;')
columsPurpose=['SUBJECT']
# columsPurpose=['LOER_NO','PLANT_CODE','PLANT_NAME','UNIT_CODE',	'UNIFORM_UNIT',	'REPORT_YEAR'	,'REPORT_SERIAL'	,'S_VER','EN_NO' ,	'EN_DATE',	'EN_SERIAL','SUBJECT',	'START_EVENT']
# columsPurpose=['LOER_NO','PLANT_CODE','PLANT_NAME','UNIT_CODE',	'UNIFORM_UNIT',	'REPORT_YEAR'	,'REPORT_SERIAL'	,'S_VER'	,'EN_NO' ,	'EN_DATE',	'EN_SERIAL',	'INES_LEVEL',	'SUBJECT',	'START_EVENT',
#                'EVENT_BEGIN'	,'EVENT_END'	,'REPORTER_TEL']
ncols=table.ncols
nrows=table.nrows
listTilte=table.row_values(0)


def txtwrite(file,data,purpose):
    for j in purpose:
        file.write(str(data[j])+' ')
    file.write('\n')

def set_style(name, height, bold=False):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = name  # 'Times New Roman'
    font.bold = bold
    font.color_index = 4
    font.height = height

    # borders= xlwt.Borders()
    # borders.left= 6
    # borders.right= 6
    # borders.top= 6
    # borders.bottom= 6

    style.font = font
    # style.borders = borders

    return style


def myXlwtWRow(isheet,index,data,purpose):
    for j in purpose:
        x=[]
        x.append(data[j])
        isheet.write_rich_text(index,j,x,default)

def Tiltle2index(title):
    return listTilte.index(title)


# txtwrite(f,listTilte,range(0,len(listTilte)))

# myXlwtWRow(sheet,0,listTilte,range(0,len(listTilte)))#录入title
# myXlwtWRow(sheet,1,table.row_values(500))#录入title

# j=1;
x=[]
purpose=[]
# columsPurpose=listTilte[:]

for i in range(0,len(columsPurpose)):
    purpose.append(Tiltle2index(columsPurpose[i]))

for i in range(0,nrows):#录入数据
    # if time.strptime(table.row_values(i)[Tiltle2index('EVENT_BEGIN')][0:10],"%Y-%m-%d")>time.strptime('2009-12-31',"%Y-%m-%d"):
    x=table.row_values(i)[:]
    txtwrite(f,x,purpose)
    # myXlwtWRow(sheet,1,table.row_values(i))
    # j=j+1
f.close()

for i in range(0,100):#录入数据
    # if time.strptime(table.row_values(i)[Tiltle2index('EVENT_BEGIN')][0:10],"%Y-%m-%d")>time.strptime('2009-12-31',"%Y-%m-%d"):
    x=table.row_values(i)[:]
    myXlwtWRow(sheet, i, x,purpose)
    sheet.flush_row_data()
    # myXlwtWRow(sheet,1,table.row_values(i))
    # j=j+1


wbk.save('output1.xls')
#
# print(table.row_values(1)[Tiltle2index('EVENT_BEGIN')])
# print(time.strptime(table.row_values(1)[Tiltle2index('EVENT_BEGIN')][0:10],"%Y-%m-%d"))
#
