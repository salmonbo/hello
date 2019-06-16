'''
Created on 2019年4月18日

@author: w00390037
'''
       
from openpyxl import Workbook
from openpyxl import load_workbook
from pickle import ADDITEMS
wb = Workbook()    #创建文件对象

# grab the active worksheet
wb = load_workbook(filename=r'C:\Users\MagicBook\Desktop\9800.xlsx')   # 打开22.xlsx从里面读数据
ws = wb.active     #获取第一个sheet
#ws=wb.get_active_sheet()
sheet0 = wb.create_sheet("2019")  

rows1 = ws.rows
RowIndex =0
for aol in rows1:
    RowIndex = RowIndex + 1
    content1 = []
    for x in aol:
        a = x.value
        if a == '编码':
           a = "编码(软件版本)" 
        elif a == 'ASD':
           a = "ASD(实际发货日期)"
        elif a == 'CPD':
           a = "CPD(承诺交单日期)"
        content1.append(a)       
    if RowIndex == 1:
        content1.extend(['version', 'IsMain', 'Date', 'Site', 'Region'])
    print (content1)
    sheet0.append(content1)    #一次写一行
# Data can be assigned directly to cells
#ws['A1'] = 42      #写入数字
#ws['B1'] = "你好"+"automation test" #写入中文（unicode中文也可）

# Rows can also be appended

wb.save('C:\\Users\MagicBook\Desktop\9800.xlsx')

# Python types will automatically be converted
import datetime
import time
ws['A2'] = datetime.datetime.now()    #写入一个当前时间
#写入一个自定义的时间格式
#ws['A3'] =time.strftime("%Y年%m月%d日 %H时%M分%S秒",time.localtime())

ws['A3'] =time.strftime('%Y{y}%m{m}%d{d} %H{h}%M{f}%S{s}').format(y='年',m='月',d='日',h='时',f='分',s='秒')


ws1 = wb.create_sheet("Mysheet")           #创建一个sheet
ws1.title = "New Title"                    #设定一个sheet的名字
ws2 = wb.create_sheet("Mysheet", 0)      #设定sheet的插入位置 默认插在后面
ws2.title = u"你好"    #设定一个sheet的名字 必须是Unicode
ws1.sheet_properties.tabColor = "1072BA"   #设定sheet的标签的背景颜色

