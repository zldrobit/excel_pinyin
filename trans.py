#!/usr/bin/env python
# coding: utf-8
import xlrd, xlwt
from xlutils.copy import copy 
from pypinyin import pinyin, lazy_pinyin

#使用xlrd读取指定excel工作中的指定表格的值并返回
def excel_read(doc,table,x,y):
     data = xlrd.open_workbook(doc)
     table = data.sheet_by_name(table)
     return table.cell(x,y).value

#使用xlwt创建指定excel工作中的指定表格的值并保存
def excel_create(sheet,value):
     data = xlwt.Workbook()
     table = data.add_sheet(sheet)
     table.write(1,4,value)
     data.save('demo.xls')

#三个结合操作同一个excel
rb = xlrd.open_workbook(u'学硕拼音1519人.xlsx')
#管道作用
wb = copy(rb) 

for i in range(0,1):
	rs = rb.sheet_by_index(i) 
	ws = wb.get_sheet(i)
	for j in range(1, rs.nrows):
		cname = rs.cell(j, 0).value
		pinyin = lazy_pinyin(cname)
		pinyin2 = (pinyin[0] + ', ' + ''.join(pinyin[1:])).title()
		ws.write(j, 1, pinyin2)

wb.save(u'3.xlsx');
