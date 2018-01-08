import xlrd , xlwt
from xlutils.copy import copy
'''
操作文件的时候注意文件的读写权限   有的电脑C盘需要修改权限  否则报错
同时操作一张表时  需要创建管道 

'''



def excel_read(doc, table, x, y):
    data = xlrd.open_workbook(doc)
    table = data.sheet_by_name(table)
    return table.cell(x, y).value


# 使用xlwt创建指定excel工作中的指定表格的值并保存

def excel_create(sheet, value):
    data = xlwt.Workbook()
    table = data.add_sheet(sheet)
    table.write(1, 4, value)
    data.save('demo.xls')


# 三个结合操作同一个excel
rb = xlrd.open_workbook('E:\\test.xls')

# 管道作用
wb = copy(rb)

# 通过get_sheet()获取的sheet有write()方法
ws = wb.get_sheet(0)  # 1代表是写到第几个工作表里，从0开始算是第一个。
ws.write(0, 0, 'changed!')
wb.save('E:\\test.xls')
