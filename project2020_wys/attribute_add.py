# !/usr/bin/env python
# -*- coding:utf-8 -*-
import xlrd
import arcpy
import sys
import time


reload(sys)
sys.setdefaultencoding('utf-8')
arcpy.env.overwriteOutput = 1
arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\JLSHYWS_0621ys\JLSHYWS.gdb'
xlsPath = r'D:\智能化管线项目\雨污水管线数据处理\部分雨水表格数据0620\8-三通探查表.xlsx'.decode('utf-8')
featuredata = r'JLYWS\Tee'
#print xlsPath
data = xlrd.open_workbook(xlsPath)
#data = xlrd.open_workbook(xlspath)
table = data.sheets()[0]  # 通过索引顺序获取
cols = table.col_values(3)
nrows = table.nrows
ncols = table.ncols
print ncols
print nrows
#point = arcpy.Point()


def getColumnIndex(table, columnName):
    columnIndex = None
    for i in range(table.ncols):
        if(table.cell_value(0, i) == columnName):
            columnIndex = i
            break
    return columnIndex

start_time = time.clock()
try:

    cur = arcpy.UpdateCursor(featuredata)

    fldname = [field.name for field in arcpy.ListFields(featuredata)]

    for row in cur:

        for i in range(1, nrows):

            if row.getValue("POINTNUMBER") == table.cell(
                    i, getColumnIndex(table, "POINTNUMBER")).value:
                for j in range(0, ncols):
                    for fldn in fldname:
                        if fldn == "POINTNUMBER":
                            continue
                        #print fldn
                        if fldn.upper() == (table.cell(0, j).value).upper():
                            print fldn,table.cell(0, j).value
                            # print table.cell(i, j).value
                            if table.cell(i, j).value !='' and (row.getValue(fldn) is None
                                                                or row.getValue(fldn) == ''):

                                row.setValue(fldn, table.cell(i, j).value)
                                print table.cell(i, j).value
                                cur.updateRow(row)
    del cur

except Exception as e:
    print e
finished_time = time.clock()
elapsetime = finished_time - start_time
print"elapsetime %s" % elapsetime
# finally:
#     if cur:
#         del cur
#