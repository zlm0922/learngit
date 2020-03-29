# !/usr/bin/env python
# -*- coding:utf-8 -*-
import xlrd
import arcpy
import sys
import time
import datetime

reload(sys)
sys.setdefaultencoding('utf-8')
arcpy.env.overwriteOutput = 1
arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\JLSHYWS_2_0618.gdb'
xlsPath = r'D:\智能化管线项目\雨污水管线数据处理\污水井线路表格0619\2-管段探查表 - 副本.xlsx'.decode('utf-8')
#print xlsPath
data = xlrd.open_workbook(xlsPath)
#data = xlrd.open_workbook(xlspath)
table = data.sheets()[0]  # 通过索引顺序获取
cols = table.col_values(3)
nrows = table.nrows
ncols = table.ncols
#print ncols
#point = arcpy.Point()

fldname = [field.name for field in arcpy.ListFields(r'JLYWS\PipeLine')]
flag = ''
cur = None
try:
    p = arcpy.Point()
    array = arcpy.Array()
    list_ref = []

    cur = arcpy.InsertCursor(r'JLYWS\PipeLine')
    for i in range(1, nrows):

        L1 = round(float(table.cell(i, 5).value), 8)
        B1 = round(float(table.cell(i, 6).value), 8)
        H1 = table.cell(i, 7).value
        ID1 = table.cell(i, 4).value

        L2 = round(float(table.cell(i, 12).value), 8)
        B2 = round(float(table.cell(i, 13).value), 8)
        H2 = table.cell(i, 14).value
        ID2 = table.cell(i,11).value
        if '' in [L1, L2, B1, B2, H1, H2]:
            continue
        #array.add(arcpy.Point(L2, B2, H2))
        if flag == '':
            flag = table.cell(i, 2).value.strip()
        if flag != table.cell(i, 2).value.strip():
            list_ref.sort(key=lambda p:p.ID)
            array = arcpy.Array(list_ref)
            row = cur.newRow()
            row.shape = array

            cur.insertRow(row)
            del list_ref[:]
            array.removeAll()
        list_ref.append(arcpy.Point(L1, B1, H1, ID=eval(ID1[3:])))
        list_ref.append(arcpy.Point(L2, B2, H2, ID=eval(ID2[3:])))

        # array.add(arcpy.Point(L1, B1, H1))
        # array.add(arcpy.Point(L2, B2, H2))

        flag = table.cell(i, 2).value.strip()
    # last feature
    list_ref.sort(key=lambda p: p.ID)
    array = arcpy.Array(list_ref)
    row = cur.newRow()
    row.shape = array

    cur.insertRow(row)
    array.removeAll()
    #del cur
except Exception as e:
    print e

finally:
    if cur:
        del cur
#     print L1, B1, H1
#     print L2, B2, H2
#     print '\n'
#     row = cur.newRow()
#     array.add([arcpy.Point(L1, B1, float(H1)),
#                          arcpy.Point(L2, B2, float(H2))])
#
#     #lineFeature = arcpy.Polyline(array)
#     #row.shape = lineFeature
#
#
#     # updates = arcpy.UpdateCursor(r'JLYWS\PipeSegment')
#     # for cursor in updates:
#     for fldn in fldname:
#         for j in range(0, ncols):
#             #print 112
#             if fldn == (table.cell(0, j).value.strip()).upper():
#                 #print table.cell(i, j).ctype
#                 if table.cell(i, j).ctype == 3:
#                    # print table.cell(i, j).ctype
#                     date = xlrd.xldate_as_tuple(table.cell(i, j).value, 0)
#                     #print(date)
#                     tt = datetime.datetime.strftime(datetime.datetime(*date),"%Y-%m-%d")
#                     row.setValue(fldn, tt)
#                 else:
#                     row.setValue(fldn, table.cell(i, j).value)
#                 #print fldn,table.cell(0, j).value
#                     # cursor.setValue(fldn, table.cell(i, j).value)
#                     # print table.cell(i, j).valueD
#                     # updates.updateRow(cursor)
#     array.removeAll()
#     cur.insertRow(row)
#
# del cur
#

