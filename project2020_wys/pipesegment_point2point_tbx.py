# !/usr/bin/env python
# -*- coding:utf-8 -*-
import xlrd
import arcpy
import sys
import chardet
import datetime
import os
import glob

reload(sys)
sys.setdefaultencoding('utf-8')
arcpy.env.overwriteOutput = 1
arcpy.env.workspace = arcpy.GetParameterAsText(0) # r'D:\智能化管线项目\雨污水管线数据处理\提交成果\JLSHYWS0730_4污水汽.gdb'
# 单个excel文件
xlsPath = arcpy.GetParameterAsText(1) # r'D:\智能化管线项目\雨污水管线数据处理\原始文件\上交资料0730\总厂补加高架线路\4污水汽提净化水\7-最新管架(墩）探查表模板0704.xls'.decode(
    #'utf-8')

# xlsPath = r'D:\智能化管线项目\雨污水管线数据处理\原始文件\上交资料0722\总厂补加高架线路\2催化高盐水0716\7-最新管架(墩）探查表模板0704.xls'.decode('utf-8')
#print xlsPath

data = xlrd.open_workbook(xlsPath)
#data = xlrd.open_workbook(xlspath)
# global  table
# global nrows
# global ncols
table = data.sheets()[0]  # 通过索引顺序获取
cols = table.col_values(3)
nrows = table.nrows
ncols = table.ncols
# print ncols
#point = arcpy.Point()

array = arcpy.Array()
polygonGeometryList = []


def getColumnIndex(table, columnName):
    columnIndex = None
    for i in range(table.ncols):
        if table.cell_value(0, i) == columnName:
            columnIndex = i
            break
    return columnIndex


def segment_created():
    arcpy.SetProgressor("step","Creating PipeRackFrame...",0,nrows-1,1)
    # cur = arcpy.InsertCursor(r'JLYWS\PipeSegment')
    cur = arcpy.InsertCursor(r'JLYWS\piperack_polyline')

    # fldname = [field.name for field in arcpy.ListFields(r'JLYWS\PipeSegment')]
    fldname = [field.name for field in arcpy.ListFields(r'JLYWS\piperack_polyline')]
    for i in range(1, nrows):
        arcpy.SetProgressorLabel("Execute {} line".format(i))

        L1 = table.cell(i, getColumnIndex(table, "x1")).value
        B1 = table.cell(i, getColumnIndex(table, "y1")).value
        H1 = table.cell(i, getColumnIndex(table, "z1")).value

        L2 = table.cell(i, getColumnIndex(table, "x2")).value
        B2 = table.cell(i, getColumnIndex(table, "y2")).value
        H2 = table.cell(i, getColumnIndex(table, "z2")).value

        if '' in [L1, L2, B1, B2, H1, H2]:
            continue

        arcpy.AddMessage(str(L1) + "\t" + str(B1) + "\t" + str(H1))
        arcpy.AddMessage(str(L1) + "\t" + str(B1) + "\t" + str(H1))
        arcpy.AddMessage('\n')
        row = cur.newRow()
        array = arcpy.Array([arcpy.Point(round(float(L1), 8), round(float(B1), 8), float(
            H1)), arcpy.Point(round(float(L2), 8), round(float(B2), 8), float(H2))])

        #lineFeature = arcpy.Polyline(array)
        #row.shape = lineFeature
        row.shape = array

        # updates = arcpy.UpdateCursor(r'JLYWS\PipeSegment')
        # for cursor in updates:
        for fldn in fldname:
            for j in range(0, ncols):
                #print 112
                if fldn == (table.cell(0, j).value.strip()).upper():
                    # print table.cell(i, j).ctype
                    if table.cell(i, j).ctype == 3:
                        # print table.cell(i, j).ctype
                        date = xlrd.xldate_as_tuple(table.cell(i, j).value, 0)
                        # print(date)
                        tt = datetime.datetime.strftime(
                            datetime.datetime(*date), "%Y-%m-%d")
                        row.setValue(fldn, tt)
                    else:
                        try:
                            row.setValue(fldn, table.cell(i, j).value)
                        except Exception as e:
                            print e
                            arcpy.AddMessage(e)

        array.removeAll()
        cur.insertRow(row)
        arcpy.SetProgressorPosition()

    del cur
    arcpy.ResetProgressor()
    arcpy.AddMessage("Created {} piperackframe .".format(nrows-1))

if __name__ == "__main__":
    # tem_fun(xlsfilePath)
    segment_created()
