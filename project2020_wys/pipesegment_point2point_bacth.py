# !/usr/bin/env python
# -*- coding:utf-8 -*-
import xlrd
import arcpy
import sys
# import chardet
import datetime
import os
import glob
"""
批量生成管廊墩
"""
reload(sys)
sys.setdefaultencoding('utf-8')
arcpy.env.overwriteOutput = 1
arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\提交成果\JLSHYWS0806.gdb'
# 单个excel文件
# xlsPath = r'D:\智能化管线项目\雨污水管线数据处理\原始文件\中转事故水0716\中转事故水222222\2-管段探查表.xlsx'.decode('utf-8')
# 包含多个excel的文件夹
xlsfilePath = r"D:\智能化管线项目\雨污水管线数据处理\原始文件\上交资料0806\管廊带（路）\65号路管廊带探查表".decode('utf-8')
# target_excel = None


def iter_folder(xlsfilePath):

    for path, file, filename in os.walk(xlsfilePath):
        # print path,filename
        target_excels = glob.glob(path + os.sep + "*管架(墩）探查表*.xls")
        for target_excel in target_excels:
            read_excel(target_excel)
            segment_created(table, nrows, ncols, target_excel)


def read_excel(xlsPath):
    # xlsPath = r'D:\智能化管线项目\雨污水管线数据处理\原始文件\0710上交\事故水22路\7-最新管架(墩）探查表模板0704.xls'.decode('utf-8')
    #print xlsPath
    data = xlrd.open_workbook(xlsPath)
    #data = xlrd.open_workbook(xlspath)
    global table
    global nrows
    global ncols
    table = data.sheets()[0]  # 通过索引顺序获取
    cols = table.col_values(3)
    nrows = table.nrows
    ncols = table.ncols
    # print ncols
    #point = arcpy.Point()


def getColumnIndex(table, columnName):
    columnIndex = None
    for i in range(table.ncols):
        if table.cell_value(0, i) == columnName:
            columnIndex = i
            break
    return columnIndex


def segment_created(table, nrows, ncols, xlspath):
    cal = 0
    cur = None
    try:
        # cur = arcpy.InsertCursor(r'JLYWS\PipeSegment')
        cur = arcpy.InsertCursor(r'JLYWS\piperack_polyline')

        # fldname = [field.name for field in arcpy.ListFields(r'JLYWS\PipeSegment')]
        fldname = [field.name for field in arcpy.ListFields(
            r'JLYWS\piperack_polyline')]
        for i in range(1, nrows):
            # array = arcpy.Array()
            # polygonGeometryList = []
            #global name,id
            # if i == 0:
            #     continue
            L1 = table.cell(i, getColumnIndex(table, "x1")).value
            B1 = table.cell(i, getColumnIndex(table, "y1")).value
            H1 = table.cell(i, getColumnIndex(table, "z1")).value

            L2 = table.cell(i, getColumnIndex(table, "x2")).value
            B2 = table.cell(i, getColumnIndex(table, "y2")).value
            H2 = table.cell(i, getColumnIndex(table, "z2")).value

            if '' in [L1, L2, B1, B2, H1, H2]:
                cal += 1
                continue

            print L1, B1, H1
            print L2, B2, H2
            print '\n'
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
                            date = xlrd.xldate_as_tuple(
                                table.cell(i, j).value, 0)
                            # print(date)
                            tt = datetime.datetime.strftime(
                                datetime.datetime(*date), "%Y-%m-%d")
                            row.setValue(fldn, tt)
                        else:
                            try:
                                row.setValue(fldn, table.cell(i, j).value)
                            except Exception as e:
                                print e

            array.removeAll()
            cur.insertRow(row)
    except Exception as e:
        print e

    else:
        print "完成{}入库，入库数据{}条".format(
            os.path.basename(xlspath), nrows - 1 - cal)

    finally:
        if cur:
            del cur


if __name__ == "__main__":
    iter_folder(xlsfilePath)
    # segment_created()
