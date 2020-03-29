#!/usr/bin/env python
# -*- coding:utf-8 -*-

# 根据点坐标生成 point feature
# 坐标保留8位小数
import arcpy
import sys
import xlrd
import os
import datetime

reload(sys)
sys.setdefaultencoding('utf-8')

xlspath = r"D:\智能化管线项目\雨污水管线数据处理\原始文件\总厂数据0714\高架\3循环水表格\10-弯头探查表.xlsx".decode('utf-8')
arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\提交成果\JLSHYWS循环水0714.gdb'
# featuredata = r'JLYWS\ControlPoint'  # 层名修改项
# featuredata = r'JLYWS\Tee'  # 层名修改项
# featuredata = r'JLYWS\valve'  # 层名修改项
featuredata = r'JLYWS\Elbow'  # 层名修改项
# featuredata = r'JLYWS\Manhole'  # 层名修改项
# featuredata = r'JLYWS\FourLink'  # 层名修改项

# template = r'D:\智能化管线项目\雨污水管线数据处理\JLSHYWS_0621ys\JLSHYWS.gdb\JLYWS\ControlPoint'



data = xlrd.open_workbook(xlspath)
table = data.sheet_by_index(0)
sheetname = data.sheet_names()[0]

nrows = table.nrows
ncols = table.ncols

# des = arcpy.Describe(template)
# sparef = des.spatialReference


def getColumnIndex(table, columnName):
    columnIndex = None
    for i in range(table.ncols):
        if table.cell_value(0, i) == columnName:
            columnIndex = i
            break
    return columnIndex


def create_point(featuredata):
    cur = None
    try:
        cur = arcpy.InsertCursor(featuredata)
        fldname = [fldn.name for fldn in arcpy.ListFields(featuredata)]
        for i in range(1, nrows):
            L1 = table.cell(i, getColumnIndex(table,"X")).value
            B1 = table.cell(i, getColumnIndex(table,"Y")).value
            H1 = table.cell(i, getColumnIndex(table,"Z")).value
            row =cur.newRow()
            if L1:
                point = arcpy.Point(round(float(L1),8), round(float(B1),8), float(H1))
                row.shape = point
            for fldn in fldname:
                for j in range(0, ncols):
                    #print 112
                    if fldn == (str(table.cell(0, j).value).strip()).upper():
                        # print table.cell(i, j).ctype
                        if table.cell(i, j).ctype == 3:
                            # print table.cell(i, j).ctype
                            date = xlrd.xldate_as_tuple(table.cell(i, j).value, 0)
                            # print(date)
                            tt = datetime.datetime.strftime(datetime.datetime(*date), "%Y-%m-%d")
                            try:
                                row.setValue(fldn, tt)
                            except Exception as e:
                                print e

                        else:
                            try:
                               row.setValue(fldn, table.cell(i, j).value)
                            except Exception as e:
                                print e

            cur.insertRow(row)
    except Exception as e:
        print e
        arcpy.AddMessage(e)
    finally:
        if cur:

            del cur


if __name__ == "__main__":
    create_point(featuredata)