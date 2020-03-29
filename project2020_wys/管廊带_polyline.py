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

xlspath_coordinate = r"D:\智能化管线项目\雨污水管线数据处理\部分高架线路0624\北区含油污B线水高架\7-管架(墩）探查表  北区B管线.xls".decode('utf-8')
xlspath_attribute = r"D:\智能化管线项目\雨污水管线数据处理\部分高架线路0624\北区含油污B线水高架\6-管廊（带）探查表.xls".decode('utf-8')
arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\JLSHYWS_0624\JLSHYWSb线.gdb'
featuredata = r'JLYWS\PipeGallery'

# template = r'D:\智能化管线项目\雨污水管线数据处理\JLSHYWS_0621ys\JLSHYWS.gdb\JLYWS\ControlPoint'


data_coordinate = xlrd.open_workbook(xlspath_coordinate)
table_coor = data_coordinate.sheet_by_index(0)
sheetname_coor = data_coordinate.sheet_names()[0]

nrows_coor = table_coor.nrows
ncols_coor = table_coor.ncols

data_attr = xlrd.open_workbook(xlspath_attribute)
table_attr = data_attr.sheet_by_index(0)
sheetname_attr = data_attr.sheet_names()[0]

nrows_attr = table_attr.nrows
ncols_attr = table_attr.ncols

# des = arcpy.Describe(template)
# sparef = des.spatialReference


def getColumnIndex(table, columnName):
    columnIndex = None
    for i in range(table.ncols):
        if table.cell_value(0, i) == columnName:
            columnIndex = i
            break
    return columnIndex


def getRowIndex(table, RowName,colfld):
    RowIndex = None
    for i in range(table.nrows):
        #for j in range(table.ncols):
        if (table.cell_value(i, getColumnIndex(table,colfld))).upper() == RowName.upper():
            RowIndex = i
            break
    return RowIndex


def create_polyline_geometry(cur,ncols_attr,X,Y,Z,RowNameS,RowNameE,fldname):
    point = arcpy.Point()
    array = arcpy.Array()
    print getRowIndex(table_coor,RowNameS,"PipeRackName"), getRowIndex(table_coor,RowNameE,"PipeRackName")+1
    for i in range(getRowIndex(table_coor,RowNameS,"PipeRackName"), getRowIndex(table_coor,RowNameE,"PipeRackName")+1):

        point.X = round(float(table_coor.cell(i, getColumnIndex(table_coor,X)).value),8)
        point.Y = round(float(table_coor.cell(i, getColumnIndex(table_coor,Y)).value),8)
        point.Z = round(float(table_coor.cell(i, getColumnIndex(table_coor,Z)).value),8)
        #array.add(point)
        array.add(point)
    #print array.getObject(8).X

    #row = cur.newRow()
    polyline = arcpy.Polyline(array)
   # row.shape = polyline
    # for fldn in fldname:
    #     print fldn
    #     # k = getRowIndex(table_attr, RowNameS,'PipeCorridSName')
    #     # print "行:%s" % k
    #     for j in range(0, ncols_attr):
    #         if fldn == (str(table_attr.cell(0, j).value).strip()).upper():
    #             print fldn, table_attr.cell(0, j).value
    #             # print table.cell(i, j).ctype
    #             k = getRowIndex(table_attr, RowNameS,'PipeCorridSName')
    #             print "行:%s" % k
    #             if table_coor.cell(k, j).ctype == 3:
    #                 # print table.cell(i, j).ctype
    #                 date = xlrd.xldate_as_tuple(table_attr.cell(k, j).value, 0)
    #                 # print(date)
    #                 tt = datetime.datetime.strftime(datetime.datetime(*date), "%Y-%m-%d")
    #                 try:
    #                     row.setValue(fldn, tt)
    #                 except Exception as e:
    #                     print e
    #             else:
    #                 try:
    #                     row.setValue(fldn, table_attr.cell(k, j).value)
    #                 except Exception as e:
    #                     print e

    cur.insertRow([polyline])
    array.removeAll()

def main(featuredata):
    cur = None
    try:
        cur = arcpy.da.InsertCursor(featuredata,["SHAPE@"])
        fldname = [fldn.name for fldn in arcpy.ListFields(featuredata)]

        create_polyline_geometry(cur,ncols_attr,"X1","Y1","Z1",'BGJ1','18JB1',fldname)
       # create_polyline_geometry(cur,ncols_attr,"X2","Y2","Z2",'BGJ1','BGJ3',fldname)

        # create_polyline_geometry(cur, ncols_attr, "X1", "Y1", "Z1", '20JB80', '20J2', fldname)
        # create_polyline_geometry(cur, ncols_attr, "X2", "Y2", "Z2", '20JB80', '20J2', fldname)

    except Exception as e:
        print e
        arcpy.AddMessage(e)
    finally:
        if cur:
            del cur


if __name__ == "__main__":
    main(featuredata)
    #getRowIndex(table_attr,"BGJ1")