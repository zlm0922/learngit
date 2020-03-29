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

xlspath_coordinate = r"D:\智能化管线项目\雨污水管线数据处理\原始文件\上交资料0806\管廊带（路）\7-65管架(墩）探查表.xls".decode('utf-8')
xlspath_attribute = r"D:\智能化管线项目\雨污水管线数据处理\原始文件\上交资料0806\管廊带（路）\6-65号路管廊（带）探查表.xls".decode('utf-8')
arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\提交成果\JLSHYWS0806.gdb'
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
    else:
        print "请确认%s输入项是否正确"% RowName
    return RowIndex


def create_polyline_geometry(cur,ncols_attr,X,Y,Z,RowNameS,RowNameE,fldname):
    point = arcpy.Point()
    array = arcpy.Array()

    print getRowIndex(table_coor,RowNameS,"PipeRackName"), getRowIndex(table_coor,RowNameE,"PipeRackName")+1
    for i in range(getRowIndex(table_coor,RowNameS,"PipeRackName"), getRowIndex(table_coor,RowNameE,"PipeRackName")+1):

        try:
        # point.X = table_coor.cell(i, getColumnIndex(table_coor,X)).value
        # point.Y = table_coor.cell(i, getColumnIndex(table_coor,Y)).value
        # point.Z = table_coor.cell(i, getColumnIndex(table_coor,Z)).value
            if table_coor.cell(i, getColumnIndex(table_coor, X)).value:
                point.X = round(float(table_coor.cell(i, getColumnIndex(table_coor, X)).value), 8)
                point.Y = round(float(table_coor.cell(i, getColumnIndex(table_coor, Y)).value), 8)
                point.Z = round(float(table_coor.cell(i, getColumnIndex(table_coor, Z)).value), 8)
                #if point.X:
                array.add(point)
            else:break
        except Exception as eee:
            print eee
            print "请检查非英文字段行是否删除，坐标字段名称是否正确"
    row = cur.newRow()
    #polyline = arcpy.Polyline(array)
    row.shape = array
    for fldn in fldname:
        print fldn
        # k = getRowIndex(table_attr, RowNameS,'PipeCorridSName')
        # print "行:%s" % k
        for j in range(0, ncols_attr):
            if fldn == (str(table_attr.cell(0, j).value).strip()).upper():
                print fldn, table_attr.cell(0, j).value
                # print table.cell(i, j).ctype
                k = getRowIndex(table_attr, RowNameS,'PipeCorridSName')
                print "行:%s" % k
                if table_coor.cell(k, j).ctype == 3:
                    # print table.cell(i, j).ctype
                    date = xlrd.xldate_as_tuple(table_attr.cell(k, j).value, 0)
                    # print(date)
                    tt = datetime.datetime.strftime(datetime.datetime(*date), "%Y-%m-%d")
                    try:
                        row.setValue(fldn, tt)
                    except Exception as e:
                        print e
                else:
                    try:
                        row.setValue(fldn, table_attr.cell(k, j).value)
                    except Exception as e:
                        print e

    cur.insertRow(row)
    array.removeAll()


def main(featuredata):
    cur = None
    try:
        cur = arcpy.InsertCursor(featuredata)
        fldname = [fldn.name for fldn in arcpy.ListFields(featuredata)]
        for kk in range(1,table_attr.nrows):
            RowNameS = table_attr.cell(kk,getColumnIndex(table_attr,"PipeCorridSName")).value
            RowNameE = table_attr.cell(kk,getColumnIndex(table_attr,"PipeCorridEName")).value
            if RowNameE == "":
                continue


            create_polyline_geometry(cur,ncols_attr,"x1","y1","z1",RowNameS,RowNameE,fldname)
            create_polyline_geometry(cur,ncols_attr,"x2","y2","z2",RowNameS,RowNameE,fldname)

        # create_polyline_geometry(cur, ncols_attr, "X1", "Y1","Z1", "6-GJ63", '6-GJ85', fldname)
        # create_polyline_geometry(cur, ncols_attr, "X2", "Y2", "Z2", "6-GJ63", '6-GJ85', fldname)

    except Exception as e:
        print e
        arcpy.AddMessage(e)
    finally:
        if cur:
            del cur


if __name__ == "__main__":
    main(featuredata)
    #getRowIndex(table_attr,"BGJ1")