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

arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\提交成果\JLSHYWS0730雨补.gdb'
# 单个excel文件
xlsPath = r'D:\智能化管线项目\雨污水管线数据处理\原始文件\上交资料0730\总厂雨水最新补测数据\2-管段探查表.xlsx'.decode('utf-8')


xlspath_coordinate = r"D:\智能化管线项目\雨污水管线数据处理\原始文件\上交资料0730\总厂雨水最新补测数据\3-中线控制点探查表.xlsx".decode('utf-8')
data_coordinate = xlrd.open_workbook(xlspath_coordinate)
table_coor = data_coordinate.sheet_by_index(0)
sheetname_coor = data_coordinate.sheet_names()[0]

# xlsPath = r'D:\智能化管线项目\雨污水管线数据处理\原始文件\0710上交\事故水22路\7-最新管架(墩）探查表模板0704.xls'.decode('utf-8')
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


def getRowIndex(table, RowName, colfld):
    try:
        RowIndex = None
        for i in range(table.nrows):
            # for j in range(table.ncols):
            if (table.cell_value(i, getColumnIndex(table, colfld))
                ).upper() == RowName.upper():
                RowIndex = i
                break
        else:
            print "请确认%s输入项是否正确" % RowName
            # arcpy.AddMessage("请确认%s输入项是否正确" % RowName)
        return RowIndex
    except Exception as e:
        print e


def segment_created():
    cal = 0
    cur = arcpy.InsertCursor(r'JLYWS\PipeSegment')
    # cur = arcpy.InsertCursor(r'JLYWS\piperack_polyline')

    fldname = [field.name for field in arcpy.ListFields(r'JLYWS\PipeSegment')]
    # fldname = [field.name for field in arcpy.ListFields(r'JLYWS\piperack_polyline')]
    for i in range(1, nrows):

        #global name,id
        # if i == 0:
        #     continue
        try:
            L1 = table_coor.cell(
                    getRowIndex(
                        table_coor, table.cell(
                            i, getColumnIndex(
                                table, "STARTPOINTNUMBER")).value, "POINTNUMBER"), getColumnIndex(
                        table_coor, "X")).value
            B1 = table_coor.cell(
                    getRowIndex(
                        table_coor, table.cell(
                            i, getColumnIndex(
                                table, "STARTPOINTNUMBER")).value, "POINTNUMBER"), getColumnIndex(
                        table_coor, "Y")).value
            H1 = table_coor.cell(
                    getRowIndex(
                        table_coor, table.cell(
                            i, getColumnIndex(
                                table, "STARTPOINTNUMBER")).value, "POINTNUMBER"), getColumnIndex(
                        table_coor, "Z")).value

            L2 = table_coor.cell(
                    getRowIndex(
                        table_coor, table.cell(
                            i, getColumnIndex(
                                table, "ENDPOINTNUMBER")).value, "POINTNUMBER"), getColumnIndex(
                        table_coor, "X")).value
            B2 = table_coor.cell(
                    getRowIndex(
                        table_coor, table.cell(
                            i, getColumnIndex(
                                table, "ENDPOINTNUMBER")).value, "POINTNUMBER"), getColumnIndex(
                        table_coor, "Y")).value
            H2 = table_coor.cell(
                    getRowIndex(
                        table_coor, table.cell(
                            i, getColumnIndex(
                                table, "ENDPOINTNUMBER")).value, "POINTNUMBER"), getColumnIndex(
                        table_coor, "Z")).value

            if '' in [L1, L2, B1, B2, H1, H2]:
                cal += 1
                continue

            print L1, B1, H1
            # arcpy.AddMessage(L1+","+B1+","+ H1)
            print L2, B2, H2
            # arcpy.AddMessage(L2 + "," + B2 + "," + H2)
            print '\n'
            # arcpy.AddMessage('\n')
            row = cur.newRow()
            array = arcpy.Array([arcpy.Point(round(float(L1),8), round(float(B1),8), float(H1)),
                                 arcpy.Point(round(float(L2),8), round(float(B2),8), float(H2))])

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
                            tt = datetime.datetime.strftime(datetime.datetime(*date), "%Y-%m-%d")
                            row.setValue(fldn, tt)
                        else:
                            try:
                                row.setValue(fldn, table.cell(i, j).value)
                            except Exception as e:
                                print e
                                arcpy.AddMessage(e)

                        # print fldn,table.cell(0, j).value
                        #print fldn,table.cell(0, j).value
                            # cursor.setValue(fldn, table.cell(i, j).value)
                            # print table.cell(i, j).value
                            # updates.updateRow(cursor)
            array.removeAll()
            cur.insertRow(row)
        except Exception as e:
            print e
            arcpy.AddMessage(e)

    del cur
    print "完成{}入库，总计入库{}条数据".format(os.path.basename(xlsPath), nrows - 1-cal)
    # arcpy.AddMessage("完成{}入库，总计入库{}条数据".format(os.path.basename(xlsPath), nrows - 1-cal))


if __name__ == "__main__":
    # tem_fun(xlsfilePath)
    segment_created()

