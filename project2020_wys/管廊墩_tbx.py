#!/usr/bin/env python
# -*- coding:cp936 -*-

# ���ݵ��������� point feature
# ���걣��8λС��
import arcpy
import sys
import xlrd
import os
import datetime

# reload(sys)
# sys.setdefaultencoding('utf-8')

xlspath = arcpy.GetParameterAsText(0) #r"D:\���ܻ�������Ŀ\����ˮ�������ݴ���\ԭʼ�ļ�\�Ͻ�����0730\�ܳ����Ӹ߼���·\4��ˮ���Ά��ˮ\7-���¹ܼ�(�գ�̽���ģ��0704.xls".decode('utf-8')
arcpy.env.workspace = arcpy.GetParameterAsText(1) #r'D:\���ܻ�������Ŀ\����ˮ�������ݴ���\�ύ�ɹ�\JLSHYWS0730_4��ˮ��.gdb'
featuredata = r'JLYWS\PipeRack'

# template = r'D:\���ܻ�������Ŀ\����ˮ�������ݴ���\JLSHYWS_0621ys\JLSHYWS.gdb\JLYWS\ControlPoint'


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


def create_point_geometry(cur,nrows,ncols,X,Y,Z,fldname):
    for i in range(1, nrows):

        l = table.cell(i, getColumnIndex(table,X)).value
        b = table.cell(i, getColumnIndex(table,Y)).value
        h = table.cell(i, getColumnIndex(table,Z)).value
        if l == "":
            continue
        point = arcpy.Point(round(float(l), 8), round(float(b), 8), float(h))
        row = cur.newRow()
        row.shape = point
        for fldn in fldname:
            for j in range(0, ncols):
                if fldn == (str(table.cell(0, j).value).strip()).upper():
                    print fldn, table.cell(0, j).value
                    arcpy.AddMessage(fldn+","+ table.cell(0, j).value)
                    # print table.cell(i, j).ctype
                    if table.cell(i, j).ctype == 3:
                        # print table.cell(i, j).ctype
                        date = xlrd.xldate_as_tuple(table.cell(i, j).value, 0)
                        # print(date)
                        tt = datetime.datetime.strftime(datetime.datetime(*date), "%Y-%m-%d")
                        try:
                            row.setValue(fldn, tt)
                        except Exception as e:
                            # print e
                            arcpy.AddError(e)
                    else:
                        try:
                            row.setValue(fldn, table.cell(i, j).value)
                        except Exception as e:
                            # print e
                            arcpy.AddError(e)

        cur.insertRow(row)


def main(featuredata):
    cur = None
    try:
        cur = arcpy.InsertCursor(featuredata)
        fldname = [fldn.name for fldn in arcpy.ListFields(featuredata)]

        create_point_geometry(cur,nrows,ncols,"x1","y1","z1",fldname)
        create_point_geometry(cur,nrows,ncols,"x2", "y2", "z2", fldname)

    except Exception as e:
        print e
        arcpy.AddError(e)
    else:
        # print "���{}��⣬�������{}��".format(os.path.basename(xlspath),nrows-1)
        arcpy.AddMessage(u"���{0}��⣬�������{1}��".format(os.path.basename(xlspath),nrows-1))

    finally:
        if cur:
            del cur


if __name__ == "__main__":
    main(featuredata)