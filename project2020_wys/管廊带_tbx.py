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

xlspath_coordinate = arcpy.GetParameterAsText(0) # r"D:\���ܻ�������Ŀ\����ˮ�������ݴ���\ԭʼ�ļ�\�Ͻ�����0730\�ܳ����Ӹ߼���·\4��ˮ���Ά��ˮ\7-���¹ܼ�(�գ�̽���ģ��0704.xls".decode('utf-8')
xlspath_attribute = arcpy.GetParameterAsText(1)# r"D:\���ܻ�������Ŀ\����ˮ�������ݴ���\ԭʼ�ļ�\�Ͻ�����0730\�ܳ����Ӹ߼���·\4��ˮ���Ά��ˮ\6-���ȣ�����̽���.xls".decode('utf-8')
arcpy.env.workspace = arcpy.GetParameterAsText(2) #r'D:\���ܻ�������Ŀ\����ˮ�������ݴ���\�ύ�ɹ�\JLSHYWS0730_4��ˮ��.gdb'
featuredata = r'JLYWS\PipeGallery'

# template = r'D:\���ܻ�������Ŀ\����ˮ�������ݴ���\JLSHYWS_0621ys\JLSHYWS.gdb\JLYWS\ControlPoint'


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
        # print "��ȷ��%s�������Ƿ���ȷ"% RowName
        arcpy.AddMessage("��ȷ��%s�������Ƿ���ȷ"% RowName)
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
            # print eee
            # print "�����Ӣ���ֶ����Ƿ�ɾ���������ֶ������Ƿ���ȷ"
            arcpy.AddError(eee)
            arcpy.AddMessage("�����Ӣ���ֶ����Ƿ�ɾ���������ֶ������Ƿ���ȷ")
    row = cur.newRow()
    #polyline = arcpy.Polyline(array)
    row.shape = array
    for fldn in fldname:
        print fldn
        # k = getRowIndex(table_attr, RowNameS,'PipeCorridSName')
        # print "��:%s" % k
        for j in range(0, ncols_attr):
            if fldn == (str(table_attr.cell(0, j).value).strip()).upper():
                # print fldn, table_attr.cell(0, j).value
                arcpy.AddMessage(fldn + ","+table_attr.cell(0, j).value)
                # print table.cell(i, j).ctype
                k = getRowIndex(table_attr, RowNameS,'PipeCorridSName')
                # print "��:%s" % k
                # arcpy.AddMessage("�����Ӣ���ֶ����Ƿ�ɾ���������ֶ������Ƿ���ȷ")
                if table_coor.cell(k, j).ctype == 3:
                    # print table.cell(i, j).ctype
                    date = xlrd.xldate_as_tuple(table_attr.cell(k, j).value, 0)
                    # print(date)
                    tt = datetime.datetime.strftime(datetime.datetime(*date), "%Y-%m-%d")
                    try:
                        row.setValue(fldn, tt)
                    except Exception as e:
                        # print e
                        arcpy.AddMessage(e)
                        arcpy.AddError(e)
                else:
                    try:
                        row.setValue(fldn, table_attr.cell(k, j).value)
                    except Exception as e:
                        # print e
                        arcpy.AddError(e)

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
        arcpy.AddError(e)
        # arcpy.AddMessage(e)
    finally:
        if cur:
            del cur


if __name__ == "__main__":
    main(featuredata)
    #getRowIndex(table_attr,"BGJ1")