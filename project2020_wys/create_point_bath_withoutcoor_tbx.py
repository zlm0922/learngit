#!/usr/bin/env python
# -*- coding:cp936 -*-

# ���ݵ��������� point feature
# ���걣��8λС��
import arcpy
import sys
import xlrd
import os
import datetime

reload(sys)
sys.setdefaultencoding('utf-8')


def getRowIndex(table, RowName, colfld):
    RowIndex = None
    for i in range(table.nrows):
        # for j in range(table.ncols):
        if (table.cell_value(i, getColumnIndex(table, colfld))
            ).upper() == RowName.upper():
            RowIndex = i
            break
    else:
        # print "��ȷ��%s�������Ƿ���ȷ" % RowName
        arcpy.AddMessage("check input %s is correct or not.".format(RowName))
    return RowIndex


def getColumnIndex(table, columnName):
    columnIndex = None
    for i in range(table.ncols):
        if table.cell_value(0, i) == columnName:
            columnIndex = i
            break
    return columnIndex


def create_control_point(featuredata):
    cur = None
    try:
        cur = arcpy.InsertCursor(featuredata)
        fldname = [fldn.name for fldn in arcpy.ListFields(featuredata)]
        for i in range(1, nrows):
            L1 = table.cell(i, getColumnIndex(table, "X")).value
            B1 = table.cell(i, getColumnIndex(table, "Y")).value
            H1 = table.cell(i, getColumnIndex(table, "Z")).value
            row = cur.newRow()
            if L1:
                point = arcpy.Point(
                    round(
                        float(L1), 8), round(
                        float(B1), 8), float(H1))
                row.shape = point
            for fldn in fldname:
                for j in range(0, ncols):
                    #print 112
                    if fldn == (str(table.cell(0, j).value).strip()).upper():
                        # print table.cell(i, j).ctype
                        if table.cell(i, j).ctype == 3:
                            # print table.cell(i, j).ctype
                            date = xlrd.xldate_as_tuple(
                                table.cell(i, j).value, 0)
                            # print(date)
                            tt = datetime.datetime.strftime(
                                datetime.datetime(*date), "%Y-%m-%d")
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
        # print e
        arcpy.AddMessage(e)
        # print "\t �����:%s �Ƿ��ͷû��ɾ�������ֶμ�ע�Ͳ���" % filename
        arcpy.AddMessage("\t check table:%s chinese fieldname deleted" % filename)

    else:
        # print "\t�������ݱ�{0}  {1}������".format(featuredata, nrows)
        arcpy.AddMessage("\t table{0} finished dataloading {1}items".format(featuredata, nrows))

    finally:
        if cur:
            del cur


def create_other_point(featuredata):
    cur = None
    try:
        cur = arcpy.InsertCursor(featuredata)
        fldname = [fldn.name for fldn in arcpy.ListFields(featuredata)]
        for i in range(1, nrows):

            L1 = table_coor.cell(
                getRowIndex(
                    table_coor, table.cell(
                        i, getColumnIndex(
                            table, "POINTNUMBER")).value, "POINTNUMBER"), getColumnIndex(
                    table_coor, "X")).value
            B1 = table_coor.cell(
                getRowIndex(
                    table_coor, table.cell(
                        i, getColumnIndex(
                            table, "POINTNUMBER")).value, "POINTNUMBER"), getColumnIndex(
                    table_coor, "Y")).value
            H1 = table_coor.cell(
                getRowIndex(
                    table_coor, table.cell(
                        i, getColumnIndex(
                            table, "POINTNUMBER")).value, "POINTNUMBER"), getColumnIndex(
                    table_coor, "Z")).value
            row = cur.newRow()
            if L1:
                point = arcpy.Point(
                    round(
                        float(L1), 8), round(
                        float(B1), 8), float(H1))
                row.shape = point
            for fldn in fldname:
                for j in range(0, ncols):
                    #print 112
                    if fldn == (str(table.cell(0, j).value).strip()).upper():
                        # print table.cell(i, j).ctype
                        if table.cell(i, j).ctype == 3:
                            # print table.cell(i, j).ctype
                            date = xlrd.xldate_as_tuple(
                                table.cell(i, j).value, 0)
                            # print(date)
                            tt = datetime.datetime.strftime(
                                datetime.datetime(*date), "%Y-%m-%d")
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
        # print e
        # print "\t �����:%s �Ƿ��ͷû��ɾ�������ֶμ�ע�Ͳ���" % filename
        arcpy.AddMessage(e)
        arcpy.AddMessage("\t check table:%s chinese fieldname deleted" % filename)
        arcpy.AddMessage(e)
    else:
        # print "�������ݱ�{0}  {1}������".format(featuredata, nrows-1)
        arcpy.AddMessage("\t table{0} finished dataloading {1}items".format(featuredata, nrows))

    finally:
        if cur:
            del cur


if __name__ == "__main__":
    index_dic = {'ControlPoint':u"3-���߿��Ƶ�̽���",'Tee': u"8-��ͨ̽���",'valve': u"5-����̽���",'Elbow': u"10-��ͷ̽���",'Manhole': u"4-񿾮̽���",'RainGate': u"15-������̽���",'FourLink': u"9-��ͨ̽���"}
    # arcpy.AddMessage(index_dic['ControlPoint'])
    # index_dic = {
    #     'RainGate': '15-������̽���'
    #     }

    # ���߳ɹ����ļ�·��
    # global xlspath_coordinate
    xlspath_coordinate = arcpy.GetParameterAsText(1)
    data_coordinate = xlrd.open_workbook(xlspath_coordinate)
    table_coor = data_coordinate.sheet_by_index(0)
    sheetname_coor = data_coordinate.sheet_names()[0]

    # ���ݿ��ļ�
    arcpy.env.workspace = arcpy.GetParameterAsText(0)#r'D:\���ܻ�������Ŀ\����ˮ�������ݴ���\�ύ�ɹ�\JLSHYWS0730�겹.gdb'
    # ���ݱ�excel�����ļ�path
    wks = arcpy.GetParameterAsText(2)#r'D:\���ܻ�������Ŀ\����ˮ�������ݴ���\ԭʼ�ļ�\�Ͻ�����0730\�ܳ���ˮ���²�������'.decode('utf-8')
    # def iteration_excel(wks):
    for root, files, filenames in os.walk(wks):
        for filename in filenames:
            # print filename
            arcpy.AddMessage(filename)
            pre_filename = os.path.splitext(filename)[0]
            # print os.path.join(root, filename)
            try:
                # if filename != os.path.split(xlspath_coordinate)[-1]:
                data = xlrd.open_workbook(os.path.join(root, filename))
                table = data.sheet_by_index(0)
                sheetname = data.sheet_names()[0]

                nrows = table.nrows
                ncols = table.ncols
                for key in index_dic.keys():
                    if index_dic[key] == pre_filename:
                        # print key,pre_filename
                        arcpy.AddMessage("key:{},value:{}".format(key, pre_filename))
                        # print fcls
                        featuredata = r'JLYWS\{}'.format(key)
                        # print featuredata
                        if key == "ControlPoint":
                            create_control_point(featuredata)
                        else:
                            create_other_point(featuredata)
            except IOError as e:
                # print e
                # print "�����ļ�Ϊ��excel�ļ�"
                arcpy.AddMessage(e)
                arcpy.AddMessage("Inut file is not excel file")
            except Exception as ee:
                arcpy.AddMessage(ee)
                # print ee


