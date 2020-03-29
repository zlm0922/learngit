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


def modify_table_field(tablename):
    # fldname = [fldn.name for fldn in arcpy.ListFields(tablename)]
    # flag = ''
    # flag_j = -1
    for j in range(ncols):
        # flag = table.cell(4, j).value
        # flag_j = j
        if ' ' in table.cell(4, j).value:

            print 'excel存在空格的字段名:%s'% table.cell(4, j).value

        for fld in arcpy.ListFields(tablename):
            if fld.aliasName == table.cell(
                3,
                j).value.replace(
                ' ',
                '') and fld.name.upper() != table.cell(
                4,
                j).value.upper().replace(
                ' ',
                    ''):

                print "修改字段名%s为%s" % (fld.name,
                                      table.cell(
                                          4,
                                          j).value.replace(
                                          ' ',
                                          ''))
                try:
                    arcpy.AlterField_management(
                        tablename, fld.name, table.cell(
                            4, j).value.replace(
                            ' ', ''))
                except arcpy.ExecuteError as e1:
                    print e1

                break
            elif fld.name.upper() == table.cell(4, j).value.upper().replace(' ', '') and \
                    fld.aliasName != table.cell(3, j).value.replace(' ', ''):
                print "修改字段别名%s为%s" % (fld.aliasName,
                                       table.cell(
                                           3,
                                           j).value.replace(
                                           ' ',
                                           ''))
                try:
                    fld.aliasName = table.cell(3, j).value.replace(' ', '')
                except Exception as e2:
                    print e2

                break
            elif fld.aliasName == table.cell(3, j).value.replace(' ', '') and \
                    fld.name.upper() == table.cell(4, j).value.upper().replace(' ', ''):

                break
        else:
            print '添加的字段名:',table.cell(4, j).value
            # print 'flag_j:',flag_j
            try:
                if table.cell(4, j).value.endswith('date') or table.cell(4, j).value.endswith('time'):
                    arcpy.AddField_management(
                        tablename,
                        table.cell(4, j).value.replace(
                            ' ',
                            ''),
                        "DATE",
                        field_alias=table.cell(
                            3,
                            j).value.replace(
                            ' ',
                            ''))
                    print "执行添加字段%s" % table.cell(4, j).value
                else:
                    arcpy.AddField_management(
                        tablename,
                        table.cell(
                            4,
                            j).value.replace(
                            ' ',
                            ''),
                        "TEXT",
                        field_length=50,
                        field_alias=table.cell(
                            3,
                            j).value.replace(
                            ' ',
                            ''))
                    print "执行添加字段%s" % table.cell(4, j).value
            except arcpy.ExecuteError as e:
                print e


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
        print e
        print "\t 请检查表:%s 是否表头没有删除中文字段及注释部分" % filename
        arcpy.AddMessage(e)
    else:
        print "导数数据表{0}  {1}条数据".format(featuredata, nrows - 1)

    finally:
        if cur:

            del cur


if __name__ == "__main__":

    arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\JLSHYWS_kkgdb\JLSHYWS_1.gdb'
    # 数据表excel所在文件path
    wks = r'D:\智能化管线项目\雨污水管线数据处理\雨污水采集模板\污雨水模板整理0318'.decode('utf-8')
    # def iteration_excel(wks):
    feature_classes = arcpy.ListFeatureClasses(feature_dataset='JLYWS')
    table_classes = arcpy.ListTables()
    for ft_cls in (feature_classes + table_classes):
        # print ft_cls
        for root, files, filenames in os.walk(wks):
            for filename in filenames:
                pre_filename = os.path.splitext(filename)[0]
                # print os.path.join(root, filename)
                try:
                    data = xlrd.open_workbook(os.path.join(root, filename))
                    table = data.sheet_by_index(0)
                    sheetname = data.sheet_names()[0]
                    nrows = table.nrows
                    ncols = table.ncols
                    # if table_1 in data.sheet_by_name()
                    # for i in range(data.nsheets):
                    #     if data.sheet_names()[i] in ['domain','Domain','阈值']:

                            # print data.sheet_names()[i]
                    table_1 = data.sheet_by_name('Domain')
                    # print table_1.cell(0, 0).value.strip()
                    if ft_cls.upper() == table_1.cell(0,0).value.strip().upper():
                        print filename
                        print table_1.cell(0,0).value.strip()
                        modify_table_field(ft_cls)
                        break

                except IOError as e:
                    print e
                    print "输入文件为非excel文件"
                except Exception as ee:
                    pass
                # finally:
                #     if data:
                #         data.close()