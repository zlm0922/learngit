#!/usr/bin/env python
# -*- coding:utf-8 -*-

# 根据点坐标生成 point feature
# 坐标保留8位小数
import arcpy
import sys
import xlrd
import os
import datetime
from xlutils.copy import copy
from openpyxl import load_workbook
reload(sys)
sys.setdefaultencoding('utf-8')


def getColumnIndex(table, columnName):
    columnIndex = None
    for i in range(table.ncols):
        if table.cell_value(0, i) == columnName:
            columnIndex = i
            break
    return columnIndex

if __name__ == "__main__":
    # index_dic = {
    #     'RainGate': '15-雨篦子探查表'
    #     }

    # 数据库文件
    arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\JLSHYWS_kkgdb\JLSHYWS_1.gdb'
    # 数据表excel所在文件path
    wks = r'D:\智能化管线项目\雨污水管线数据处理\雨污水采集模板\污雨水模板整理0318'.decode('utf-8')
    # def iteration_excel(wks):

    # for ft_cls in (feature_classes + table_classes):
        # print ft_cls
    for root, files, filenames in os.walk(wks):
        for filename in filenames:

            pre_filename = os.path.splitext(filename)[0]
            # print os.path.join(root, filename)
            try:
                # data = xlrd.open_workbook(os.path.join(root, filename))
                wb = load_workbook(os.path.join(root, filename))
                # table = data.sheet_by_index(0)
                #                 # sheetname = data.sheet_names()[0]
                #                 # nrows = table.nrows
                #                 # ncols = table.ncols
                # if table_1 in data.sheet_by_name()
                # for i in range(data.nsheets):
                #     if data.sheet_names()[i] in ['domain','阈值']:
                ws = wb['阈值']
                ws.title = 'Domain'



                wb.save(os.path.join(root, filename))
                wb.close()
            except IOError as e:
                print e
                print "输入文件为非excel文件"
            except Exception as ee:
                pass

