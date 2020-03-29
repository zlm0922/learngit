# !/usr/bin/env python
# -*- encoding:utf-8 -*-
import arcpy
import xlrd
import os
import sys

reload(sys)
sys.setdefaultencoding("utf-8")
arcpy.env.overwriteOutput = 1


def addfld():
    for j in range(nrows):
        if table.cell(j, 1).value == "":
            continue
        # for k in range(1,ncols):
        try:
            arcpy.AddField_management("Pipe_Environment/{}".format(table.cell(0,0).value),table.cell(j, 2).value,table.cell(j, 4).value.upper(),
                                  field_length=table.cell(j, 5).value,field_alias=table.cell(j, 3).value)

        except Exception as e:

            print e


template = r"D:\智能化管线项目\新气\新气数据处理\新气数据入库_0916\新库0910.gdb\Pipe_Environment"
arcpy.env.workspace = r"D:\智能化管线项目\新气\新气数据处理\新气数据入库_0916\新库0910.gdb"
xlspath = r"D:\智能化管线项目\新气\新气数据处理\新气数据入库_0916\creattable.xlsx".decode('utf-8')
myworkbook = xlrd.open_workbook(xlspath)
for i in range(myworkbook.nsheets):
    table = myworkbook.sheet_by_index(i) #定位到sheet页
    print myworkbook.sheet_names()[i]
    ncols = table.ncols
    nrows = table.nrows
    if myworkbook.sheet_names()[i] == 'Sheet2':
        print 1

        # arcpy.CreateFeatureclass_management("Pipe_Environment",table.cell(0,0).value,"POINT","",'ENABLED', 'ENABLED', template)
        # arcpy.AlterAliasName(table.cell(0,0).value,table.cell(0,1).value)
        # addfld()
    elif myworkbook.sheet_names()[i] == 'Sheet1':
        print 2
        # arcpy.CreateTable_management("Pipe_Risk", table.cell(2, 0).value, "POLYLINE", "", 'ENABLED', 'ENABLED',
        #                                     template)
        # arcpy.CreateTable_management(arcpy.env.workspace, table.cell(0, 0).value)
        # arcpy.AlterAliasName(table.cell(0, 0).value, table.cell(0, 1).value)
        # for j in range(nrows):
        #     if table.cell(j, 1).value == "":
        #         continue
        #     # for k in range(1,ncols):
        #     try:
        #         arcpy.AddField_management(table.cell(0, 0).value, table.cell(j, 2).value,
        #                                   table.cell(j, 4).value.upper(),
        #                                   field_length=table.cell(j, 5).value, field_alias=table.cell(j, 3).value)
        #
        #     except Exception as e:
        #         print e

    elif myworkbook.sheet_names()[i] == 'Sheet4':
        arcpy.CreateTable_management(arcpy.env.workspace,table.cell(2,0).value)
        arcpy.AlterAliasName(table.cell(0, 0).value, table.cell(0, 1).value)
        for j in range(nrows):
            if table.cell(j, 1).value == "":
                continue
            # for k in range(1,ncols):
            try:
                arcpy.AddField_management(table.cell(0, 0).value, table.cell(j, 2).value,
                                          table.cell(j, 4).value.upper(),
                                          field_length=table.cell(j, 5).value, field_alias=table.cell(j, 3).value)

            except Exception as e:
                print e



