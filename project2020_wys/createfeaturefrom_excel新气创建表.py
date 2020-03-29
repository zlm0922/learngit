# !/usr/bin/env python
# -*- encoding:utf-8 -*-
import arcpy
import xlrd
import os
import sys

reload(sys)
sys.setdefaultencoding("utf-8")
arcpy.env.overwriteOutput = 1

#  适用于表名和字段同样行数，需要修改表名所在单元格index
def addfld():
    for j in range(1,nrows):
        if table.cell(j, 2).value == "":
            continue
        # for k in range(1,ncols):
        try:
            # arcpy.AddField_management()
            arcpy.AddField_management(table.cell(2, 0).value,table.cell(j, 1).value,table.cell(j, 3).value.upper(),
                                  field_length = table.cell(j, 4).value,field_alias=table.cell(j, 2).value)
            print "为{}添加字段{}完成!".format(table.cell(2, 0).value,table.cell(j, 1).value)
        except arcpy.ExecuteError as e:
            print e


template = r"D:\智能化管线项目\新气\新气数据处理\新气建设期转到新气数据库\new20190709gdb_hww\new20190709.gdb\Pipe_Risk"
arcpy.env.workspace = r'Database Connections\Connection to 127.0.0.1.sde'
# arcpy.env.workspace = r'D:\智能化管线项目\新气\新气数据处理\新气数据入库_0916\新库0910zlm.gdb'#r'Database Connections\Connection to 10.246.146.120.sde'#r"D:\智能化管线项目\销售华北1023数据表字段add\总部空库20180105.gdb"
arcpy.env.workspace = r'Database Connections\Connection to 10.246.146.120.sde'#r"D:\智能化管线项目\销售华北1023数据表字段add\总部空库20180105.gdb"
xlspath = r"D:\智能化管线项目\新气\新气数据处理\标准数据库-20200115.xlsx".decode('utf-8')
myworkbook = xlrd.open_workbook(xlspath)
for i in range(myworkbook.nsheets):

    table = myworkbook.sheet_by_index(i)  # 定位到sheet页
    ncols = table.ncols
    nrows = table.nrows
    #全部为属性表
    print myworkbook.sheet_names()[i]
    arcpy.CreateTable_management(arcpy.env.workspace, table.cell(2, 0).value)
    arcpy.AlterAliasName(table.cell(2, 0).value, table.cell(1, 0).value)
    addfld()
    #有空间表的情况
    # if i == 0 or i == myworkbook.nsheets-1:
    #     print myworkbook.sheet_names()[i]
    #     # arcpy.CreateFeatureclass_management()
    #     arcpy.CreateFeatureclass_management("SDE.Pipe_Integrity",table.cell(2,0).value,"POLYLINE","",'ENABLED', 'DISABLED', template)
    #     arcpy.AlterAliasName(table.cell(2,0).value,table.cell(1,0).value)
    #     addfld()
    # # elif i == 1:
    # #     arcpy.CreateFeatureclass_management("Pipe_Risk", table.cell(2, 0).value, "POLYLINE", "", 'ENABLED', 'ENABLED',
    # #                                         template)
    # #     arcpy.AlterAliasName(table.cell(2, 0).value, table.cell(1, 0).value)
    # #     addfld()
    # else:
    #     print myworkbook.sheet_names()[i]
    #     arcpy.CreateTable_management(arcpy.env.workspace,table.cell(2,0).value)
    #     arcpy.AlterAliasName(table.cell(2, 0).value, table.cell(1, 0).value)
    #     addfld()
    #     # for j in range(1, nrows):
    #     #     if table.cell(j, 1).value == "":
    #     #         continue
    #     #     # for k in range(1,ncols):
    #     #     try:
    #     #         arcpy.AddField_management(table.cell(2, 0).value, table.cell(j, 1).value,
    #     #                                   table.cell(j, 3).value.upper(),
    #     #                                   field_length=table.cell(j, 4).value, field_alias=table.cell(j, 2).value)
    #     #
    #     #     except Exception as e:
    #     #         print e



