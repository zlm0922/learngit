#!/usr/bin/env python
# -*- coding:utf-8 -*-

# 根据点坐标生成 point feature
# 坐标保留8位小数
import arcpy
import sys
import xlrd
import os

reload(sys)
sys.setdefaultencoding('utf-8')

xlspath = r"D:\智能化管线项目\雨污水管线数据处理\原始文件\部分雨水表格数据0620\15-雨篦子探查表.xlsx".decode('utf-8')
arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\提交成果\JLSHYWS0620_zlm.gdb'
template = r'D:\智能化管线项目\雨污水管线数据处理\JLSHYWS_0621ys\JLSHYWS.gdb\JLYWS\ControlPoint'
featuredata = r'JLYWS\RainGate'
data = xlrd.open_workbook(xlspath)
table = data.sheet_by_index(0)
sheetname = data.sheet_names()[0]
# print sheetname
# print table.name
nrow = table.nrows
ncols = table.ncols

des = arcpy.Describe(template)
sparef = des.spatialReference
fm = arcpy.FieldMap()
fms = arcpy.FieldMappings()
# 没有使用此函数
def getColumnIndex(table, columnName):
    columnIndex = None
    for i in range(table.ncols):
        if(table.cell_value(0, i) == columnName):
            columnIndex = i
            break
    return columnIndex
arcpy.env.overwriteOutput = 1
tablelyr = os.path.join(xlspath,sheetname)
# x = table.cell(0,getColumnIndex(table,'X'))
# y = table.cell(0,getColumnIndex(table,'Y'))
# z = table.cell(0,getColumnIndex(table,'Z'))

xytable = arcpy.ExcelToTable_conversion(xlspath,sheetname)
controlp_lyr = arcpy.MakeXYEventLayer_management(xytable,"X","Y","controlP_lyr",sparef,"Z")
fieldmappings = arcpy.FieldMappings()

controlp = arcpy.CopyFeatures_management(controlp_lyr,"controlp")
# tblfldn = [tblfld.name for tblfld in arcpy.ListFields(controlp)]
# fldn = [fldn.name for fldn in arcpy.ListFields('JLYWS\ControlPoint')]
fieldmappings.addTable(featuredata)
fieldmappings.addTable(controlp)
for tblfld in arcpy.ListFields(controlp):

    for fld in arcpy.ListFields(featuredata):
        if fld.name.upper() == tblfld.name.upper():
            fieldmap = fieldmappings.getFieldMap(fieldmappings.findFieldMapIndex(fld.name))
            fieldmap.addInputField(controlp, tblfld.name)
            fieldmappings.replaceFieldMap(fieldmappings.findFieldMapIndex("TRACT2000"), fieldmap)

            # Remove the TRACTCODE fieldmap.
            #
            fieldmappings.removeFieldMap(fieldmappings.findFieldMapIndex("TRACTCODE"))

            # print fld.name
            try:
                fm.addInputField(controlp, tblfld.name)
                print tblfld.name
            except Exception as e:
                print e

fms.addFieldMap(fm)
try:
    arcpy.Append_management(controlp,featuredata,"NO_TEST",fms)
except Exception as e:
    print e
arcpy.Delete_management(controlp)


