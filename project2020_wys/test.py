#!/usr/bin/env python
# -*- coding:utf-8 -*-
import sys
import xlrd
import os
import datetime
import arcpy

reload(sys)
sys.setdefaultencoding('utf-8')

arcpy.env.overwriteOutput = 1

xlspath_coordinate = r"D:\智能化管线项目\雨污水管线数据处理\部分高架线路0624\北区含油污B线水高架\ttt.xlsx".decode('utf-8')
xlspath_attribute = r"D:\智能化管线项目\雨污水管线数据处理\部分高架线路0624\北区含油污B线水高架\6-管廊（带）探查表.xls".decode('utf-8')
arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\JLSHYWS_0624\JLSHYWSb线.gdb'
featuredata = r'JLYWS\PipeGallery'




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
fc = "D:/data/gdb.gdb/roads"
outFC = "D:/data/gdb.gdb/roads9"
# spatial_ref = arcpy.Describe(fc).spatialReference
spatial_ref = arcpy.Describe(featuredata).spatialReference
arcpy.CreateFeatureclass_management(os.path.dirname(outFC),
                                        os.path.basename(outFC),
                                        "POLYLINE",spatial_reference= spatial_ref)


des = arcpy.Describe(outFC)
cursor = arcpy.da.InsertCursor(outFC,["Shape@"])
shapename = des.shapeFieldName
print shapename
array = arcpy.Array()
point = arcpy.Point()
for i in range(1,114):
    point.X = round(float(table_coor.cell(i, 7).value), 8)
    point.Y = round(float(table_coor.cell(i, 8).value), 8)
    point.Z = round(float(table_coor.cell(i, 9).value), 8)
    # array.add(point)
    array.add(point)
for feat in array:
    print feat
polyline = arcpy.Polyline(array,spatial_ref)
# row = cursor.newRow()
# row.setValue(shapename,array)
#polyline = arcpy.Polyline(array)

cursor.insertRow([polyline])
array.removeAll()
