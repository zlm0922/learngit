# !/usr/bin/env python
# -*- encoding:utf-8 -*-
import arcpy
import xlrd
import os
import sys

reload(sys)
sys.setdefaultencoding("utf-8")

arcpy.env.overwriteOutput = 1
xlspath = r"D:\智能化管线项目\新气\新气数据处理\site.xlsx".decode("utf-8")

line = r"D:\智能化管线项目\新气\新气数据处理\新气建设期转到新气数据库\new20190709gdb_hww\new20190709.gdb\Pipe_Core\Cor_StationSeries"
#distance = 200  #等距离分割时使用
output = r"D:\赋能\视频及教程_esri\2019年公开课地理处理及制图\demos\Demo3\mygdb.gdb\mypoint_site"
# output1 = r"D:\赋能\视频及教程_esri\2019年公开课地理处理及制图\demos\Demo3\mygdb.gdb\mypolyline"
mem_point = arcpy.CreateFeatureclass_management("in_memory", "mem_point", "POINT", "", "ENABLED", "ENABLED", line)
arcpy.AddField_management(mem_point, "SITENAME", "TEXT")
arcpy.AddField_management(mem_point, "STATION", "DOUBLE")

result = arcpy.GetCount_management(line)
features = int(result.getOutput(0))
fields = ["SHAPE@"]

myworkbook = xlrd.open_workbook(xlspath)
table = myworkbook.sheet_by_index(0)
nrows = table.nrows
ncols = table.ncols

with arcpy.da.SearchCursor(line, fields) as search:
    with arcpy.da.InsertCursor(mem_point, ("SHAPE@", "SITENAME", "STATION")) as insert:
        for row in search:
            try:
                line_geom = row[0]
                length = float(line_geom.length)
                # count = distance
                for i in range(1,nrows):
                    station = table.cell(i,2).value
                    sitename = table.cell(i,1).value
                    ratio = table.cell(i,4).value
                    if station != "":
                        print station
                    # oid = str(row[1])
                    # start = arcpy.PointGeometry(line_geom.firstPoint)
                    # end = arcpy.PointGeometry(line_geom.lastPoint)
                    # while count <= length:# 等分时需要循环
                        point = line_geom.positionAlongLine(ratio, True)

                        insert.insertRow((point, sitename, station))
                        # count += distance
            except Exception as e:
                print e.message
arcpy.CopyFeatures_management(mem_point, output)
arcpy.Delete_management(mem_point)


