# !/usr/bin/env python
# -*- encoding:utf-8 -*-
import arcpy
arcpy.env.overwriteOutput =1
line = r"D:\赋能\视频及教程_esri\2019年公开课地理处理及制图\demos\Demo3\line.shp"
distance = 200
output = r"D:\赋能\视频及教程_esri\2019年公开课地理处理及制图\demos\Demo3\mygdb.gdb\mypoint_71"
# output1 = r"D:\赋能\视频及教程_esri\2019年公开课地理处理及制图\demos\Demo3\mygdb.gdb\mypolyline"
mem_point = arcpy.CreateFeatureclass_management("in_memory", "mem_point", "POINT", "", "DISABLED", "DISABLED", line)
arcpy.AddField_management(mem_point, "LineOID", "TEXT")
arcpy.AddField_management(mem_point, "Value", "TEXT")

result = arcpy.GetCount_management(line)
features = int(result.getOutput(0))
fields = ["SHAPE@", "OID@"]

with arcpy.da.SearchCursor(line, (fields)) as search:
    with arcpy.da.InsertCursor(mem_point, ("SHAPE@", "LineOID", "Value")) as insert:
        for row in search:
            try:
                line_geom = row[0]
                length = float(line_geom.length)
                count = distance
                oid = str(row[1])
                start = arcpy.PointGeometry(line_geom.firstPoint)
                end = arcpy.PointGeometry(line_geom.lastPoint)
                while count <= length:
                    point = line_geom.positionAlongLine(count, False)

                    insert.insertRow((point, oid, count))
                    count += distance
            except Exception as e:
                print e.message
arcpy.CopyFeatures_management(mem_point, output)
arcpy.Delete_management(mem_point)


