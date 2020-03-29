# !/usr/bin/env python
# -*- encoding:utf-8 -*-
import arcpy
arcpy.env.overwriteOutput =1

line = r"D:\赋能\视频及教程_esri\2019年公开课地理处理及制图\demos\Demo3\line.shp"
distance = 200
beginstation = 0
# output = r"D:\赋能\视频及教程_esri\2019年公开课地理处理及制图\demos\Demo3\mygdb.gdb\mypoint_1"
output1 = r"D:\赋能\视频及教程_esri\2019年公开课地理处理及制图\demos\Demo3\mygdb.gdb\mypolyline"
mem_line = arcpy.CreateFeatureclass_management("in_memory", "mem_line", "Polyline", "", "DISABLED", "DISABLED", line)
#mem_line = arcpy.CreateFeatureclass_management('in_memory','mem_line','Polygon','','disabled','disabled',line)
print "Create mem_line successfully"
arcpy.AddField_management(mem_line, "LineOID", "TEXT")
arcpy.AddField_management(mem_line, "Value", "TEXT")

result = arcpy.GetCount_management(line)
features = int(result.getOutput(0))
fields = ["SHAPE@", "OID@"]

with arcpy.da.SearchCursor(line, (fields)) as search:
    with arcpy.da.InsertCursor(mem_line, ("SHAPE@", "LineOID", "Value")) as insert:
        for row in search:
            try:
                #beginstation = 0
                endstation = 200
                # print "endpoint_m:%f" % endstation
                line_geom = row[0]
                length = float(line_geom.length)
                #count = distance
                oid = str(row[1])
                start = arcpy.PointGeometry(line_geom.firstPoint)
                end = arcpy.PointGeometry(line_geom.lastPoint)
                while endstation <= length:

                    line1 = line_geom.segmentAlongLine(beginstation,endstation)

                    insert.insertRow((line1, oid, str(beginstation) + '-' + str(endstation)))
                    print "add line compelete"
                    beginstation = endstation
                    if beginstation < length:
                        
                        print "begin_m:%f" % beginstation
                    else:
                        break


                    endstation = beginstation + distance
                    if endstation > length:
                        endstation = length
                        print "endpoint_m:%f" % endstation
            except Exception as e:
                print e.message
arcpy.CopyFeatures_management(mem_line, output1)
arcpy.Delete_management(mem_line)
