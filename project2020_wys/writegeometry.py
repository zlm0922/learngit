# !/usr/bin/env python
# -*- coding:utf-8 -*-
import xlrd
import arcpy
import sys
import chardet

reload(sys)
sys.setdefaultencoding('utf-8')
arcpy.env.overwriteOutput = 1
arcpy.env.workspace = r'D:\赋能\视频及教程_esri\2019年公开课地理处理及制图\demos\Demo2\data'
xlsPath = r'D:\赋能\视频及教程_esri\2019年公开课地理处理及制图\demos\Demo2\Lot.xls'.decode('utf-8')
#print xlsPath
data = xlrd.open_workbook(xlsPath)
table = data.sheets()[0]  # 通过索引顺序获取
cols = table.col_values(3)
nrows = table.nrows
# point = arcpy.Point()
array = arcpy.Array()
polygonGeometryList = []
cur = arcpy.da.InsertCursor('polygon.shp',['Shape@','name','Id','type','type1'])

for i in range(1, nrows):
    global name,id
    str1 = table.cell(i, 3).value
    name = table.cell(i, 2).value
    type1 = table.cell(i, 1).value
    id = table.cell(i, 0).value
    points = str1.split(';')
    for j in points:
        xy = j.split(',')
        print xy[0]
        print xy[1]
        print '\n'
        point = arcpy.Point(float(xy[0]),float(xy[1]))
        # point.X = float(xy[0])
        # point.Y = float(xy[1])

        array.add(point)
    #print array
    #row = cur.newRow()
    row = arcpy.Polygon(array)

    #row.shape = array
    # row.name = name
    # row.Id = id


    array.removeAll()
    cur.insertRow([row,name,id,type1,type1])

#cur.insertRow(row)

del cur
# with arcpy.da.UpdateCursor('Polygon', ['name', 'Id']) as rows:
#     for rw in rows:
#         rw[0] = name
#         rw[1] = id
#         rw.updateRow()