# !/usr/bin/env python
# -*- coding:utf-8 -*-
import xlrd
import arcpy
import sys
import time
import os


reload(sys)
sys.setdefaultencoding('utf-8')
arcpy.env.overwriteOutput = 1

arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\部分雨水表格数据0620\JLSHYWS.gdb'
#first_p = arcpy.FeatureVerticesToPoints_management(pipeline,"first_p","Start")
iuputworkspace = r'D:\智能化管线项目\雨污水管线数据处理\JLSHYWS_0621ys\JLSHYWS.gdb\JLYWS'
#arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\JLSHYWS_0621ys\JLSHYWS.gdb'
xlsPath = r'D:\智能化管线项目\雨污水管线数据处理\部分雨水表格数据0620\8-三通探查表.xlsx'.decode('utf-8')

featuredata = os.path.join(iuputworkspace,'Tee')
#print xlsPath
data = xlrd.open_workbook(xlsPath)
#data = xlrd.open_workbook(xlspath)
table = data.sheets()[0]  # 通过索引顺序获取
#print table
#cols = table.col_values(3)
nrows = table.nrows
ncols = table.ncols
print ncols
print nrows
sheetname = table.name
exceltable = arcpy.ExcelToTable_conversion(xlsPath,sheetname)
#print sheetname
#point = arcpy.Point()


def colums_field_isnull():
    lst_colums_field = []
    for j in range(ncols):
        if len(set(table.col_values(j)))== 2 and '' in set(table.col_values(j)):
            #print table.cell(0,j).value
            lst_colums_field.append(table.cell(0,j).value.upper())
    return lst_colums_field


def attribute_join(sheetname,featuredata):
    try:
        #cur = arcpy.UpdateCursor(featuredata)
        tablename = [tblfield.name for tblfield in arcpy.ListFields(sheetname)]
        fldname = [field.name for field in arcpy.ListFields(featuredata)]
        layername = arcpy.MakeFeatureLayer_management(featuredata,featuredata.split("//")[-1]+"_lyr")
        arcpy.AddJoin_management(layername,"POINTNUMBER",exceltable,"POINTNUMBER")

        for fldn in fldname:

            for tblfldn in tablename:
                if tblfldn in ['OBJECTID','POINTNUMBER']:
                    continue
                if fldn == 'OPERATIONALS':
                    arcpy.CalculateField_management(layername, fldn, "!%s.OPERATIONALSTATUS!" % sheetname, "PYTHON_9.3")
                if fldn.upper() == tblfldn.upper() and fldn.upper() not in colums_field_isnull():
                    print fldn,   "!%s.%s!" %(sheetname,tblfldn)
                    try:
                        arcpy.CalculateField_management(layername,fldn,"!%s.%s!" %(sheetname,tblfldn),"PYTHON_9.3")
                    except arcpy.ExecuteError as e:
                        print e
        arcpy.RemoveJoin_management(layername,sheetname)

        savelyr = arcpy.CopyFeatures_management(layername,sheetname+"_frlyr")
        #print savelyr
        arcpy.DeleteFeatures_management(featuredata)
        arcpy.Append_management(savelyr,featuredata)
        arcpy.Delete_management(savelyr)


    except Exception as e:
        print e


if __name__ == "__main__":
    start_time = time.clock()
    attribute_join(sheetname,featuredata)
    # colums_field_isnull()
    finished_time = time.clock()
    elapsetime = finished_time - start_time
    print"elapsetime %s" % elapsetime