#!/usr/bin/env python
# -*-coding:utf-8 -*-

import arcpy
import os
import sys

reload(sys)
sys.setdefaultencoding('utf-8')


def select_identify_point(pipeline, controlP, pointlocation):
    try:
        selectP = arcpy.FeatureVerticesToPoints_management(
            pipeline, "selectP" + pointlocation, pointlocation)
        #print selectP
        resultP = arcpy.SpatialJoin_analysis(
            selectP,
            controlP,
            "resultP" + pointlocation,
            search_radius=0.0000001195)  # ,search_radius=0.0000001195
        # print os.path.abspath( "resultP" + pointlocation)
        # print os.path.(resultP)
        # print type(resultP)
        # print resultP[0].split("\\")[-1]

        return resultP
    except arcpy.ExecuteError as e:
        print e


def update_fielditem(featuredata, startcode, endcode):
    with arcpy.da.UpdateCursor(featuredata, ["SHAPE@", startcode, endcode]) as cur:
        for row in cur:
            Pstart = select_identify_point(row[0], controlP, "START")
            Pend = select_identify_point(row[0], controlP, "END")
            for ssrow in arcpy.da.SearchCursor(Pstart, ["POINTNUMBER"]):
                row[1] = ssrow[0]
                cur.updateRow(row)
            for serow in arcpy.da.SearchCursor(Pend, ["POINTNUMBER"]):
                row[2] = serow[0]
                cur.updateRow(row)


if __name__ == "__main__":
    arcpy.env.overwriteOutput = 1
    arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\过程数据库文件\JLSHYWS0714temp.gdb'
    # first_p = arcpy.FeatureVerticesToPoints_management(pipeline,"first_p","Start")
    iuputworkspace = r'D:\智能化管线项目\雨污水管线数据处理\multipart_pipeline.gdb'  # 设置要添加属性的数据库
    pipeline = os.path.join(iuputworkspace, r"pipeline\multipart_pipeline")
    controlP = os.path.join(
        iuputworkspace,
        r"pipeline\ControlPoint")  # 注意修改数据集
    update_fielditem(pipeline, "STARTPOINTNUM", "ENDPOINTNUM")
