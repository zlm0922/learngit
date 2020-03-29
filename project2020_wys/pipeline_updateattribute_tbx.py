#!/usr/bin/env python
# -*-coding:cp936 -*-

import arcpy
import os
import sys

# reload(sys)
# sys.setdefaultencoding('utf-8')

# 设置过程文件的存储位置
arcpy.env.workspace = arcpy.GetParameterAsText(0)#r'D:\智能化管线项目\雨污水管线数据处理\过程数据库文件\JLSHYWS0714temp.gdb'
#first_p = arcpy.FeatureVerticesToPoints_management(pipeline,"first_p","Start")
iuputworkspace =arcpy.GetParameterAsText(1)# r'D:\智能化管线项目\雨污水管线数据处理\提交成果\JLSHYWS0730_4污水汽.gdb' #设置要添加属性的数据库
pipeline = os.path.join(iuputworkspace, r"JLYWS\PipeLine")
controlP = os.path.join(iuputworkspace, r"JLYWS\ControlPoint")


def select_identify_point(pointlocation):
    try:
        selectP = arcpy.FeatureVerticesToPoints_management(
            pipeline, "selectP" + pointlocation, pointlocation)
        #print selectP
        resultP = arcpy.SpatialJoin_analysis(
            selectP, controlP, "resultP" + pointlocation)
        # print os.path.abspath( "resultP" + pointlocation)
        #print os.path.(resultP)
        # print type(resultP)
        # print resultP[0].split("\\")[-1]

        return resultP
    except arcpy.ExecuteError as e:
        print e


def pipe_attribute_add(pipeline):
    try:
        pipelyr = arcpy.MakeFeatureLayer_management(pipeline, "pipelyr")
        print "开始执行起始点"
        arcpy.AddMessage(u"开始执行起始点")
        P_start = select_identify_point("START")
        print P_start[0]
        base_p_start = os.path.basename(P_start[0])
        P_end = select_identify_point('END')
        base_p_end = os.path.basename(P_end[0])
        arcpy.AddJoin_management(
            pipelyr,
            "PipeLineNumber",
            P_start,
            "PipeLineNumber")
        arcpy.CalculateField_management(
            pipelyr, 'StartPointNum', "!%s.POINTNUMBER!"% base_p_start , "PYTHON_9.3")
        arcpy.CalculateField_management(
            pipelyr, 'OrgName', "!%s.OrgName_1!" % base_p_start, "PYTHON_9.3")
        arcpy.CalculateField_management(
            pipelyr, 'OrgCode', "!%s.OrgCode_1!" % base_p_start, "PYTHON_9.3")
        arcpy.RemoveJoin_management(
            pipelyr,base_p_start)
        print "起始点字段计算结束"
        arcpy.AddMessage(u"起始点字段计算结束")
        print "\n"
        print "开始执行终点"
        arcpy.AddMessage(u"开始执行终点")
        arcpy.AddJoin_management(
            pipelyr,
            "PipeLineNumber",
            P_end,
            "PipeLineNumber")
        arcpy.CalculateField_management(
            pipelyr,
            'EndPointNum',
            "!POINTNUMBER!",
            "PYTHON_9.3")
        arcpy.RemoveJoin_management(
            pipelyr, base_p_end)
        pipewithattr = arcpy.CopyFeatures_management(pipelyr, "pipewithattr")
        arcpy.DeleteFeatures_management(pipeline)
        arcpy.Append_management(pipewithattr,pipeline)
        print "点号提取完成"
        arcpy.AddMessage(u"点号提取完成")
    except arcpy.ExecuteError as e:
        print e

    else:
        arcpy.Delete_management(pipewithattr)


def join_attr2_field(featuredata,com_fld_name,jointable,join_field):
    ad_fld_dic = {"PIPECODE":"管道编码","PIPENAME":"管道名称"}
    for key,val in ad_fld_dic.items():
        arcpy.AddField_management(featuredata,key,"String",field_alias=val,field_length=255)
    tem_lyr = arcpy.MakeFeatureLayer_management(featuredata)
    arcpy.AddJoin_management(tem_lyr,com_fld_name,jointable,join_field)
    for key in ad_fld_dic.keys():
        arcpy.CalculateField_management(tem_lyr,key,"!{0}.{1}!".format(jointable,key))
    arcpy.RemoveJoin_management(tem_lyr,jointable)


if __name__ == "__main__":
    arcpy.env.overwriteOutput = 1
    # select_identify_point("START")
    pipe_attribute_add(pipeline)
