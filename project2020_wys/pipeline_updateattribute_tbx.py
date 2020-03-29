#!/usr/bin/env python
# -*-coding:cp936 -*-

import arcpy
import os
import sys

# reload(sys)
# sys.setdefaultencoding('utf-8')

# ���ù����ļ��Ĵ洢λ��
arcpy.env.workspace = arcpy.GetParameterAsText(0)#r'D:\���ܻ�������Ŀ\����ˮ�������ݴ���\�������ݿ��ļ�\JLSHYWS0714temp.gdb'
#first_p = arcpy.FeatureVerticesToPoints_management(pipeline,"first_p","Start")
iuputworkspace =arcpy.GetParameterAsText(1)# r'D:\���ܻ�������Ŀ\����ˮ�������ݴ���\�ύ�ɹ�\JLSHYWS0730_4��ˮ��.gdb' #����Ҫ������Ե����ݿ�
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
        print "��ʼִ����ʼ��"
        arcpy.AddMessage(u"��ʼִ����ʼ��")
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
        print "��ʼ���ֶμ������"
        arcpy.AddMessage(u"��ʼ���ֶμ������")
        print "\n"
        print "��ʼִ���յ�"
        arcpy.AddMessage(u"��ʼִ���յ�")
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
        print "�����ȡ���"
        arcpy.AddMessage(u"�����ȡ���")
    except arcpy.ExecuteError as e:
        print e

    else:
        arcpy.Delete_management(pipewithattr)


def join_attr2_field(featuredata,com_fld_name,jointable,join_field):
    ad_fld_dic = {"PIPECODE":"�ܵ�����","PIPENAME":"�ܵ�����"}
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
