# !/usr/bin/ env python
# -*-coding:utf-8 -*-

import arcpy
import os
# 获得可用的ArcSDE连接文件
# def GetArcSDEconnectionfile():
# def GetArcSDEConnectionFile():
#     sdefilename = 'Connection to 10.246.146.120.sde'
#     # arcpy.CreateArcSDEConnectionFile_management(sdefilepath,sdefilename,"")
#     arcpy.CreateDatabaseConnection_management("Database Connections", sdefilename, "ORACLE", 'ipgis', 'DATABASE_AUTH',
#                                               'JASFRAMEWORK',
#                                               '123', 'SAVE_USERNAME')
#     return os.path.join("Database Connections", sdefilename)


folderPath = r"D:\智能化管线项目\新气\MXDchange".decode('utf-8')
dessdeFile = r"Database Connections\Connection to 10.246.146.101.sde"
# print dessdeFile
for filename in os.listdir(folderPath):
    fullpath = os.path.join(folderPath, filename)
    if os.path.isfile(fullpath):
        basename, extension = os.path.splitext(fullpath)
        if extension.lower() == ".mxd":
            mxd = arcpy.mapping.MapDocument(fullpath)
            # 枚举出有问题数据源的图层列表
            brknList = arcpy.mapping.ListBrokenDataSources(mxd)
            if len(brknList) == 0:
                print filename + ':No problem!'
            else:

                for brknItem in brknList:

                    # 将有问题的地图文档的图层信息打印输出
                    print "\t" + brknItem.name + "-Original source:" + brknItem.dataSource
                else:
                    # 进行数据源的更换
                    print "************starting Replace Datasource******************"
                    mxd.findAndReplaceWorkspacePaths(brknItem.workspacePath, dessdeFile,False)
                    print brknItem.workspacePath,dessdeFile
                    # print brknItem.dataSource
                    newmxdpath = basename + "_copy.mxd"
                    # 生成一个新的MXD
                    mxd.saveACopy(newmxdpath)
                    print "***********successful Replace Datasource******************"
                    # print "***********starting export JPEG******************"
                    # # 获得新MXD的对象
                    # newmxd = arcpy.mapping.MapDocument(newmxdpath)
                    # # 导出MXD的JPEG文件
                    # arcpy.mapping.ExportToJPEG(newmxd, basename + "_copy.jpg")
                    # print "***********successful export JPEG******************"

                del mxd
                # del newmxd
