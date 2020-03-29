#!/usr/bin/python
# -*- coding: utf-8 -*-

import arcpy, time, sys,os
from datetime import datetime, date
from arcpy import env
from lstfc import list_feature_batch

env.overwriteOutput = 1
# if len(sys.argv) != 6:
#     print "Usage  : python export-sde-fgdb.py basepath-ending-with-slash sde-password sde-servername sde-service " \
#           "sde-database "
#     sys.exit()

try:
    start = time.clock()
    basedir = r'D:'
    env.workspace = r"Database Connections\Connection to 127.0.0.1.sde"
    env.workspace = r"Database Connections\Connection to 10.246.146.120.sde"
    # 获得指定ArcSDE数据源的版本列表
    versionList = arcpy.ListVersions(env.workspace)
    # 获得数据集、要素类、表的列表对象
    fcList = arcpy.ListFeatureClasses()
    dsList = arcpy.ListDatasets()
    tbList = arcpy.ListTables()
    # print tbList

    folderName = basedir+ os.sep + "ConnectionFiles"

    # serverName = sys.argv[3]
    # serviceName = sys.argv[4]
    # databaseName = sys.argv[5]
    # authType = "DATABASE_AUTH"
    # username = "sde"
    # password = sys.argv[2]
    # saveUserInfo = "SAVE_USERNAME"
    # saveVersionInfo = "SAVE_VERSION"

    connFiles = []

    baseFilename = datetime.now().strftime("%Y-%m-%d_%H-%M")
    print "BaseFileDatabaseName: " + baseFilename
    # 循环每一个版本信息，针对每一个版本进行数据备份
    for version in versionList:
        fileName = baseFilename + "-VER-" + str(version)
        fileName = fileName.replace('.', '_')
        fileName = fileName.replace(':', '_')
        fileName = fileName.replace(' ', '-')
        # 根据不同的子版本信息，创建相应的ArcSDE连接信息
        # arcpy.CreateDatabaseConnection_management("Database Connections", "Connection to 10.246.146.120.sde", "ORACLE", 'ipgis', 'DATABASE_AUTH',
        #                                           'JASFRAMEWORK',
        #                                           '123', 'SAVE_USERNAME')
        connFiles.append(fileName)

    for conn in connFiles:
        fileGeoDbLocation = basedir + os.sep +"databases"
        fileGeoDb = conn + ".gdb"
        print conn
        # 创建针对某个版本数据存储的文件地理数据库对象
        arcpy.CreateFolder_management(basedir,"databases")
        arcpy.CreateFileGDB_management(fileGeoDbLocation, fileGeoDb)

        # env.workspace = folderName + "\\" + conn + ".sde"
        # print env.workspace
        totaldest = fileGeoDbLocation + "\\" + fileGeoDb
        print totaldest
        if tbList:
            for lmitc in tbList:
                if lmitc[:lmitc.find('.')] == "SDE":
                    desttc = lmitc[lmitc.find('.')+1:]
                    totaltcdest = totaldest + os.sep + desttc
                    print lmitc + " ---->  " + totaltcdest
                    try:
                        arcpy.Copy_management(lmitc, totaltcdest)
                    except arcpy.ExecuteError as ee:
                        print ee
                        # arcpy.Copy_management(lmitc, totaltcdest+"_1")
        for lmids in dsList:
            srcds = lmids
            fcinds = arcpy.ListFeatureClasses("", "All", lmids)
            destds = lmids[lmids.find('.')+1:]
            totaldsdest = totaldest + os.sep + destds
            print lmids + " ---->  " + totaldsdest
            try:
                arcpy.Copy_management(lmids, totaldsdest,"Dataset")
            except arcpy.ExecuteError as ee:
                print ee
                # arcpy.Copy_management(lmids, totaldsdest+'_1', "Dataset")

            # if (destds != "RELATED_DS") and (destds.find("RELATED_FC") == -1):
        #     for fccopy in fcinds:
        #         srcfc = lmids + "\\" + fccopy
        #         destdsfc = fccopy[fccopy.find('.')+1:]
        #         totaldsfcdest = totaldsdest + os.sep + destdsfc
        #         print srcfc + " ---->  " + totaldsfcdest
        #         # 使用Copy函数将对象从ArcSDE数据源拷贝到文件地理数据库
        #         try:
        #             arcpy.Copy_management(srcfc, totaldsfcdest, "FeatureClass")
        #         except arcpy.ExecuteError as ee:
        #             print ee
        # #
        if fcList:
            for lmifc in fcList:
                destfc = lmifc[lmifc.find('.')+1:]
                totalfcdest = totaldest + os.sep +destfc
                print lmifc + " ---->  " + totalfcdest
                try:
                    arcpy.Copy_management(lmifc, totalfcdest)
                except arcpy.ExecuteError as ee:
                    print ee
                    # arcpy.Copy_management(lmifc, totalfcdest+"_1")

        print "Compressing : " + fileGeoDbLocation + "\\" + fileGeoDb
        # 压缩一下文件地理数据库
        arcpy.CompressFileGeodatabaseData_management(fileGeoDbLocation + "\\" + fileGeoDb)

    elapsed = (time.clock() - start)
    print "Total time : " + str(elapsed) + " seconds"
    print "or : " + str(elapsed / 60) + " minutes"
    print "or : " + str(elapsed / 3600) + " hours"
    resultFile = open(basedir + "result.txt", "w")
    resultFile.write("yes")
    resultFile.close()

except Exception as e:
    print e
    exceptFile = open(basedir + "exception.txt", "w")
    resultFile = open(basedir + "result.txt", "w")
    resultFile.write("no")
    resultFile.close()
    exceptFile.write(str(e))
    exceptFile.close()
