#!usr/bin/env python
# -*- coding:utf-8 -*-
import arcpy
import os
import sys
from arcpy import env
# reload(sys)
# sys.setdefaultencoding('utf-8')

env.overwriteOutput = True
# env.workspace =r"D:\全生命周期"#D:\全生命周期\


def list_feature(gdb, dataset):
    #dir_name = os.getcwd(gdb)
    # env.workspace = r"D:\全生命周期"  # D:\全生命周期\
    # for fgdb in arcpy.ListWorkspaces("*gdb"):
    for fgdb in list(set(arcpy.ListWorkspaces("*gdb")) |
                     set(arcpy.ListWorkspaces("*.mdb"))):
        # print fgdb
        wk = os.path.basename(fgdb)

        if str(wk) == gdb:  # str("SPDM2_datasets.gdb"):
            print wk
            env.workspace = fgdb
            # print arcpy.env.workspace
            datasets = arcpy.ListDatasets()
            datasets = [''] + datasets if datasets is not None else []
            ## if datasets is not None:
                    #datasets = [''] + datasets
               #else:
                #    datasets = []

            # print datasets
            for ds in datasets:

                #print env.workspace
                if str(ds).strip() == str(dataset).strip():
                    # print ds
                    # env.workspace = fgdb + os.sep + ds
                    return fgdb + os.sep + ds
                    #  print env.workspace


def list_feature_batch(gdb):
    for fgdb in arcpy.ListWorkspaces("*gdb"):
        #print fgdb
        wk = os.path.basename(fgdb)
        if str(wk) == gdb:  # str("SPDM2_datasets.gdb"):
            print wk
            env.workspace = fgdb
            # print arcpy.env.workspace
            datasets = arcpy.ListDatasets()
            datasets = [''] + datasets if datasets is not None else []
            #print datasets
            wklst = []
            for ds in datasets:
                #print ds
                env.workspace = fgdb + os.sep + ds
                wk = fgdb + os.sep + ds
                #  print env.workspace
                wklst.append(wk)
            return wklst
