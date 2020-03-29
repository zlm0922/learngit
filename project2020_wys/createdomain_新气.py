#!/usr/bin/env python
# -*- coding:utf-8 -*-
from lstfc import list_feature
import arcpy
import xlrd
import sys

reload(sys)
sys.setdefaultencoding('utf-8')

# 属性表包含多个域值项时
arcpy.env.overwriteOutput = True


def create_domain(nrows):
    # global flg
    flg = ''
    flg_bk = ''
    item_flg = ''
    for k in range(nrows):
        domains = arcpy.da.ListDomains(arcpy.env.workspace)
        domname_lst = [domain.name for domain in domains]
        if table.cell(k, 0).value == '':
            continue
        # print table.row_values(k)
        alst = table.row_values(k)
        if '' in alst:
            alst = list(set(alst))
            alst.remove('')
        # print alst

        if len(alst) > 2:  # 注意截取的字符串

            # item_flg = table.cell(k, 2).value

            # print flg+"1"
            try:
                if table.cell(k, 0).value not in domname_lst:
                    arcpy.CreateDomain_management(
                        arcpy.env.workspace, table.cell(k, 0).value,table.cell(k,1).value,
                        table.cell(k, 2).value.upper(),"CODED")
                    flg = table.cell(k, 0).value
                    # print flg
                    item_flg = table.cell(k,2).value.upper()
                    # print item_flg
                    flg_bk = ''

                else:
                    print "{} 该阈值名称已经存在！".format(table.cell(k, 0).value)
                    flg_bk = table.cell(k, 0).value
                    flg = ''

            except RuntimeError,arcpy.ExecuteError :
                print arcpy.GetMessages()
            # except arcpy.ExecuteError as err:
            #     print err
        elif len(alst) == 2:
            # print 10*"#"
            # print flg
            # print flg_bk
            # print 10 * "#"
            if flg_bk:
                # print flg_bk
                # print "因为阈值{}已存在，无需再为其添加阈值项！".format(flg_bk)
                continue
            elif flg:
                # print 222
                try:
                    # print 112
                    # print item_flg
                    # print 113
                    if item_flg == 'SHORT':
                        arcpy.AddCodedValueToDomain_management(arcpy.env.workspace,flg,
                                                               int(table.cell(k, 0).value),table.cell(k, 1).value)
                        print int(table.cell(k, 0).value)
                        print type(int(table.cell(k, 0).value)),table.cell(k, 1).value
                    else:
                        arcpy.AddCodedValueToDomain_management(arcpy.env.workspace, flg,
                                                               table.cell(k, 0).value, table.cell(k, 1).value)
                except RuntimeError:
                    print arcpy.GetMessages()


def add_domain():  # 挂阈值
    domain_flag = ''
    tb_flag = ""
    tbfld_flag = ""
    for k in range(nrows):
        if table.cell(k, 0).value == '':
            continue
        # print table.row_values(k)
        alst = table.row_values(k)
        if '' in alst:
            alst = list(set(alst))
            alst.remove('')
        if len(alst) == 1:
            tb_flag = table.cell(k,0).value
        elif len(alst) > 2:
            domain_flag = table.cell(k,0).value
            tbfld_flag = table.cell(k,3).value
            # print tb_flag
            # print "\t",domain_flag,tbfld_flag
            try:
                if tbfld_flag == "ENGINEETYPE":
                    arcpy.AlterField_management(tb_flag,tbfld_flag,new_field_alias="工程类型",field_type = "SHORT")
                elif tbfld_flag == "SYSDATAFROM":
                    arcpy.AlterField_management(tb_flag, tbfld_flag, field_type="SHORT")
                arcpy.AssignDomainToField_management(tb_flag, tbfld_flag, domain_flag)
            except RuntimeError:
                print arcpy.GetMessages()


if __name__ == "__main__":
    arcpy.env.workspace = r'Database Connections\Connection to 127.0.0.1.sde'
    arcpy.env.workspace = r'Database Connections\Connection to 10.246.146.120.sde'
    # arcpy.env.workspace = r"D:\智能化管线项目\新气\新气数据处理\新气数据入库_0916\新库0910zlm.gdb"
    # domains = arcpy.da.ListDomains(arcpy.env.workspace)
    # domname_lst = [domain.name for domain in domains]
    xlspath = r"D:\新库0910.xlsx".decode('utf-8')
    myworkbook = xlrd.open_workbook(xlspath)
    for i in range(myworkbook.nsheets):
        if myworkbook.sheet_names()[i] == "添加阈值5":
            table = myworkbook.sheet_by_index(i)  # 定位到sheet页
            print myworkbook.sheet_names()[i]
            ncols = table.ncols
            nrows = table.nrows
            create_domain(nrows)
    else:
        for i in range(myworkbook.nsheets):
            if myworkbook.sheet_names()[i] == "添加阈值5":
                table = myworkbook.sheet_by_index(i)  # 定位到sheet页
                print myworkbook.sheet_names()[i]
                ncols = table.ncols
                nrows = table.nrows
                add_domain()

