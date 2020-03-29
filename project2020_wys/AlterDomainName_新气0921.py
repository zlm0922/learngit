#!/usr/bin/env python
# -*- coding:utf-8 -*-

from lstfc import list_feature,list_feature_batch
import arcpy
import os


def domain_dictionary(gdb,domaintype):
    # arcpy.env.workspace = r"D:\全生命周期"
    #def alter_domain_type(database):
    domain_dic = {}
    for domain in arcpy.da.ListDomains(gdb): # "Pipe.gdb"
        # print domain.name
        if domain.name == 'gn_Reportstatus':  # 限制
            print 112,domain.type
            if domain.type.upper() == domaintype and domain.domainType == 'CodedValue': # "Long"设置限制条件

                print 113

                domain_lst = []
                domain_lst.append(domain.description)
                domain_lst.append(domain.codedValues)#字典

                domain_dic[domain.name] = domain_lst
            #domain_lst.append(domain.codedValues)
            elif domain.domainType == 'Range':  # "Long"设置限制条件
                domain_lst = []
                domain_lst.append(domain.description)
                domain_lst.append(domain.Range)  # 字典

                domain_dic[domain.name] = domain_lst
    return domain_dic


def save_defaultvalue_domain():  #保存
    #a = domain_dictionary()
    all_lst = []
    del_domain = []
    ds_dic ={}
    for ds in list_feature_batch('Pipe.gdb'):
        print ds
        arcpy.env.workspace = ds
        fc_dic = {}
        for fc in arcpy.ListFeatureClasses(): # 为要素类是使用
            fld_dic = {}
            print "正在执行函数"
            print fc + '\n'
            for fld in arcpy.ListFields(fc):
                if fld.domain:
                    #try:
                        if fld.type == 'Integer': # 读取字段类型时,长整型：integer   字符串：String  短整型：SmallInteger
                            domname_defvalue = []
                            default_val = fld.defaultValue
                            domain_name = fld.domain
                            domname_defvalue.append(default_val)
                            domname_defvalue.append(domain_name)
                            print "\t正在执行函数"
                            print '\t',domain_name,default_val
                            if domain_name not in del_domain:
                                   del_domain.append(domain_name)
                            #arcpy.AssignDefaultToField_management(fc,fld.name,"","","True")#有默认值的情况
                            # 去除域值，去除default value
                            try:
                                arcpy.AssignDefaultToField_management(fc, fld.name, "", "", "true")  # 去除default value
                                arcpy.RemoveDomainFromField_management(fc, fld.name)  # 去除域值
                                arcpy.AlterField_management(fc,fld.name,'','','Short')
                            except arcpy.ExecuteError:
                                print arcpy.GetMessages()
                        #fld_dic[fld.name] = domname_defvalue
                            #print domname_defvalue
                        #fldname_domname_defvalue.append(fld_dic)
                            fld_dic[fld.name] = domname_defvalue
           # print fld_dic
            fc_dic[fc] = fld_dic
        ds_dic[os.path.basename(ds)] = fc_dic
        #print fc_dic
    all_lst.append(ds_dic)
    all_lst.append(del_domain)
    return all_lst


#arcpy.DeleteDomain_management(outdataset,fld.domain)
# 创建新的domain.type

# for key,val in domain_dictionary().items():
#     if domain_name == key:
#         #创建域值
#         arcpy.CreateDomain_management(outdataset,domain_name,val[0],'Short','CODED')
#         for cod,decs in val[1].items():
#             print cod,decs
#             arcpy.AddCodedValueToDomain_management(outdataset,domain_name,cod,decs)
#
#
# arcpy.AlterField_management(fc,fld.name,"","","Short")
# # print fld.aliasName,fld.type
# # print fld.domain,fld.type
# #arcpy.RefreshCatalog(fc)
# arcpy.AssignDomainToField_management(fc,fld.name,domain_name)
# #fld.defaultValue = default_val
# arcpy.AssignDefaultToField_management(fc, fld.name, default_val,"","false")


# 该语句的适用条件是要素类表现挂有的字段域值的重新创建
def alter_domiantype():
    outdataset = r'D:\智能化管线项目\新气\新气数据处理\新气数据入库_0916\新库0910.gdb'

    save_domain = domain_dictionary(outdataset,'LONG') #
    blst = save_defaultvalue_domain()
    print blst
    for del_nm in blst[1]:
        arcpy.DeleteDomain_management(outdataset,del_nm)
        for dmname,alias_codesc in save_domain.items():
            if del_nm == dmname:
                arcpy.CreateDomain_management(outdataset, del_nm, alias_codesc[0], 'Short', 'CODED')
                print "创建域值"
                for cod, desc in alias_codesc[1].items():
                    try:
                        arcpy.AddCodedValueToDomain_management(outdataset, del_nm, int(cod),desc)
                    except arcpy.ExecuteError:
                        print arcpy.GetMessages()
                        print cod, desc
                        print type(cod)


def alter_domiantype_single():

    outdataset = r'D:\智能化管线项目\新气\新气数据处理\新气数据入库_0916\新库0910.gdb'

    save_domain = domain_dictionary(outdataset, 'LONG')  #
    print save_domain

    arcpy.DeleteDomain_management(outdataset, "gn_Reportstatus")
    for dmname, alias_codesc in save_domain.items():
        # if del_nm == dmname:
            arcpy.CreateDomain_management(outdataset, "gn_Reportstatus", alias_codesc[0], 'Short', 'CODED')
            print "创建域值"
            for cod, desc in alias_codesc[1].items():
                try:
                    arcpy.AddCodedValueToDomain_management(outdataset, "gn_Reportstatus", int(cod), desc)
                except arcpy.ExecuteError:
                    print arcpy.GetMessages()
                    print cod, desc
                    print type(cod)

# 是对照修改前的恢复域值挂接
def recover_fld_domain():
    arcpy.env.workspace = r"D:\全生命周期"
    for ds in list_feature_batch('Pipe.gdb'):
        #print ds
        blst = save_defaultvalue_domain()
        for ds_item, fc_dic_item in blst[0].items():

            if ds_item == os.path.basename(ds):
                print ds
                arcpy.env.workspace = ds
                for fc_item,fld_dic in fc_dic_item.items():
                    print "正在执行主程序"
                    for fc in arcpy.ListFeatureClasses(): # 为要素类是使用
                        if fc == fc_item:

                            print fc + '\n'
                            for fld in arcpy.ListFields(fc):
                                for fld_item,domain_dflt_lst in fld_dic.items():
                                    if fld_item == fld.name:
                                        #fld.domain = domain_dflt_lst[1]
                                        try:
                                            arcpy.AssignDomainToField_management(fc,fld.name,domain_dflt_lst[1])
                                        except arcpy.ExecuteError:
                                            print arcpy.GetMessages()
                                            print 10 * '#'
                                            print "图层:" + fc + ' ' + '字段：' + fld.name + ' ' + "名称为"  + domain_dflt_lst[1] +'的阈值不存在'

                                        #fld.defaultValue = domain_dflt_lst[0]
                                        if domain_dflt_lst[0] is not None:
                                            try:
                                                arcpy.AssignDefaultToField_management(fc,fld.name,int(domain_dflt_lst[0]),"","false")
                                            except arcpy.ExecuteError:
                                                print arcpy.GetMessages()
                                                print 10*"%"
                                                print fc,fld.name,domain_dflt_lst[0]
                                                print 10 * "%"


def alter_domain_name():
    soucedomain = domain_dictionary('SPDM2_201803.gdb','Short')
    for dmname, alias_codesc in soucedomain.items():
        new_dmname = dmname[0:2]+'_' + dmname[2:]
        #if del_nm == dmname:
        if dmname[2] != '_':
            arcpy.CreateDomain_management(outdataset, new_dmname, alias_codesc[0], 'Short', 'CODED')
            print "创建域值"
            for cod, desc in alias_codesc[1].items():
                try:
                    arcpy.AddCodedValueToDomain_management(outdataset, new_dmname, int(cod), desc)
                except arcpy.ExecuteError:
                    print arcpy.GetMessages()
                    print cod, desc
                    print type(cod)
            arcpy.DeleteDomain_management(outdataset,dmname)


if __name__== "__main__":
    arcpy.env.workspace = r"D:\智能化管线项目\新气\新气数据处理\新气数据入库_0916\新库0910.gdb"
    outdataset = r"D:\智能化管线项目\新气\新气数据处理\新气数据入库_0916\新库0910.gdb"
    alter_domiantype_single()






