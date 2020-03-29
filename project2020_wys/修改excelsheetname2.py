#!/usr/bin/env python
# -*- coding:utf-8 -*-
from win32com.client import Dispatch
import win32com.client
import arcpy
import os
import sys

import time
reload(sys)
sys.setdefaultencoding('utf-8')


def getSheetCount():
    '''get the number of sheet'''
    return xlbook.Worksheets.Count


# def getMaxRow( sheet):
#     '''get the max row number, not the count of used row number'''
#     return getSheet(sheet).Rows.Count
#
#
# def getMaxCol( sheet):
#     '''get the max col number, not the count of used col number'''
#     return getSheet(sheet).Columns.Count

def getRow(sheet=1, row=1):
    '''get the row object'''
    assert row > 0, 'the row index must bigger then 0'
    return getSheet(sheet).Rows(row)


def getCol(sheet, col):
    '''get the column object'''
    assert col > 0, 'the column index must bigger then 0'
    return getSheet(sheet).Columns(col)


def getCellValue(sheet, row, col):
    '''Get value of one cell'''
    return getCell(sheet, row, col).Value


def getColsIndex(activeshet, rowindex, kn_value):
    '''get colindex when cellvalue knowed'''
    colindex = None
    for col in range(1, activeshet.UsedRange.Columns.Count + 1):
        if activeshet.Cells(rowindex, col).Value == kn_value:
            colindex = col
            break
    return colindex


def setCellformat(sheetname, row, col):  # 设置单元格的数据
    "set value of one cell sheet ；sheetname"
    sht = xlbook.Worksheets(sheetname)
    # print sht.Cells(0,0).value
    sht.Cells(row, col).Font.Size = 10  # 字体大小
    sht.Cells(row, col).Font.Bold = True  # 是否黑体
    sht.Cells(row, col).Font.Name = u"宋体"  # 字体类型
    # sht.Cells(row, col).Interior.ColorIndex = 3  # 表格背景
    # sht.Range("A1").Borders.LineStyle = xlDouble
    # sht.Cells(row, col).BorderAround(1, 4)  # 表格边框

    # sht.Rows(row).RowHeight = 9  # 行高
    # sht.Cells(row, col).HorizontalAlignment = -4131  # 水平居中xlCenter
    # sht.Cells(row, col).VerticalAlignment = -4160  #


def getSheet(sheet=1):
    '''get the sheet object by the sheet index'''

    assert sheet > 0, 'the sheet index must bigger then 0'

    return xlbook.Worksheets(sheet)


def getCell(sheet=1, row=1, col=1):
    '''get the cell object'''
    assert row > 0 and col > 0, 'the row and column index must bigger then 0'
    return getSheet(sheet).Cells(row, col)


def setCellValue(sheet, row, col):
    '''set value of one cell'''
    # getCell(sheet, row, col).Value = value
    getCell(sheet, row, col).Font.Size = 9
    getCell(sheet, row, col).Font.Name = u'宋体'


def getRowValue(sheet, row):
    '''get the row values'''
    return getRow(sheet, row).Value


def inserColumns(activeshet, col):
    sht = activeshet  # xlBook.Worksheets(sheet)
    sht.Columns(col).Insert(1)


if __name__ == '__main__':
    arcpy.env.workspace = r'D:\智能化管线项目\雨污水管线数据处理\JLSHYWS_kkgdb\JLSHYWS_1.gdb'
    # 数据表excel所在文件path
    wks = r'D:\智能化管线项目\雨污水管线数据处理\雨污水采集模板\污雨水模板整理0318'.decode('utf-8')
    # global xlbook
    # app = xw.App(visible=True, add_book=False)
    # app.display_alerts = False  # 关闭一些提示信息，可以加快运行速度。 默认为 True。
    # app.screen_updating = True
    try:

        for root, files, filenames in os.walk(wks):
            for filename in filenames:
                xlapp = Dispatch('Excel.Application')
                xlpath = os.path.join(root, filename)
                # print xlpath
                if not os.path.basename(xlpath).startswith('~$'):
                    xlbook = xlapp.Workbooks.Open(xlpath)
                    pre_filename = os.path.splitext(filename)[0]
                    # wb = app.books.active
                    # wb = app.books.open(xlpath)
                    sheet0 = getSheet(1)
                    # print type(getRow(1,1))
                    # print getRowValue(1,1)[0][0]
                    #
                    # print '-----'
                    # inserColumns(sheet0, getColsIndex(sheet0, 4, '备注'))# 插入列，可以指定列号
                    usedrange = sheet0.UsedRange
                    # sht = wb.sheets[sheet0.Name]
                    print sheet0.Name
                    # print '运行状态 所在列号:',getColsIndex(sheet0,4,'运行状态')
                    for col in range(1, usedrange.Columns.Count + 1):
                        print col
                        # color4 = getCell(1, 4, col).Font.ColorIndex
                        # print color4
                        # getCell(1, 4, col).clearformats()
                        getCell(1, 4, col).Font.Size = 10
                        getCell(1, 4, col).Font.Name = u'微软雅黑'
                        # getCell(1, 4, col).Font.ColorIndex = color4
                        # getCell(1, 4, col).autofit()
                        getCell(1, 5, col).Font.Size = 10

                        getCell(1, 5, col).Font.Name = 'Tahoma'
                        getRow(1, 4).RowHeight = 15
                        getRow(1, 5).RowHeight = 15
                        getRow(1, 1).RowHeight = 28.5
                        getRow(1, 2).RowHeight = 12.75
                        # sheet0.autofit()
                        setCellValue(1, 3, col)
                        setCellformat(sheet0.Name, 2, col)
                        # sht.autofit()
                        # getCell(1, 4, getColsIndex(sheet0,4,'备注')-1).Value = u'运行状态'
                        # getCell(1, 5, getColsIndex(sheet0, 4, '备注')-1).Value = 'ppp'
                        # getCell(1, 4,getColsIndex(sheet0,4,'备注')-1).Font.ColorIndex =3 #5

                    for i in range(1, xlbook.Worksheets.Count + 1):

                        # print i
                        sheet = getSheet(i)
                        print sheet.Name
                        if sheet.Name.startswith('Dom'):
                            sheet.Name = 'Domain'

                    xlbook.Save()
                    #
                    xlbook.Close(SaveChanges=0)
                    del xlapp
                    # time.sleep(3)

    except Exception as e:
    #     # print e[1].decode('cp936')
        print e
    # finally:
    #     if xlbook:
    #         # xlbook.Save()
    #
    #         xlbook.Close(SaveChanges=0)
    #         del xlapp
