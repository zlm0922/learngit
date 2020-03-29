#! /usr/bin/env python
# -*-coding:utf-8 -*-
# 依据坐标增量计算角度
90-math.degrees(math.atan2(( !END_Y! - !START_Y! ),( !END_X! - !START_X! )))

def quadrant(NorthAzimuth):

    if (NorthAzimuth >= 0) & (NorthAzimuth < 90):
        quad = "N" + str(NorthAzimuth) + " E"
    elif (NorthAzimuth >= 90) & (NorthAzimuth < 180):
        quad = "S " + str(180-NorthAzimuth) + "E"
    elif (NorthAzimuth >= 180) & (NorthAzimuth < 270):
        quad = "S " + str(NorthAzimuth-180)+"W"
    else:
        quad = "N " + str(360-NorthAzimuth) + "W"
    return quad


def quadrant(NorthAzimuth):

    if (NorthAzimuth >= 0) & (NorthAzimuth < 90):
        quad = str(NorthAzimuth)
    elif (NorthAzimuth >= 90) & (NorthAzimuth < 180):
        quad = str(180-NorthAzimuth)
    elif (NorthAzimuth >= 180) & (NorthAzimuth < 270):
        quad = str(NorthAzimuth-180)
    else:
        quad = str(360-NorthAzimuth)
    return quad


def quadrant(NorthAzimuth):
    if NorthAzimuth < 0:
        quad = NorthAzimuth + 360
    elif NorthAzimuth >= 360:
        quad = NorthAzimuth - 360

    return quad
