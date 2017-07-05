# coding=utf-8

import os
import arcpy
from arcpy import env
import arcpy.mapping as mapping
from datetime import datetime

path = r"D:\python\coords"  # 工作空间路径
excel = r"D:\python\ArcpyBook\Ch4\Book1.xlsx"  # Excel 文件路径
sheet = r"sheet1"  # Excel 表格名称
point = r"D:\python\ArcpyBook\Ch4\Export_Output_2.shp"  # 点数据路径
mxdpath = r"C:\ArcpyBook\Ch4\tojpg.mxd"  # 地图文档路径
name = r"C:\ArcpyBook\Ch4\xxxx.jpg"  # 输出图片路径及名称


def to_jpg():
    # 创建工作空间
    filename = datetime.now().strftime('%Y%m%d%H%M%S')
    if os.path.exists(path):
        pass
    else:
        os.makedirs(path)
    env.Workspace = path
    arcpy.CreateFileGDB_management(env.Workspace, filename + "data.gdb", "9.3")
    gdb = os.path.join(path, filename + "data.gdb")
    # 导入Excel 源数据
    dbf = os.path.join(gdb, "excle_to_dbf")
    arcpy.ExcelToTable_conversion(excel, dbf, sheet)
    # 导入点数据
    data = os.path.join(gdb, "data")
    arcpy.CopyFeatures_management(point, data)
    # 属性表连接
    data2 = os.path.join(gdb, "data2")
    arcpy.JoinField_management(data, "OBJECTID", dbf, "OBJECTID_1", "cahzhi")
    # 筛选关联数据
    arcpy.Select_analysis(data, data2, "\"cahzhi_1\" >0")
    # 插值分析生成栅格数据
    idw = os.path.join(gdb, "IDW")
    arcpy.CheckOutExtension("Spatial")
    outidw = arcpy.sa.Idw(data2, "cahzhi_1", "", "2", "VARIABLE 12", "")
    outidw.save(idw)
    # 替换地图文档数据源
    mxd = mapping.MapDocument(mxdpath)
    df = mapping.ListDataFrames(mxd, "Crime")[0]
    lyr1 = mapping.ListLayers(mxd, "data", df)[0]
    lyr2 = mapping.ListLayers(mxd, "idw4", df)[0]
    lyr1.replaceDataSource(gdb, "FILEGDB_WORKSPACE", "data2")
    lyr2.replaceDataSource(gdb, "FILEGDB_WORKSPACE", "IDW")
    # 另存新地图文档
    newpath = os.path.join(os.path.dirname(mxdpath), filename + "idw.mxd")
    mxd.saveACopy(newpath)
    del mxd
    # 输出图片
    mxd2 = mapping.MapDocument(newpath)
    title = mapping.ListLayoutElements(mxd2,"TEXT_ELEMENT","xxx")[0]
    title.text = r"5月份热力图"
    mapping.ExportToJPEG(mxd2, name, resolution=200)
    del mxd2


if __name__ == '__main__':
    to_jpg()
