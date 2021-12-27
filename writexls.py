import bpy
import xlrd
import os
import math
from easybpy import *
import xlsxwriter

if object_exists("excel"):

    obj = get_object("excel")
    verts = get_vertices(obj)
    
    #print(verts[0].co[0])
    workbook = xlsxwriter.Workbook(r'E:\master\build\coordinateData\testData2.xlsx')
    worksheet = workbook.add_worksheet()
    
    # Start from the first cell below the headers.
    row = 0
    col = 0
    
    data = []
    for point in verts:
        data.append(point.co)
        
    for x, y, z in data:
        worksheet.write_number(row, col, x)
        worksheet.write_number(row, col + 1, y)
        worksheet.write_number(row, col + 2, z)
        row += 1
    
    
    workbook.close()