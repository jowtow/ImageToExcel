import openpyxl
import sys
from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.cell import Cell
from PIL import Image

def getColumnStr(num):
    columnStr = ''
    headerNum = num / 26
    while(headerNum >= 1):
        columnStr = columnStr + chr(int(headerNum%26) + 64)
        headerNum = headerNum / 26
    columnStr = columnStr + chr(num % 26 + 65)
    return columnStr

def normalizeCells(ws,height,width):
    for i in range(width):
        ws.column_dimensions[getColumnStr(i)].width = 1
    for i in range(1,height):
        ws.row_dimensions[i].height = 5

def getSensitivityPixels(pixels,height,width,pixelation):
    newPixels = []
    for x in range(width):
        pixelRow = []
        for y in range(height):
            pixel = ''
            r = []
            g = []
            b = []
            
            for i in range(pixelation):
                x1 = x*pixelation
                for j in range(pixelation):
                    y1 = y*pixelation
                    rgb = pixels[x1 + i, y1 + j] 
                    r.append(rgb[0])
                    g.append(rgb[1])
                    b.append(rgb[2])
            rAvg = int(sum(r) / float(len(r)))
            gAvg = int(sum(g) / float(len(g)))
            bAvg = int(sum(b) / float(len(b)))
            pixelRow.append((rAvg,gAvg,bAvg))
        newPixels.append(pixelRow)
    return newPixels


def fillCellsWithPixels(ws,pixels,height,width):
    for i in range(1,width):
        x = getColumnStr(i)
        for y in range(1,height):
            hexColor = '%02x%02x%02x' % pixels[i][y]
            fill1 = PatternFill(fill_type='solid',start_color=hexColor,end_color=hexColor)
            ws.cell(row=y,column=i).fill = fill1

pixelation = 1
if(len(sys.argv) > 2):
    pixelation = int(sys.argv[2])
#Setup image
im = Image.open(sys.argv[1])
pixels = im.load()
width = int(im.size[0] / pixelation)
height = int(im.size[1] / pixelation)
pixels = getSensitivityPixels(pixels,height,width,pixelation)

#Setup worksheet
wb = Workbook()
ws = wb.active
ws.insert_cols(width)
ws.insert_rows(height)

normalizeCells(ws,height,width)
fillCellsWithPixels(ws,pixels,height,width)

wb.save('test.xlsx')