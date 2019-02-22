import openpyxl
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

#Setup image
im = Image.open("genericCowboy.jpg")
pixels = im.load()
width = im.size[0]
height = im.size[1]

#Setup worksheet
wb = Workbook()
ws = wb.active
ws.insert_cols(width)
ws.insert_rows(height)

normalizeCells(ws,height,width)


wb.save('test.xlsx')