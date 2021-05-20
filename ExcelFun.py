# -*- coding: utf-8 -*-
"""
Created on Tue May 11 16:27:12 2021

@author: noneofyourbusiness
"""
import xlsxwriter
from PIL import Image

################INITIATE#################



# creating a image object
im = Image.open(r"C:\Program Files\test.JPG") #####<-CHANGE PATH #####
im.mode #should be RGB
px = im.convert('RGB').load()
im_width, im_height = im.size #im.size gives length of width and height, so e.g. 0-9 is width 10

#initiate Excel workbook
workbook = xlsxwriter.Workbook(r"C:\Program Files\Intel\test.xlsx") ######<- CHANGE PATH ######
worksheet = workbook.add_worksheet()


#######################FUNCTIONS##########

####convert RGB to HEX
def rgb2hex(r, g, b):
    return '#{:02x}{:02x}{:02x}'.format(r, g, b)


####Function to convert a given number to an Excel column
def getColumnName(n):
    # initialize output string as empty
    result = "" 
    while n > 0:
        # find the index of the next letter and concatenate the letter
        # to the solution
        # here index 0 corresponds to `A`, and 25 corresponds to `Z`
        index = (n - 1) % 26
        result += chr(index + ord('A'))
        n = (n - 1) // 26
    return result[::-1]


###Set color to Excel cell
def setCellColor(rrggbb, cell_coord):
    cell_format = workbook.add_format()
    cell_format.set_bg_color(rrggbb)
    worksheet.write(cell_coord, "", cell_format)
    del cell_format


##############RUN##############

for x in range(im_width):
        for y in range(im_height):
            r, g, b = px[x, y]
            hex_col = rgb2hex(r, g, b)
            col_name = getColumnName(x+1)
            row_num = y+1
            cell_coord = col_name + str(row_num)
            setCellColor(hex_col, cell_coord)

worksheet.set_column(0, im_width-1, 2.5) #make cells approximately square-sized
workbook.close()
