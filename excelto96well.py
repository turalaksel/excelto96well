#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Date    : 2017-07-16 15:48:50
# @Author  : Your Name (you@example.org)
# @Link    : http://example.org
# @Version : $Id$

import xlrd
import xlwt
import os
import sys
import numpy as np

# Checking the number of arguments
if len(sys.argv) < 2:
    print('Please enter the excel file to be processed')
    sys.exit()

# Checks if the file exists
fname = sys.argv[1]
if not os.path.isfile(fname):
    print('File doesnt exist. Please enter a valid file')
    sys.exit()

# Check if it is an excel file
header, ext = os.path.splitext(fname)

if not ext == '.xls':
    print('File doesnt have the valid format')
    sys.exit()

# Read the file
workbook = xlrd.open_workbook(fname)
worksheet = workbook.sheet_by_index(0)

# Read the cytokine labels
num_conditions  = len(worksheet.row(7))
name_conditions = []

for col in range(3, num_conditions):
    name_conditions.append(worksheet.cell(7, col).value)
num_conditions -= 3

# Read the well locations
empty_cell      = False
current_row     = 9
rowto96well_map = {}
while not empty_cell:
    # Get cell type in second column
    cell_type = worksheet.cell_type(current_row, 1)

    if cell_type == 0:
        break
    
    # Get the current value from third column
    current_str = worksheet.cell_value(current_row, 2)

    # Find the position of Well
    start_index = current_str.find('Well')

    # Check if the word Well is found 
    if start_index == -1:
        # Update current row position
        current_row += 1
        continue 
    else:
        well_position = current_str[start_index+5:-1]
        
        # Determine row id
        row_id    = ord(well_position[0]) - ord('A')

        # Column id
        column_id = int(well_position[1:])-1

        # Store row and column ids
        rowto96well_map[current_row] = (row_id,column_id)

        # Update current row position
        current_row += 1

# Read the data
data = {}
for i in range(num_conditions):
    data[i] = np.zeros((8,12),dtype=np.dtype('S25'))
    col     = i + 3
    for key in rowto96well_map:
        # Get the row and col id
        row_id, col_id = rowto96well_map[key]

        # Assign the values
        data[i][row_id, col_id] = worksheet.cell_value(key, col)


# Write the data
head, ext  = os.path.splitext(fname)
out_fname = head+'_96well_format.xls'
workbook = xlwt.Workbook()

for i in range(num_conditions):
    sheet = workbook.add_sheet(name_conditions[i])
    for row in range(8):
        for col in range(12):
            sheet.write(row, col, data[i][row, col])

workbook.save(out_fname)
