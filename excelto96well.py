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

#Checking the number of arguments
if len(sys.argv) < 2:
    print 'Please enter the excel file to be processed'
    sys.exit()

#Checks if the file exists
fname = sys.argv[1]
if not os.path.isfile(fname):
    print 'File doesnt exist. Please enter a valid file'
    sys.exit()

#Check if it is an excel file
header,ext = os.path.splitext(fname)

if not ext == '.xls':
    print 'File doesnt have the valid format' 
    sys.exit()

#Read the file
workbook = xlrd.open_workbook(fname)
worksheet = workbook.sheet_by_index(0)

#Read the cytokine labels
num_conditions  = len(worksheet.row(7))
name_conditions = [] 

for col in range(3,num_conditions):
    name_conditions.append(worksheet.cell(7,col).value)
num_conditions -= 3 

#Read the data
data = {}
for i in range(num_conditions):
    data[i] = []
    col     = i + 3
    for row in range(21,93):
        data[i].append(worksheet.cell(row,col).value)
    data[i] = np.array(data[i])
    data[i] = np.reshape(data[i],(8,9),order='F')

#Write the data
head,ext  = os.path.splitext(fname)
out_fname = head+'_96well_format.xls' 
workbook = xlwt.Workbook()

for i in range(num_conditions):
    sheet = workbook.add_sheet(name_conditions[i])
    for row in range(8):
        for col in range(9):
            sheet.write(row, col+3,data[i][row,col])

workbook.save(out_fname)

