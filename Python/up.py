# -*- coding: utf-8 -*-
"""
Created on Mon Aug  6 08:42:53 2018

@author: Najib
"""

import os
import pandas as pd
import xlrd

book = xlrd.open_workbook(r"C:\Users\Najib\Documents\AGIT\Customer Insight\Data\AHM\1Januari2018.xlsx")
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))
for rx in range(sh.nrows):
    print(sh.row(rx))

#cwd = os.getcwd()
#cwd
#
#os.chdir(r"C:\Users\Najib\Documents\AGIT\Customer Insight\Data\AHM")
#os.listdir('.')
#
#file = '1.Januari 2018.xlsx'
#
#xl = pd.ExcelFile(file)
#
#print(xl.sheet_names)
#
#df1 = xl.parse('januari')
#
#print(df1)