import os
import xlrd
book = xlrd.open_workbook('C:/1/l01gcal-poewmuh/imi2018.xls')
sh = book.sheets()[8]
for i in range(38):
        print (sh.cell(i,8).value)
        print (sh.cell(i,10).value)