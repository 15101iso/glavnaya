import openpyxl
import xlrd
from openpyxl import Workbook
wb = Workbook()

# Load the workbook
book = xlrd.open_workbook('C:\glavnaya\BaseMRT.xls')
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))

sh = book.sheet_by_index(0)

#print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
#print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))

dictionary = dict()
for rx in range(sh.nrows):
    #print(sh.row(rx))
    row_index = 0
    for cell in sh.row(rx):
        if (row_index == 6):
                #print(cell.value)
                atr = "/"
                
                for j in str(cell.value).split(" "):
                    if str(j).find(atr) >= 0:
                        
                        datafind = sh.row(rx)

                        findKey = False

                        for x in dictionary.keys():
                            if x == str(j):
                                findKey = True
                                break

                        if findKey:
                            arrData = dictionary.get(j)
                            arrData.append(datafind)
                        else:
                            arrData = []
                            arrData.append(datafind)
                            dictionary[j] = arrData
        row_index += 1
print(dictionary)


