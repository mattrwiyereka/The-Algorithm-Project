Python 2.7.14 (v2.7.14:84471935ed, Sep 16 2017, 20:25:58) [MSC v.1500 64 bit (AMD64)] on win32
Type "copyright", "credits" or "license()" for more information.
>>> import xlrd
>>> file_location = "C:\Users\Noella\Desktop/data.xlsx"
>>> workbook = xlrd.open_workbook(file_location)
>>> sheet = workbook.sheet_by_index(0)
>>> sheet.cell_value(0, 0)
u''
>>> sheet.cell_value(1, 1)
u'Name1'
>>> sheet.cell_value(1, 3)
u'Age'
>>> sheet.nrows
7
>>> sheet.ncols
6
>>> for col in range (sheet.ncols):
	print sheet.cell_value(1, col)

	

Name1
Name2
Age
ID
Lost item
>>> for col in range (sheet.ncols):
	print sheet.cell_value(2, col)

	

Mathieu
RWIYEREKA
30.0
123.0
Phone
>>> 
