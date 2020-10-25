##

import openpyxl
wb = openpyxl.reader.excel.load_workbook(filename="karamel.xlsx")
print(wb.sheetnames)
wb.active = 0

sheet = wb.active
#print(sheet['A1'].value)

for i in range(6, 7):
    print(sheet['AR'+str(i)].value)
