import openpyxl
wb = openpyxl.load_workbook('file.xlsx')
print(type(wb))
print(wb.sheetnames)
sheet = wb['S1']
print(sheet)
print(type(sheet))
print(sheet.title)
print(wb.active)
print("-------")
print(sheet['A1'])
print(sheet['A1'].value)
cell = sheet['B2']
print(cell.value)
print("Row " + str(cell.row) + ", column " + str(cell.column) + " has value " + str(cell.value) + ".")
print("Cell " + str(cell.coordinate) + ' has value ' + str(cell.value) + '.')
print("-------")
print(sheet.cell(row=1, column=2).value)
for i in range(1, 8, 2):
    print(i, sheet.cell(row=i, column=2).value)


