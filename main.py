import openpyxl
wb = openpyxl.load_workbook('file.xlsx')
print(type(wb))
print(wb.sheetnames)
sheet = wb['S1']
print(sheet)
print(type(sheet))
print(sheet.title)
print(wb.active)

