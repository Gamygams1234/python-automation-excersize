import openpyxl as xl

# loading the sheet
wb =xl.load_workbook('transactions.xlsx')

sheet = wb["Sheet1"]

# getting the cell of the sheet
cell = sheet["a1"]
# second way of getting the cell
cell = sheet.cell(1,1)
# this is showing the value of the cell in row 1 column 1
print(cell.value)


for row in range(2, sheet.max_row + 1):
    cell = sheet.cell(row,3)
    corrected_price = cell.value * 0.9
    corrected_price_cell = sheet.cell(row, 4)
    corrected_price_cell.value = corrected_price
    print(corrected_price)

wb.save('transactions2.xlsx')


