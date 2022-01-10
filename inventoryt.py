# Python program to read an excel file

# import openpyxl module
import openpyxl

# Give the location of the file
path = "C:\\Users\\wareh\\OneDrive\\Masaüstü\\INVENTORY.xlsx"

# To open the workbook
# workbook object is created
wb_obj = openpyxl.load_workbook(path)

wb_obj.active = 2

# Get workbook active sheet object
# from the active attribute
sheet_obj = wb_obj.active

# Cell objects also have a row, column,
# and coordinate attributes that provide
# location information for the cell.

# Note: The first row or
# column integer is 1, not 0.

print(wb_obj.sheetnames)

treshhold = input('Please enter minimum quantity on location: ')

#sheet_obj = wb_obj.get_sheet_by_name('Pallet locations')

# Cell object is created by using
# sheet object's cell() method.
cell_obj = sheet_obj.cell(row = 1, column = 1)

for i in range(1, len(list(sheet_obj.rows))):
    cell_obj = sheet_obj.cell(row = i, column = 1)
    cell_obj_v = cell_obj.value
    cell_obj_next = sheet_obj.cell(row = i+1, column = 1)
    cell_obj_next_v = cell_obj_next.value
    if (cell_obj_v == cell_obj_next_v):
        if (int((sheet_obj.cell(row = i, column = 4)).value) <= int(treshhold) ):
            print(str(cell_obj_v) +
                ' at location ' + str(((sheet_obj.cell(row = i, column = 3)).value)) +
                ' have ' + str(((sheet_obj.cell(row = i, column = 4)).value)) + ' avaible')
# Print value of cell object
# using the value attribute

#print(sheets)

#print(list(sheet_obj.rows))
