# Python program to read an excel file

# import openpyxl module
import openpyxl

file1 = open("MyFile.txt","w")
path = input('Please enter path of excel file')

# Give the location of the file
if (path == ''):
    path = "C:\\Users\\wareh\\OneDrive\\Masaüstü\\INVENTORY.xlsx"
blacklist = ['96PF250', '96PF444', '96PF251', '96GT15P', '96BOX004', '96FM003A', '96FM003P'
            '96FM005A', '96FM005P']
# To open the workbook
# workbook object is created
print('Loading the file... Please wait')
wb_obj = openpyxl.load_workbook(path)
print(wb_obj.sheetnames)

def inventory():
    wb_obj.active = 2
    sheet_obj = wb_obj.active
    treshhold_min = input('Please enter minimum quantity on location: ')


    for i in range(1, len(list(sheet_obj.rows))):
        cell_obj = sheet_obj.cell(row = i, column = 1)
        cell_obj_v = cell_obj.value
        cell_obj_next = sheet_obj.cell(row = i+1, column = 1)
        cell_obj_next_v = cell_obj_next.value
        if (cell_obj_v == cell_obj_next_v):
            if (int((sheet_obj.cell(row = i, column = 4)).value) <= int(treshhold_min) ):
                print(str(cell_obj_v) +
                    ' at location ' + str(((sheet_obj.cell(row = i, column = 3)).value)) +
                    ' have ' + str(((sheet_obj.cell(row = i, column = 4)).value)) + ' avaible')
    while(True):
        menu = input('Please enter 0 to go to main menu')
        if (menu == '0'):
            break
        else:
            print('Wrong action, please try again')

def replenishment():
    wb_obj.active = 2
    sheet_obj = wb_obj.active
    treshhold_min = input('Please enter minimum quantity on location: ')
    threshold_max = input('Please enter maximum quantity on location: ')
    file1.write('Minimum quanttity: ' + str(treshhold_min) + '\n')
    file1.write('Maximum quanttity: ' + str(threshold_max) + '\n')    
    black = False
    for i in range(1, len(list(sheet_obj.rows))):
        cell_obj = sheet_obj.cell(row = i, column = 1)
        cell_obj_v = cell_obj.value
        cell_obj_next = sheet_obj.cell(row = i+1, column = 1)
        cell_obj_next_v = cell_obj_next.value
        if (cell_obj_v == cell_obj_next_v):
            if (int((sheet_obj.cell(row = i, column = 4)).value) <= int(treshhold_min) ):
                result = int((sheet_obj.cell(row = i, column = 4)).value) + int((sheet_obj.cell(row = i+1, column = 4)).value)
                if (result <= int(threshold_max)):
                    for j in range(0, len(blacklist)):
                        if (blacklist[j] == str(cell_obj_v)):
                            black = True
                    if (black == False):
                        message = ('Transfer ' + str(cell_obj_v) +
                            ' from ' + str(((sheet_obj.cell(row = i, column = 3)).value)) +
                            ' to ' + str(((sheet_obj.cell(row = i+1, column = 3)).value)) +
                            ' it will be ' + str(result)  + ' total')
                        print(message)
                        file1.write((message + '\n'))
                        black = False

    while(True):
        menu = input('Please enter 0 to go to main menu')
        if (menu == '0'):
            break
        else:
            print('Wrong action, please try again')

while(True):

    print('1 - Inventory')
    print('2 - Replenishment')
    print('3 - Receiving')
    print('9 - Close application')
    print('0 - Main menu')
    menu = input('Please enter action:')
    if (menu == '9'):
        file1.close()
        break
    elif (menu == '1'):
        inventory()
    elif (menu == '2'):
        replenishment()
    else:
        print('Wrong Action, please try again')
