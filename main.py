import openpyxl
import re

excel_file = openpyxl.load_workbook('Table.xlsx')

# print(excel_file.sheetnames) # ['TDSheet'] - получение название листа
employees_sheet = excel_file['TDSheet']
currently_active_sheet = excel_file.active

maxRow = employees_sheet.max_row
maxColumn = employees_sheet.max_column
listName = list()
strCode = list()

for i in range(maxRow - 9):
    i = i + 9
    i = str(i)
    m = employees_sheet["A" + i].value
    strCode.append(str(m))  # - заполняем первый столбец

for k in range(len(strCode)):
    listName.append(("list" + str(k)))  # создаем массив всех элементов
    listName[k] = [1, 2, 3, 4]
f = 0
for j in range(len(strCode)):
    if not re.search(r'\d{2}.\d{2}\D{3}', strCode[j]) is None:  # делаем нужную выборку
        listName[f][0] = strCode[j]
        #print(j)
        listName[f][1] = employees_sheet["A" + str(j + 11)].value
        listName[f][2] = employees_sheet["A" + str(j + 12)].value
        listName[f][3] = employees_sheet["A" + str(j + 13)].value
        # m = employees_sheet["A" + j].value
        f = f + 1

for i, e in reversed(list(enumerate(listName))):
    if str(listName[i][0]) == '1':
        del listName[i]
print(listName)


#listName.insert(1, 2)
#print(employees_sheet["A15"].value, employees_sheet["B15"].value, employees_sheet["C15"].value)

# print(len(listName))
# print(maxRow)
