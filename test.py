import openpyxl
import json
import re

excel_file = openpyxl.load_workbook('Table2.xlsx')

# print(excel_file.sheetnames) # ['TDSheet'] - получение название листа
employees_sheet = excel_file['TDSheet']
currently_active_sheet = excel_file.active

maxRow = employees_sheet.max_row
maxColumn = employees_sheet.max_column
maxposition = 0
listName = list()
strCode = list()
twoStolbesh = list()

for i in range(maxRow - 9):
    i = i + 10
    i = str(i)
    m = employees_sheet["C" + i].value
    twoStolbesh.append(str(m))
    n = employees_sheet["A" + i].value
    strCode.append(str(n))

for k in range(len(twoStolbesh)):
    listName.append(("list" + str(k)))  # создаем массив всех элементов
    listName[k] = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]

f = 0
for j in range(len(twoStolbesh)):
    if str(twoStolbesh[j]) != 'None':
        listName[f][0] = str(f + 1)
        listName[f][5] = twoStolbesh[j]
        listName[f][6] = employees_sheet["J" + str(j + 10)].value
        listName[f][7] = employees_sheet["K" + str(j + 10)].value
        listName[f][8] = employees_sheet["M" + str(j + 10)].value
        listName[f][9] = employees_sheet["O" + str(j + 10)].value
        listName[f][10] = employees_sheet["P" + str(j + 10)].value
        listName[f][11] = employees_sheet["Q" + str(j + 10)].value
        listName[f][12] = str(employees_sheet["R" + str(j + 10)].value)
        listName[f][13] = str(employees_sheet["S" + str(j + 10)].value)
        listName[f][14] = str(employees_sheet["T" + str(j + 10)].value)
        f = f + 1

for i, e in reversed(list(enumerate(listName))):
    if str(listName[i][5]) == '5':
        del listName[i]
# for k in range(len(listName)):
#     NameStolb = [
#         "ID",
#         "Категория",
#         "Номер",
#         "ФИО ответственного",
#         "Отдел",
#         "Основное средство",
#         "Инвентарный номер",
#         "ОКОФ",
#         "Амортизационная группа",
#         "Способ начисления амортизации",
#         "Дата ввода в эксплуатацию",
#         "Состояние",
#         "Срок полезного использования",
#         "Мес. норма износа,",
#         "Износ"]
#     fruit_dictionary = dict(zip(NameStolb, listName[k]))
#     with open("data_file.json", "a") as write_file:
#         json.dump(fruit_dictionary, write_file, ensure_ascii=False, indent=2)

collposit = list()
collp = 0
for m in range(len(strCode)):
    k = str(strCode[m])
    if not re.search(r'\d{2}.\d{2}\D{3}', strCode[m]) is None:
        collposit.append(collp)
        collposit.append((strCode[m]))
        collp = 0
    else:
        collp = collp + 1
del collposit[0]

shagi = 11
shetchik = 0
for i in range(1, len(collposit), 2):
    shagi = shagi + collposit[i] - 1
    shagi1 = employees_sheet["A" + str(shagi)].value
    #print(shagi1)
    print(collposit[shetchik])
    shetchik = shetchik + 2



#print(strCode)
print(collposit)
#print(listName)
