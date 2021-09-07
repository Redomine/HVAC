#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = '1.Импорт данных'
__doc__ = "Перенос расхода воздуха из расчетной таблицы в пространство модели"


import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

import sys
import WriteLog
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from rpw.ui.forms import select_file
from rpw.ui.forms import TextInput
from rpw.ui.forms import SelectFromList
from rpw.ui.forms import Alert

from Microsoft.Office.Interop import Excel
from System.Runtime.InteropServices import Marshal

exel = Excel.ApplicationClass()

project_todo = SelectFromList('Выберите тип таблицы', ['Вентиляция','Отопление','Кондиционирование'])


filepath = select_file()

try:
    workbook = exel.Workbooks.Open(filepath)
except Exception:    
    Alert('Файл не найден!', title= 'Ошибка', header = 'Неверный ввод')
    sys.exit()
sheet_name = TextInput('Введите имя листа с данными', default='Лист1')

#проверяем есть ли такой лист в книге
try:
    worksheet = workbook.Sheets[sheet_name]
except Exception:
    Alert('В файле нет листа с таким названием!', title= 'Ошибка', header = 'Неверный ввод')
    sys.exit()

xlrange = worksheet.Range["A1", "AZ500"]

scan_status = 'start'
current_row_number = 0

#опреляем имена для проверки в зависимости от типа таблицы
def cell_names(project_todo):
    names = []
    if project_todo == 'Вентиляция':
        names.append('Приток')
        names.append('Вытяжка')
    if project_todo == 'Отопление':
        names.append('Теплопотери')
        names.append('-')
    if project_todo == 'Кондиционирование':
        names.append('Q полное')
        names.append('-')
    return names


names = cell_names(project_todo)

#Логика такая - прогоняем каждую строку в таблице в поисках заголовка, как только его находим
#пробегаемся по строке-заголовку и забираем номера столбцов с номером помещения и расходами
#в моменте где мы читаем значения объединенной ячейки, заглавляющей две других - номер первой
#под объединенными будет тем же что и у заголовка, а у второй на один больше
current_row = []
while scan_status == 'start':
    if current_row_number > 300:
        Alert('Заголовки не обнаружены, проверьте имена в таблице!', title="Ошибка", header="Проблема с таблицей")
        sys.exit()
    current_row_number+=1
    current_row = []
    for cell in range(52):    
        current_row.append(xlrange.value2[current_row_number, cell])
    for cell in current_row:
        if cell == 'Помещение':
            space_column = current_row.index(cell)
            for cell in current_row:
                if cell == names[0]:
                    supply_column = current_row.index(cell)
                if cell == names[1]:
                    extract_column = current_row.index(cell)
            scan_status = 'stop'
    
    




#85-103 строка - собираем список помещений и прогоняем его элементы на предмет повторений
error_row = current_row_number

error_list = []
while True:
    if xlrange.value2[error_row, space_column] == None:
        break
    error_list.append(xlrange.value2[error_row, space_column])
    error_row += 1


error_alert = 'Помещения:'
alert_status = 'off'
for space in error_list:
    if error_list.count(space) > 1:
        alert_status = 'on'
        error_alert += space
        error_alert += ', '
if alert_status == 'on':
    Alert('В списке помещений есть повторяющиеся элементы!', title= 'Ошибка', header = error_alert)
    exel.ActiveWorkbook.Close(True)
    Marshal.ReleaseComObject(worksheet)
    Marshal.ReleaseComObject(workbook)
    Marshal.ReleaseComObject(exel)
    sys.exit()


spaces = []



counter = 0 #число прогонов цикла для корректного заполнения списка,
#номер прогона = номер вложенного списка

#далее присваиваем списку пространств данные вида помещение-расход
#до тех пор пока новая строка не окажется пустой
#если строка расхода пустая - присваиваем ноль
try:
    while True:
        if xlrange.value2[current_row_number, space_column] == None:
            break
        spaces.append([])
    
        spaces[counter].append(xlrange.value2[current_row_number, space_column])
    
        if (xlrange.value2[current_row_number, supply_column] == None) or (xlrange.value2[current_row_number, supply_column] == ''): spaces[counter].append(0)
        else: spaces[counter].append(xlrange.value2[current_row_number, supply_column])
    
        if project_todo == 'Вентиляция':
            if (xlrange.value2[current_row_number, extract_column] == None) or (xlrange.value2[current_row_number, extract_column] == ''): spaces[counter].append(0)
            else: spaces[counter].append(xlrange.value2[current_row_number, extract_column])
        current_row_number += 1
        counter += 1
except Exception:
    Alert('Ошибка в обработке списка помещений!', title= 'Ошибка', header = 'Проблемы с таблицей')
    exel.ActiveWorkbook.Close(True)
    Marshal.ReleaseComObject(worksheet)
    Marshal.ReleaseComObject(workbook)
    Marshal.ReleaseComObject(exel)
    sys.exit()
    
#закрытие невидимых процессов эксель, которые мусорят память
exel.ActiveWorkbook.Close(True)
Marshal.ReleaseComObject(worksheet)
Marshal.ReleaseComObject(workbook)
Marshal.ReleaseComObject(exel)

#начинаем работу с ревитом

doc = __revit__.ActiveUIDocument.Document

#получаем список пространств в активной модели
colSpaces = FilteredElementCollector(doc)\
                            .OfCategory(BuiltInCategory.OST_MEPSpaces)\
                            .WhereElementIsNotElementType()\
                            .ToElements()

t = Transaction(doc, 'Расходы воздуха')

t.Start()

#у ревита свои единицы измерения, переводим в них
if project_todo == 'Вентиляция':
    k = 101.940647731554 #Коэффициент для перевода в метры кубические.
else:
    k = 0.0929026687598116 #Коэффициент для перевода в ватты


#опреляем имена для проверки в зависимости от типа таблицы
def parameter_names(project_todo):
    names = []
    if project_todo == 'Вентиляция':
        names.append('ИОС_Расход воздуха приточный')
        names.append('ИОС_Расход воздуха вытяжной')
    if project_todo == 'Отопление':
        names.append('ИОС_Теплопотери')
        names.append('-')
    if project_todo == 'Кондиционирование':
        names.append('ИОС_Теплопритоки')
        names.append('-')
    return names

names = parameter_names(project_todo)

#далее перебираем список пространств, выдергивая их номера и
#сравнивая с номерами в списке полученном из экселя
#получив соответствие, из нужного вложенного списка забираем расходы
#и присваиваем их параметрам нужного пространства в модели

space_in_model = []

try:
    for Space in colSpaces:
        if Space.Location:
            if Space.LookupParameter("Номер"):
                SpaceNumber = Space.LookupParameter("Номер")
                SpaceNumber = SpaceNumber.AsString()
                space_in_model.append(SpaceNumber)
                for number in spaces:
                    if number[0] == SpaceNumber:
                        Supply = number[1]
                        if project_todo == 'Вентиляция':
                            Extract = number[2]
                        if Space.LookupParameter(names[0]):
                            SpaceSupply = Space.LookupParameter(names[0])
                            print(Supply)
                            SpaceSupply.Set(Supply/k)
                    
                        if project_todo == 'Вентиляция':
                            if Space.LookupParameter(names[1]):
                                SpaceExtract = Space.LookupParameter(names[1])
                                SpaceExtract.Set(Extract/k)

                    
                        if project_todo == 'Вентиляция':
                            print(SpaceNumber, Supply, Extract)
                        else: 
                            print(SpaceNumber, Supply)
except Exception:
    Alert('Ошибка при работе с моделью!', title= 'Ошибка', header = 'Проблемы с моделью')
    sys.exit()
    
for space in error_list:
    if space not in space_in_model:
        print space, 'Помещение есть в спецификации, но не было найдено в модели!'

t.Commit()

WriteLog.SetLogFile("1.Импорт данных", doc)