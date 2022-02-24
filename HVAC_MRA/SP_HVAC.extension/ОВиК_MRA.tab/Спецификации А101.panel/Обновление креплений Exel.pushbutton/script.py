#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Импорт ведомости \n элементов'
__doc__ = "Генерирует в модели элементы в соответствии с их ведомостью"


import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

import sys
import System

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit

from Microsoft.Office.Interop import Excel
from System.Runtime.InteropServices import Marshal
from rpw.ui.forms import select_file
from rpw.ui.forms import TextInput
from rpw.ui.forms import SelectFromList
from rpw.ui.forms import Alert

exel = Excel.ApplicationClass()
filepath = select_file()



doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView

def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col 
    
colPipes = make_col(BuiltInCategory.OST_PipeCurves)
colCurves = make_col(BuiltInCategory.OST_DuctCurves)
colModel = make_col(BuiltInCategory.OST_GenericModel)
# create a filtered element collector set to Category OST_Mass and Class FamilySymbol
collector = FilteredElementCollector(doc)
collector.OfCategory(BuiltInCategory.OST_GenericModel)
collector.OfClass(FamilySymbol)
famtypeitr = collector.GetElementIdIterator()
famtypeitr.Reset()



is_temporary_in = False
for element in famtypeitr:
    famtypeID = element
    famsymb = doc.GetElement(famtypeID)

    if famsymb.Family.Name == '_Заглушка для спецификаций':
        temporary = famsymb
        is_temporary_in = True

if is_temporary_in == False:
    print 'Не обнаружено семейство-заглушка для спецификаций, проверьте не менялось ли его имя или загружалось ли оно'
    sys.exit()


def setElement(element, name, setting):
    if setting == None:
        pass
    else: element.LookupParameter(name).Set(setting)


def new_position(calculation_elements):
    #создаем заглушки по элементов собранных из таблицы
    loc = XYZ(0, 0, 0)
    for element in calculation_elements:
        familyInst = doc.Create.NewFamilyInstance(loc, temporary, Structure.StructuralType.NonStructural)

    #собираем список из созданных заглушек
    colModel = make_col(BuiltInCategory.OST_GenericModel)
    Models = []
    for element in colModel:
        if element.LookupParameter('Семейство').AsValueString() == '_Заглушка для спецификаций':
            try:
                element.CreatedPhaseId = phaseid
            except Exception:
                print 'Не удалось присовить стадию спецификация, проверьте список стадий'

            Models.append(element)

    #для первого элмента списка заглушек присваиваем все параметры, после чего удаляем его из списка
    for element in calculation_elements:
        dummy = Models[0]
        setElement(dummy, 'ADSK_Имя системы', element[0])
        setElement(dummy, 'ФОП_ВИС_Группирование', element[1])
        setElement(dummy, 'ФОП_ВИС_Наименование комбинированное', element[2])
        setElement(dummy, 'ADSK_Завод-изготовитель', element[3])
        setElement(dummy, 'ФОП_ВИС_Единица измерения', element[4])
        setElement(dummy, 'ФОП_ВИС_Число', element[5])
        Models.pop(0)

ADSK_System_Names = []
System_Named = True

try:
    workbook = exel.Workbooks.Open(filepath)
except Exception:
    sys.exit()
sheet_name = 'Импорт'

try:
    worksheet = workbook.Sheets[sheet_name]
except Exception:
    print 'Не найден лист с названием Импорт, проверьте файл формы.'
    sys.exit()

xlrange = worksheet.Range["A1", "AZ500"]

ADSK_System = 0
FOP_Group = 1
FOP_Name = 2
ADSK_Maker = 3
ADSK_Izm = 4
FOP_Number = 5


with revit.Transaction("Добавление расчетных элементов"):
    #при каждом повторе расчета удаляем старые версии
    for element in colModel:
        if element.LookupParameter('Семейство').AsValueString() == '_Заглушка для спецификаций':
            doc.Delete(element.Id)

    calculation_elements = []
    row = 2
    while True:
        if xlrange.value2[row, FOP_Name] == None:
            break
        System = xlrange.value2[row, ADSK_System]
        Group = xlrange.value2[row, FOP_Group]
        Name = xlrange.value2[row, FOP_Name]
        Maker = xlrange.value2[row, ADSK_Maker]
        Izm = xlrange.value2[row, ADSK_Izm]
        Number = xlrange.value2[row, FOP_Number]
        row += 1
        calculation_elements.append([System, Group, Name, Maker, Izm, Number])


    for phase in doc.Phases:
        if phase.Name == 'Спецификация':
            phaseid = phase.Id

    # в следующем блоке генерируем новые экземпляры пустых семейств куда уйдут расчеты
    new_position(calculation_elements)


exel.ActiveWorkbook.Close(True)
Marshal.ReleaseComObject(worksheet)
Marshal.ReleaseComObject(workbook)
Marshal.ReleaseComObject(exel)








