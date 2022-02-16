#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Обновление креплений'
__doc__ = "Обновляет число подсчетных элементов"


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


# from Microsoft.Office.Interop import Excel
# from System.Runtime.InteropServices import Marshal
# from rpw.ui.forms import select_file
# from rpw.ui.forms import TextInput
# from rpw.ui.forms import SelectFromList
# from rpw.ui.forms import Alert
#
#
# exel = Excel.ApplicationClass()
#
# filepath = select_file()
#
# try:
#     workbook = exel.Workbooks.Open(filepath)
# except Exception:
#     Alert('Файл не найден!', title= 'Ошибка', header = 'Неверный ввод')
#     sys.exit()
# sheet_name = TextInput('Введите имя листа с данными', default='Лист1')
#
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
loc = XYZ(0, 0, 0)

Names = ['Шпилька, резьба M8, оцинкованная']

for element in colPipes:
    if str(element.Name) == 'Тр_Труба_Стальная Оцинкованная_Грувлок':
        new_name = 'Хомута под грувлок на трубу наружным диаметром ' + str(element.LookupParameter('Внешний диаметр').AsDouble() * 304.8)
        if new_name not in Names:
            Names.append(new_name)

for element in famtypeitr:
    famtypeID = element
    famsymb = doc.GetElement(famtypeID)
    if famsymb.Family.Name == '_Заглушка для спецификаций':
        temporary = famsymb

def bracing_curves_v2(collection):
    for element in collection:
        try:
            if element.LookupParameter('Диаметр'):
                dy = element.LookupParameter('Диаметр').AsValueString()
            if element.LookupParameter('Эквивалентный диаметр'):
                dy = element.LookupParameter('Эквивалентный диаметр').AsValueString()
                dy = float(dy)
            if element.LookupParameter('Площадь'):
                long = element.LookupParameter('Площадь').AsDouble()
                long = float(long) * 0.0928886438809261
                
            if dy < 251:
                kg = 0.712
            elif dy < 561:
                kg = 1.22
            else:
                kg = 2.55
            
            if element.LookupParameter('Количество креплений'):
                element.LookupParameter('Количество креплений').Set(long*kg)
        except Exception:
            pass

def bracing_pipes(collection):
    try:
        for element in collection:
            if element.LookupParameter('Диаметр'):
                dy = element.LookupParameter('Диаметр').AsDouble() * 304.8
            if element.LookupParameter('Площадь'):
                long = (element.LookupParameter('Длина').AsDouble() * 304.8)/1000

            if dy < 16:
                kg = 0.72
            elif dy < 21:
                kg = 0.55
            elif dy < 26:
                kg = 0.625
            elif dy < 33:
                kg = 0.7
            elif dy < 41:
                kg = 0.72
            elif dy < 51:
                kg = 0.81
            elif dy < 71:
                kg = 0.85
            elif dy < 81:
                kg = 1.07
            elif dy < 101:
                kg = 1.33
            elif dy < 126:
                kg = 1.48
            else:
                kg = 1.76
            if element.LookupParameter('Количество креплений'):
                element.LookupParameter('Количество креплений').Set(long*kg)
    except Exception:
        pass

def getLenght(Type_Name, System_Name):
    lenght = 0
    for element in colPipes:
        if str(element.Name) == Type_Name and element.LookupParameter('ADSK_Имя системы').AsString() == System_Name:
            lenght = lenght + (element.LookupParameter('Длина').AsDouble() * 304.8)/1000
    return lenght

def new_position():
    #создаем заглушки по количеству систем, для каждой из которых по заглушке на расчетный параметр
    for system in ADSK_System_Names:
        for name in Names:
            familyInst = doc.Create.NewFamilyInstance(loc, temporary, Structure.StructuralType.NonStructural)

    colModel = make_col(BuiltInCategory.OST_GenericModel)
    Models = []
    for element in colModel:
        if element.LookupParameter('Семейство').AsValueString() == '_Заглушка для спецификаций':
            Models.append(element)

    for system in ADSK_System_Names:
        for name in Names:
            if name == 'Хомута под грувлок':
                gruvlock_lenght = getLenght('Тр_Труба_Стальная Оцинкованная_Грувлок', system)
                Number = round((gruvlock_lenght / 2), 0)
                Maker = 'Сантехкомплект'
            if name == 'Шпилька, резьба M8, оцинкованная':
                Number = 20
                Maker = 'Сантехкомплект'
            element = Models[0]
            element.LookupParameter('ADSK_Имя системы').Set(system)
            element.LookupParameter('ФОП_ВИС_Группирование').Set('7. Материалы креплений')
            element.LookupParameter('ФОП_ВИС_Наименование комбинированное').Set(name)
            element.LookupParameter('ADSK_Завод-изготовитель').Set(Maker)
            element.LookupParameter('ФОП_ВИС_Число').Set(Number)
            Models.pop(0)

ADSK_System_Names = []
System_Named = True

for element in colPipes:
    if element.LookupParameter('ADSK_Имя системы').AsString() == None:
        System_Named = False
        continue
    if element.LookupParameter('ADSK_Имя системы').AsString() not in ADSK_System_Names:
        ADSK_System_Names.append(element.LookupParameter('ADSK_Имя системы').AsString())

if System_Named == False:
    print 'Есть элементы труб у которых не заполнен параметр ADSK_Имя системы, для них не произведен расчет'


with revit.Transaction("Обновление общей спеки"):
    bracing_curves_v2(colCurves)
    bracing_pipes(colPipes)


    #при каждом повторе расчета удаляем старые версии
    for element in colModel:
        if element.LookupParameter('Семейство').AsValueString() == '_Заглушка для спецификаций':
            doc.Delete(element.Id)

    # в следующем блоке генерируем новые экземпляры пустых семейств куда уйдут расчеты
    new_position()





