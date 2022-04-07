#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Добавление параметров'
__doc__ = "Добавление параметров в семейство для за полнения спецификации и экономической функции"


import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

import sys
import System
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from rpw.ui.forms import SelectFromList
from System import Guid
from pyrevit import revit
from Autodesk.Revit.DB.Electrical import *

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView

def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col 

spFile = doc.Application.OpenSharedParameterFile()

#проверяем тот ли файл общих параметров подгружен
spFileName = str(doc.Application.SharedParametersFilename)
spFileName = spFileName.split('\\')
spFileName = spFileName[-1]

if "ФОП_v1.txt" != spFileName:
    try:
        doc.Application.SharedParametersFilename = "T:\\Проектный институт\\Отдел стандартизации BIM и RD\\BIM-Ресурсы\\2-Стандарты\\01 - ФОП\\!Архив\\ФОП2018_V4_ADSK.txt"
    except Exception:
        print 'По стандартному пути не найден файл общих параметров'

paraNames = ['Ш.№Изм1', 'Ш.№Изм2', 'Ш.№Изм3', 'Ш.№Изм4',
             'Ш.КолУч1', 'Ш.КолУч2', 'Ш.КолУч3', 'Ш.КолУч4',
             'Ш. Изм\№Док 1', 'Ш. Изм\№Док 2', 'Ш. Изм\№Док 3', 'Ш. Изм\№Док 4',
             'Ш.ДатаИзм1', 'Ш.ДатаИзм2', 'Ш.ДатаИзм3', 'Ш.ДатаИзм4',
             'Ш.Значение.Изм/Зам№1', 'Ш.Значение.Изм/Зам№2', 'Ш.Значение.Изм/Зам№3', 'Ш.Значение.Изм/Зам№4']

catSheets = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Sheets)

cats = [catSheets]



#проверка на наличие нужных параметров
map = doc.ParameterBindings
it = map.ForwardIterator()
while it.MoveNext():
    newProjectParameterData = it.Key.Name
    if str(newProjectParameterData) in paraNames:
        paraNames.remove(str(newProjectParameterData))

uiDoc = __revit__.ActiveUIDocument
sel = uiDoc.Selection

catSet = doc.Application.Create.NewCategorySet()


for cat in cats:
    catSet.Insert(cat)


with revit.Transaction("Добавление параметров"):
    if len(paraNames) > 0:
        addedNames = []
        for name in paraNames:
            for dG in spFile.Groups:
                group = "12_Штамп_ZADSK"
                if str(dG.Name) == group:
                    myDefinitions = dG.Definitions
                    eDef = myDefinitions.get_Item(name)

                    newIB = doc.Application.Create.NewInstanceBinding(catSet)
                    doc.ParameterBindings.Insert(eDef, newIB)
