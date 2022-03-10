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
    print 'Подгружен не тот файл общих параметров, переключитесь на ФОП_v1'
    sys.exit()

paraNames = ['ФОП_ВИС_Группирование', 'ФОП_ВИС_Единица измерения', 'ФОП_ВИС_Масса', 'ФОП_ВИС_Минимальная толщина воздуховода',
             'ФОП_ВИС_Наименование комбинированное', 'ФОП_ВИС_Число']

catFittings = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctFitting)
catPipeFittings = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeFitting)
catPipeCurves = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeCurves)
catCurves = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctCurves)
catFlexCurves = doc.Settings.Categories.get_Item(BuiltInCategory.OST_FlexDuctCurves)
catFlexPipeCurves = doc.Settings.Categories.get_Item(BuiltInCategory.OST_FlexPipeCurves)
catTerminals = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctTerminal)
catAccessory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctAccessory)
catPipeAccessory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeAccessory)
catEquipment = doc.Settings.Categories.get_Item(BuiltInCategory.OST_MechanicalEquipment)
catInsulations = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctInsulations)
catPipeInsulations = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeInsulations)
catPlumbingFixtures = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PlumbingFixtures)

cats = [catFittings, catPipeFittings, catPipeCurves, catCurves, catFlexCurves, catFlexPipeCurves, catTerminals, catAccessory,
        catPipeAccessory, catEquipment, catInsulations, catPipeInsulations, catPlumbingFixtures]


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
                group = "03_ВИС"
                if str(dG.Name) == group:
                    myDefinitions = dG.Definitions
                    eDef = myDefinitions.get_Item(name)
                    newIB = doc.Application.Create.NewInstanceBinding(catSet)
                    doc.ParameterBindings.Insert(eDef, newIB)
