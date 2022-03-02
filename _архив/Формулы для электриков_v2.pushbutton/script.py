#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Добавление формул'
__doc__ = "Добавление формул"


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

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView


def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col 
    
# create a filtered element collector set to Category OST_Mass and Class FamilySymbol
collector = FilteredElementCollector(doc)
collector.OfCategory(BuiltInCategory.OST_GenericModel)
collector.OfClass(FamilySymbol)
famtypeitr = collector.GetElementIdIterator()
famtypeitr.Reset()

colModel = make_col(BuiltInCategory.OST_GenericModel)

is_temporary_in = False
for element in famtypeitr:
    famtypeID = element
    famsymb = doc.GetElement(famtypeID)

    if famsymb.Family.Name == '_Преднастроенный коннектор':
        temporary = famsymb
        is_temporary_in = True

if is_temporary_in == False:
    print 'Не обнаружено преднастроенное семейство для заданий, проверьте не менялось ли его имя или загружалось ли оно'
    sys.exit()


manager = doc.FamilyManager
with revit.Transaction("Обновление общей спеки"):
    for element in colModel:
        #вложенные семейства не удаляются так, потом переделаю на проверку вложенности и все
        try:
            if element.LookupParameter('Семейство').AsValueString() == '_Преднастроенный коннектор':
                element.LookupParameter('ADSK_Номинальная мощность')
                doc.Delete(element.Id)
        except Exception:
            pass

    loc = XYZ(0, 0, 0)
    familyInst = doc.FamilyCreate.NewFamilyInstance(loc, temporary, Structure.StructuralType.NonStructural)
