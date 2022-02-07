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

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView


def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col 
    


t = Transaction(doc, 'Добавление формул')

manager = doc.FamilyManager

t.Start()
set = doc.FamilyManager.Parameters

for x in set:
    if str(x.Definition.Name) == 'ADSK_Полная мощность':
        manager.SetFormula(x, "ADSK_Номинальная мощность / ADSK_Коэффициент мощности")

    if str(x.Definition.Name) == 'ADSK_Коэффициент мощности':
        manager.SetFormula(x, "if(ADSK_Номинальная мощность < 1 кВт, 0.65, if(ADSK_Номинальная мощность < 4.0001 кВт, 0.75, if(ADSK_Номинальная мощность > 4 кВт, 0.85, 0)))")
        
    if str(x.Definition.Name) == 'ADSK_Количество фаз':
        manager.SetFormula(x, "if(ADSK_Напряжение < 250 В, 1, 3)")
        

t.Commit()
