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
import WriteLog

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from rpw.ui.forms import CommandLink, TaskDialog
from rpw.ui.forms import select_file
from rpw.ui.forms import TextInput
from rpw.ui.forms import SelectFromList
from rpw.ui.forms import Alert
from System.Collections.Generic import List
from rpw.ui.forms import SelectFromList
import System.Drawing
import System.Windows.Forms
from System.Drawing import *
from System.Windows.Forms import *
from System import Guid
doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView



def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col 
    
colPipeCurves = make_col(BuiltInCategory.OST_PipeCurves)
colCurves = make_col(BuiltInCategory.OST_DuctCurves)
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

t = Transaction(doc, 'Обновление общей спеки')

t.Start()
            


bracing_curves_v2(colCurves)
bracing_pipes(colPipeCurves)


t.Commit()
