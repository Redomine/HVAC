#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'ИОС_наименование \n системы(Хорошкола)'
__doc__ = "Обновляет имя"


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
    
colFittings = make_col(BuiltInCategory.OST_DuctFitting)    
colPipeFittings = make_col(BuiltInCategory.OST_PipeFitting)
colPipeCurves = make_col(BuiltInCategory.OST_PipeCurves)
colCurves = make_col(BuiltInCategory.OST_DuctCurves)
colFlexCurves = make_col(BuiltInCategory.OST_FlexDuctCurves)
colFlexPipeCurves = make_col(BuiltInCategory.OST_FlexPipeCurves)
colTerminals = make_col(BuiltInCategory.OST_DuctTerminal)
colAccessory = make_col(BuiltInCategory.OST_DuctAccessory)
colPipeAccessory = make_col(BuiltInCategory.OST_PipeAccessory)
colEquipment = make_col(BuiltInCategory.OST_MechanicalEquipment)
colInsulations = make_col(BuiltInCategory.OST_DuctInsulations)
colPipeInsulations = make_col(BuiltInCategory.OST_PipeInsulations)

t = Transaction(doc, 'Обновление общей спеки')

t.Start()




def make_new_name(collection):
    for element in collection:
        try:
            if element.LookupParameter('Имя системы'):
                sys_name = element.LookupParameter('Имя системы').AsString()
                if 'СП' in sys_name:
                    pass
                else:
                    sys_name = sys_name[:-1]
                    i_name = element.LookupParameter('ИОС_Наименование системы')
                    i_name.Set(sys_name)
        except Exception:
            pass
        
def make_new_name_eq(collection):
    for element in collection:
        try:
            if element.LookupParameter('Имя системы'):
                sys_name = element.LookupParameter('Имя системы').AsString()
                if 'СП' in sys_name:
                    pass
                else:
                    sys_name = sys_name.split(',')
                    sys_name = sys_name[1]
                    sys_name = sys_name[:-1]
                    i_name = element.LookupParameter('ИОС_Наименование системы')
                    i_name.Set(sys_name)
        except Exception:
            pass  


make_new_name_eq(colEquipment)
make_new_name(colAccessory)
make_new_name(colTerminals)
make_new_name(colCurves)
make_new_name(colFlexCurves)
make_new_name(colFittings)
make_new_name(colPipeCurves)
make_new_name(colFlexPipeCurves)
make_new_name(colPipeAccessory)
make_new_name(colPipeFittings)
make_new_name(colPipeInsulations)
make_new_name(colInsulations)

t.Commit()

