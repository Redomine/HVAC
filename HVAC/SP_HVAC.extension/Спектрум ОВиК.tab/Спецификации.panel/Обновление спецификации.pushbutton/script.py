#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Обновление общей \n спецификации'
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



def common_param(element):
    Size = 0
    if element.Name == 'СП_Медная':
        if element.LookupParameter('Внешний диаметр'):
            Size = element.LookupParameter('Внешний диаметр').AsValueString()
            print Size
        
    elif element.LookupParameter('Диаметр'):
        Size = element.LookupParameter('Диаметр').AsValueString()
        if Size == None:
            if element.LookupParameter('Размер'):
                Size = element.LookupParameter('Размер').AsString()
    elif element.LookupParameter('Размер'):
        Size = element.LookupParameter('Размер').AsString()

    if element.LookupParameter('ИОС_Размер'):
        Spec_Size = element.LookupParameter('ИОС_Размер')
        Spec_Size.Set(Size)
        
    try:
        if element.LookupParameter('ИОС_Наименование'):
            if element.LookupParameter('О_Наименование') != None:
                Spec_Name = element.LookupParameter('ИОС_Наименование')
                Old_Name = element.LookupParameter('О_Наименование').AsString()
                New_Name = Old_Name + ' ' + Size

                Spec_Name.Set(New_Name)
    except Exception:
        pass

        

#этот блок для элементов с длиной или площадью(учесть что в единицах измерения проекта должны стоять м/м2, а то в спеку уйдут миллиметры
def add_spec_param(collection, param, position):
    k1 = 1.15
    k2 = 1.20
    for element in collection:
        
        try:
                if element.LookupParameter('ИОС_Позиция в спецификации'):
                    Pos = element.LookupParameter('ИОС_Позиция в спецификации')
                    Pos.Set(position)    
                if element.LookupParameter(param):
                    Length = element.LookupParameter(param).AsValueString()
                    Length = Length.split(' ')
                    Length = Length[0] 
                if element.LookupParameter('ИОС_Количество'):
                    
                    if Length == None: continue
                    Spec_Length = element.LookupParameter('ИОС_Количество')
                    if param != 'Площадь':
                        target = (float(Length)/1000)*k1
  
                        
                    else: target = float(Length)*k2
                    Spec_Length.Set(target)
                common_param(element)
        except Exception:
            pass

#этот блок для элементов которые идут поштучно 
def add_item_spec_param(collection, position):
    for element in collection:
        try:
            if element.Name == 'СП_Вспомогательное_Спецификация':
                continue
            if element.Location:
                if element.LookupParameter('ИОС_Позиция в спецификации'):
                    Pos = element.LookupParameter('ИОС_Позиция в спецификации')
                    Pos.Set(position)
                
                if element.LookupParameter('ИОС_Количество'):
                    Spec_Length = element.LookupParameter('ИОС_Количество')
                    Spec_Length.Set(1)
                common_param(element)
        except Exception:
            continue
        
add_item_spec_param(colEquipment, '1')
add_item_spec_param(colAccessory, '2')
add_item_spec_param(colTerminals, '3')
add_spec_param(colCurves, 'Длина', '4')
add_spec_param(colFlexCurves, 'Длина', '4')
add_item_spec_param(colFittings, '5')
add_spec_param(colPipeCurves, 'Длина', '4')
add_spec_param(colFlexPipeCurves, 'Длина', '4')
add_item_spec_param(colPipeAccessory, '2')
add_item_spec_param(colPipeFittings, '5')
add_spec_param(colPipeInsulations, 'Длина', '6')
add_spec_param(colInsulations, 'Площадь', '6')


    
t.Commit()

#WriteLog.SetLogFile("Распределение по рабочим наборам", doc)