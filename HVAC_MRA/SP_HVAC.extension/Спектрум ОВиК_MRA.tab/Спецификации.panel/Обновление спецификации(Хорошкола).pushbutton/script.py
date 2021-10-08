#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Обновление общей \n спецификации(Хорошкола)'
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

def duct_thickness(collection):

    try:
        for element in collection:
            if element.LookupParameter('Диаметр'):
                Size = element.LookupParameter('Диаметр').AsValueString()
                Size = float(Size)
                if Size < 501:
                    thickness = '0.5'
                elif Size < 901:
                    thickness = '0.7'
                elif Size < 1251:
                    thickness = '1'
                else:
                    thickness = '1.2'
                    
            if element.LookupParameter('Ширина'):
                SizeA = element.LookupParameter('Ширина').AsValueString()
                SizeB = element.LookupParameter('Высота').AsValueString()
                if SizeA > SizeB:
                    SizeC = SizeA
                else:
                    SizeC = SizeB
                if SizeC < 301:
                    thickness = '0.5'
                elif SizeC < 1001:
                    thickness = '0.7'
                elif SizeC < 1251:
                    thickness = '1'
                else:
                    thickness = '1.2'
                    
            if element.LookupParameter('ИОС_Толщина воздуховода'):
                duct_thickness = element.LookupParameter('ИОС_Толщина воздуховода')
                duct_thickness.Set(thickness)  
            
    except Exception:
        pass   


def make_new_name(collection, status, mark):
    SelectedLink  = __revit__.ActiveUIDocument.Document
    for element in collection:
        try:
            if status != '+':       
                Spec_Name = element.LookupParameter('ИОС_Наименование')
                if element.LookupParameter('О_Наименование'):
                    O_Name = element.LookupParameter('О_Наименование').AsString()
                else:
                    ElemTypeId = element.GetTypeId()
                    ElemType = SelectedLink.GetElement(ElemTypeId)
                    O_Name = ElemType.get_Parameter(Guid('e6e0f5cd-3e26-485b-9342-23882b20eb43')).AsString()
                New_Name = O_Name  
                Spec_Name.Set(New_Name)
            else:
                if element.LookupParameter('ИОС_Размер'):
                    Size = element.LookupParameter('ИОС_Размер').AsString()
                Spec_Name = element.LookupParameter('ИОС_Наименование')
                if element.LookupParameter('О_Наименование'):
                    Old_Name = element.LookupParameter('О_Наименование').AsString()
                    New_Name = Old_Name + ' ' + Size  
                else:
                    ElemTypeId = element.GetTypeId()
                    ElemType = SelectedLink.GetElement(ElemTypeId)
                    O_Name = ElemType.get_Parameter(Guid('e6e0f5cd-3e26-485b-9342-23882b20eb43')).AsString()
                    New_Name = O_Name + ' ' + Size
                
                if element.LookupParameter('ИОС_Толщина воздуховода'):
                    duct_thickness = element.LookupParameter('ИОС_Толщина воздуховода').AsString()
                    New_Name = New_Name + ' δ=' + duct_thickness + 'мм'
                Spec_Name.Set(New_Name)
                
            if element.LookupParameter('О_Марка'):
                O_Mark = element.LookupParameter('О_Марка').AsString()
            else:
                ElemTypeId = element.GetTypeId()
                ElemType = SelectedLink.GetElement(ElemTypeId)
                O_Mark = ElemType.get_Parameter(Guid('2204049c-d557-4dfc-8d70-13f19715e46d')).AsString()
            
            if O_Mark != None and O_Mark != "" and mark == "+" and O_Mark != "-":
                Mark_Name = element.LookupParameter('ИОС_Наименование').AsString() + " " + "(" + O_Mark +")"
                Spec_Name.Set(Mark_Name)
        except Exception:
            pass    

    
    
def common_param(element):
    Size = ''
    if element.Name == 'СП_Медная':
        if element.LookupParameter('Внешний диаметр'):
            Size = "Ø" + element.LookupParameter('Внешний диаметр').AsValueString()
            Spec_Size = element.LookupParameter('ИОС_Размер')
            Spec_Size.Set(Size)
           
    elif element.LookupParameter('Внешний диаметр'):
        outer_size = element.LookupParameter('Внешний диаметр').AsValueString()
        interior_size = element.LookupParameter('Внутренний диаметр').AsValueString()
        thickness = (float(outer_size) - float(interior_size))/2
        Size = "Ø" + outer_size + "x" + str(thickness)
        Spec_Size = element.LookupParameter('ИОС_Размер')
        Spec_Size.Set(Size)
    elif element.LookupParameter('Размер'):
        Size = element.LookupParameter('Размер').AsString()
        Spec_Size = element.LookupParameter('ИОС_Размер')
        Spec_Size.Set(Size)
        
    elif element.LookupParameter('Размер трубы'):
        Size = element.LookupParameter('Размер трубы').AsString()
        Spec_Size = element.LookupParameter('ИОС_Размер')
        Spec_Size.Set(Size)
    
    #if element.LookupParameter('ИОС_Размер'):
    #    Spec_Size = element.LookupParameter('ИОС_Размер')
    #    Spec_Size.Set(Size)
    

        


        

#этот блок для элементов с длиной или площадью(учесть что в единицах измерения проекта должны стоять м/м2, а то в спеку уйдут миллиметры
def add_spec_param(collection, param, position):
    k1 = 1.0
    k2 = 1.0
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
        
add_item_spec_param(colEquipment, '1.Оборудование')
add_item_spec_param(colAccessory, '2. Арматура')
add_item_spec_param(colTerminals, '3. Воздухораспределители')
add_spec_param(colCurves, 'Длина', '4. Воздуховоды')
add_spec_param(colFlexCurves, 'Длина', '4. Гибкие воздуховоды')
add_item_spec_param(colFittings, '5. Фасонные детали воздуховодов')
add_spec_param(colPipeCurves, 'Длина', '4. Трубопроводы')
add_spec_param(colFlexPipeCurves, 'Длина', '4. Гибкие трубопроводы')
add_item_spec_param(colPipeAccessory, '2. Трубопроводная арматура')
add_item_spec_param(colPipeFittings, '5. Фасонные детали трубопроводов')
add_spec_param(colPipeInsulations, 'Длина', '6. Материалы трубопроводной изоляции')
add_spec_param(colInsulations, 'Площадь', '6. Материалы изоляции воздуховодов')



make_new_name(colEquipment, '-', '+')
make_new_name(colAccessory, '-', '+')
make_new_name(colTerminals, '-', '+')
make_new_name(colCurves, '+', '-')
make_new_name(colFlexCurves, '+', '-')
make_new_name(colFittings, '+', '-')
make_new_name(colPipeCurves, '+', '-')
make_new_name(colFlexPipeCurves, '+', '-')
make_new_name(colPipeAccessory, '-', '+')
make_new_name(colPipeFittings, '+', '-')
make_new_name(colPipeInsulations, '+', '-')
make_new_name(colInsulations, '-', '-')

duct_thickness(colCurves)


t.Commit()

#WriteLog.SetLogFile("Распределение по рабочим наборам", doc)