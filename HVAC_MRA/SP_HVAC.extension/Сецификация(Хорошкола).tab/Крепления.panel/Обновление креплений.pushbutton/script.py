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


t = Transaction(doc, 'Обновление общей спеки')

t.Start()

def duct_thickness(collection):

    try:
        for element in collection:
            if element.LookupParameter('Диаметр'):
                Size = element.LookupParameter('Диаметр').AsValueString()
                Size = float(Size)
                
                if element.LookupParameter('ИОС_Толщина воздуховода').AsString() == '1.0':
                    continue
                if Size < 201:
                    thickness = '0.5'
                elif Size < 451:
                    thickness = '0.6'
                elif Size < 801:
                    thickness = '0.7'
                elif Size < 1251:
                    thickness = '1.0'
                elif Size < 1601:
                    thickness = '1.2'
                else:
                    thickness = '1.4'
                    
            if element.LookupParameter('Ширина'):
                SizeA = float(element.LookupParameter('Ширина').AsValueString())
                SizeB = float(element.LookupParameter('Высота').AsValueString())
                if SizeA > SizeB:
                    SizeC = SizeA
                else:
                    SizeC = SizeB
                    
                if element.LookupParameter('ИОС_Толщина воздуховода').AsString() == '1.0':
                    continue
                
                if SizeC < 251:
                    thickness = '0.5'
                elif SizeC < 1001:
                    thickness = '0.7'
                elif SizeC < 2001:
                    thickness = '0.9'
                else:
                    thickness = '1.4'
                
                    
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
                    
                    if element.LookupParameter('ИОС_Позиция в спецификации').AsString() == '4. Трубопроводы':
                        Dy = element.LookupParameter('Размер').AsString()
                        Dy = Dy[1:]
                        New_Name = O_Name + ' ' + 'Ду='+ Dy + ' (Днар. х т.с. ' + Size + ')'
                    else:
                        New_Name = O_Name + ' ' + Size

                if element.LookupParameter('ИОС_Толщина воздуховода'):
                    duct_thickness = element.LookupParameter('ИОС_Толщина воздуховода').AsString()
                    New_Name = New_Name + ' толщиной ' + duct_thickness + ' мм'
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
              
    if element.LookupParameter('Внешний диаметр'):
        outer_size = element.LookupParameter('Внешний диаметр').AsDouble() * 304.8
        interior_size = element.LookupParameter('Внутренний диаметр').AsDouble() * 304.8
        thickness = (float(outer_size) - float(interior_size))/2
        outer_size = str(outer_size)
        
        a = outer_size.split('.') #убираем 0 после запятой в наружном диаметре если он не имеет значения
        if a[1] == '0':
            outer_size = outer_size.replace(".0","")
        Size = "Ø" + outer_size + "x" + str(thickness)
        Spec_Size = element.LookupParameter('ИОС_Размер')
        Spec_Size.Set(Size)
    elif element.LookupParameter('Размер'):
        Size = element.LookupParameter('Размер').AsString()
        Spec_Size = element.LookupParameter('ИОС_Размер')
        Spec_Size.Set(Size)
        
    elif element.LookupParameter('Диаметр'):
        Size = element.LookupParameter('Диаметр').AsValueString()
        Spec_Size = element.LookupParameter('ИОС_Размер')
        Spec_Size.Set(Size)
            
    elif element.LookupParameter('Размер трубы'):
        Size = element.LookupParameter('Размер трубы').AsString()
        Spec_Size = element.LookupParameter('ИОС_Размер')
        Spec_Size.Set(Size)
        
        #if element.LookupParameter('ИОС_Размер'):
        #    Spec_Size = element.LookupParameter('ИОС_Размер')
        #    Spec_Size.Set(Size)
    
    if element.LookupParameter('ИОС_Позиция в спецификации').AsString() == '3. Воздухораспределители':
        Spec_Size.Set('-')
    
    if element.LookupParameter('ИОС_Позиция в спецификации').AsString() == '2. Арматура':
        Spec_Size.Set('-')
    
    if element.LookupParameter('ИОС_Позиция в спецификации').AsString() == '2. Трубопроводная арматура':
        Spec_Size.Set('-') 
        

def bracing_curves(collection):
    for element in collection:
        try:
            if element.LookupParameter('Диаметр'):
                dy = element.LookupParameter('Диаметр').AsValueString()
            if element.LookupParameter('Эквивалентный диаметр'):
                dy = element.LookupParameter('Эквивалентный диаметр').AsValueString()
                dy = float(dy)
            if element.LookupParameter('Длина'):
                long = element.LookupParameter('Длина').AsValueString()
                long = float(long)/1000
                
            if dy < 159:
                kg = 0.33
            elif dy < 314:
                kg = 0.75
            elif dy < 499:
                kg = 1.8
            elif dy < 709:
                kg = 4
            elif dy < 899:
                kg = 6.5
            else:
                kg = 8.8
            
            if element.LookupParameter('Количество креплений'):
                element.LookupParameter('Количество креплений').Set(long*kg)
            
        except Exception:
            pass
            
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
            


bracing_curves_v2(colCurves)



t.Commit()

#WriteLog.SetLogFile("Распределение по рабочим наборам", doc)