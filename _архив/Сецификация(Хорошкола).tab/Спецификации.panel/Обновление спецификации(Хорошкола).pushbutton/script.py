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


def new_code(collection):
    for element in collection:
        if element.LookupParameter('О_Шифр комплекта РД'):
            element.LookupParameter('О_Шифр комплекта РД').Set('01Д/2020-СПЦ09-0201-ОВК01.07')

def duct_thickness(collection):
    SelectedLink  = __revit__.ActiveUIDocument.Document

    try:
        for element in collection:
            if element.LookupParameter('Диаметр'):
                Size = element.LookupParameter('Диаметр').AsValueString()
                Size = float(Size)
        
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
                    
            if element.LookupParameter('Тип изоляции').AsString() == 'ALU1 WIRED MAT 105 (25мм EI60)':
                if thickness == '0.5' or thickness == '0.6' or thickness == '0.7':
                    thickness = '0.8'
            ElemTypeId = element.GetTypeId()
            ElemType = SelectedLink.GetElement(ElemTypeId)    
            if ElemType.get_Parameter(Guid('e6e0f5cd-3e26-485b-9342-23882b20eb43')).AsString() == 'Воздуховод из нержавеющей стали':
                if thickness == '0.5' or thickness == '0.6' or thickness == '0.7':
                    thickness = '0.8'
                    
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
                    New_Name = O_Name + ' толщиной ' + duct_thickness + ' мм ' + Size
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
        


        

#этот блок для элементов с длиной или площадью(учесть что в единицах измерения проекта должны стоять м/м2, а то в спеку уйдут миллиметры
def add_spec_param(collection, position):
    SelectedLink  = __revit__.ActiveUIDocument.Document
    k1 = 1.0
    k2 = 1.0
    for element in collection:

            ElemTypeId = element.GetTypeId()
            ElemType = SelectedLink.GetElement(ElemTypeId)
            O_Izm = ElemType.get_Parameter(Guid('4289cb19-9517-45de-9c02-5a74ebf5c86d ')).AsString()

            if O_Izm == 'м.п.':
                param = 'Длина'
            else:
                param = 'Площадь'
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
                if param == 'Длина':
                    target = (float(Length)/1000)*k1 
                else:
                    Length = Length.replace(",",".")
                    target = float(Length)*k2
                    
                Spec_Length.Set(target)
            common_param(element)

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
        
def hor(collection):
    SelectedLink  = __revit__.ActiveUIDocument.Document

    for element in collection:
        ElemTypeId = element.GetTypeId()
        ElemType = SelectedLink.GetElement(ElemTypeId)
        O_Name = ElemType.get_Parameter(Guid('e6e0f5cd-3e26-485b-9342-23882b20eb43')).AsString()
        IOS_Name = element.LookupParameter('ИОС_Наименование').AsString()
        Spec_Name = element.LookupParameter('ИОС_Наименование')
        
        if 'Врезка' not in O_Name:
                
            if 'Тройник круглого сечения' in O_Name or 'Переход круглого сечения ' in O_Name or 'Отвод круглого сечения' in O_Name:
                if 'Тройник круглого сечения' in O_Name:
                    size_1 = int(element.LookupParameter('R2').AsValueString())*2
                    size_2 = int(element.LookupParameter('R0').AsValueString())*2
                    
                    size = max(size_1, size_2)
                    
                    
                if 'Переход круглого сечения ' in O_Name:
                    size_1 = int(element.LookupParameter('R1').AsValueString())*2
                    size_2 = int(element.LookupParameter('R0').AsValueString())*2
                    
                    size = max(size_1, size_2)
                    
                if 'Отвод круглого сечения' in O_Name:
                    size = int(element.LookupParameter('Радиус воздуховода').AsValueString())*2
                    
                if size < 201:
                    thickness = '0.5'
                elif size < 451:
                    thickness = '0.6'
                elif size < 801:
                    thickness = '0.7'
                elif size < 1251:
                    thickness = '1.0'
                elif size < 1601:
                    thickness = '1.2'
                else:
                    thickness = '1.4'
                
            else:
                if 'Тройник прямоугольного сечения' in IOS_Name:
                    size_1 = int(element.LookupParameter('Ширина воздуховода 1').AsValueString())
                    size_2 = int(element.LookupParameter('Ширина воздуховода 3').AsValueString())
                    size_3 = int(element.LookupParameter('Высота воздуховода').AsValueString())
                    
                    size = max(size_1, size_2, size_3)

                if 'Переход прямоугольного сечения' in O_Name:
                    size_1 = int(element.LookupParameter('Ширина1').AsValueString())
                    size_2 = int(element.LookupParameter('Ширина2').AsValueString())
                    size_3 = int(element.LookupParameter('Высота1').AsValueString())
                    size_4 = int(element.LookupParameter('Высота2').AsValueString())
                    size = max(size_1, size_2, size_3, size_4)
                    
                if 'Переход  с прямоугольного на круглое сечение' in O_Name:
                    size_1 = int(element.LookupParameter('H0').AsValueString())
                    size_2 = int(element.LookupParameter('W0').AsValueString())
                    
                    size = max(size_1, size_2)
                    
                if 'Отвод прямоугольного сечения' in O_Name:
                    if element.LookupParameter('Ширина воздуховода'):
                        size_1 = int(element.LookupParameter('Ширина воздуховода').AsValueString())
                    if element.LookupParameter('Ширина'):
                        size_1 = int(element.LookupParameter('Ширина').AsValueString())
                    if element.LookupParameter('Высота воздуховода'):
                        size_2 = int(element.LookupParameter('Высота воздуховода').AsValueString())
                    if element.LookupParameter('Высота'):
                        size_2 = int(element.LookupParameter('Высота').AsValueString())
                    
                    size = max(size_1, size_2)
                    
                if size < 251:
                    thickness = '0.5'
                elif size < 1001:
                    thickness = '0.7'
                elif size < 2001:
                    thickness = '0.9'
                else:
                    thickness = '1.4'
                    
                    
                    
            if element.LookupParameter('Тип изоляции').AsString() == 'ALU1 WIRED MAT 105 (25мм EI60)':
                if thickness == '0.5' or thickness == '0.6' or thickness == '0.7':
                    thickness = '0.8'
            ElemTypeId = element.GetTypeId()
            ElemType = SelectedLink.GetElement(ElemTypeId)    
            if 'нержавеющей' in O_Name:
                if thickness == '0.5' or thickness == '0.6' or thickness == '0.7':
                    thickness = '0.8'
                    
            Mark_Name = IOS_Name + " толщиной " + thickness + " мм"
            
            if 'Отвод' in IOS_Name:
                angle = element.LookupParameter('Угол').AsValueString()
                angle = angle[:-4]
                angle = int(angle)
                if angle < 31:
                    angle = '30'
                elif angle < 46:
                    angle= '45'
                elif angle < 61:
                    angle= '60'                
                else:
                    angle = '90'
                    
                Mark_Name = IOS_Name + " толщиной " + thickness + " мм" + " под углом " + angle + "°"    

            Spec_Name.Set(Mark_Name)

            
            
                    
            
        
add_item_spec_param(colEquipment, '1.Оборудование')
add_item_spec_param(colAccessory, '2. Арматура')
add_item_spec_param(colTerminals, '3. Воздухораспределители')
add_spec_param(colCurves, '4. Воздуховоды')
add_spec_param(colFlexCurves, '4. Гибкие воздуховоды')
add_item_spec_param(colFittings, '5. Фасонные детали воздуховодов')
add_spec_param(colPipeCurves, '4. Трубопроводы')
add_spec_param(colFlexPipeCurves, '4. Гибкие трубопроводы')
add_item_spec_param(colPipeAccessory, '2. Трубопроводная арматура')
add_item_spec_param(colPipeFittings, '5. Фасонные детали трубопроводов')
add_spec_param(colPipeInsulations, '6. Материалы трубопроводной изоляции')
add_spec_param(colInsulations, '6. Материалы изоляции воздуховодов')



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

new_code(colEquipment)
new_code(colAccessory)
new_code(colTerminals)
new_code(colCurves)
new_code(colFlexCurves)
new_code(colFittings)
new_code(colPipeCurves)
new_code(colFlexPipeCurves)
new_code(colPipeAccessory)
new_code(colPipeFittings)
new_code(colPipeInsulations)
new_code(colInsulations)


hor(colFittings)


t.Commit()

#WriteLog.SetLogFile("Распределение по рабочим наборам", doc)