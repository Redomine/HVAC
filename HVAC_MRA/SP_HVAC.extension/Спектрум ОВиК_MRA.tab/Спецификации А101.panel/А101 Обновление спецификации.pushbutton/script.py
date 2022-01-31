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
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from rpw.ui.forms import SelectFromList
from System import Guid

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView

#Переменные для расчета
stainless_types = ['Нержавеющая сталь']
EI_insulation_types = ['EI60']
length_reserve = 1.0 #запас длинны воздуховодов
area_reserve = 1.0 #запас длинны воздуховодов

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

def duct_thickness(element):
    mode = ''
    if str(element.Category.Name) == 'Соединительные детали воздуховодов':
        a = getConnectors(element)
        try:
            SizeA = a[0].Width*0.3048
            SizeB = a[1].Width*0.3048
            mode = 'W'
        except:
            Size = a[0].Radius*2*0.3048
            mode = 'R'
        
    if element.LookupParameter('Диаметр') or mode == 'R':
        
        if mode != 'R':
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
            
    if element.LookupParameter('Ширина') or mode == 'W':
        if mode != 'W':
            SizeA = float(element.LookupParameter('Ширина').AsValueString())
            SizeB = float(element.LookupParameter('Высота').AsValueString())
        if SizeA > SizeB:
            SizeC = SizeA
        else:
            SizeC = SizeB
        
        if SizeC < 251:
            thickness = '0.5'
        elif SizeC < 1001:
            thickness = '0.7'
        elif SizeC < 2001:
            thickness = '0.9'
        else:
            thickness = '1.4'
    return thickness
 



def make_new_name(collection):
    SelectedLink  = __revit__.ActiveUIDocument.Document
    for element in collection:
        
        Spec_Name = element.LookupParameter('ФОП_ВИС_Наименование комбинированное')
        if element.LookupParameter('ADSK_Наименование'):
            ADSK_Name = element.LookupParameter('ADSK_Наименование').AsString()
        else:
            ElemTypeId = element.GetTypeId()
            ElemType = SelectedLink.GetElement(ElemTypeId)
            ADSK_Name = ElemType.get_Parameter(Guid('e6e0f5cd-3e26-485b-9342-23882b20eb43')).AsString()
            
        if ADSK_Name == None:
            print 'Для категории не заполнен параметр ADSK_Наименование'
            print element.LookupParameter('ФОП_ВИС_Группирование').AsString()
            continue
        
            
        New_Name = ADSK_Name
        
        if element.LookupParameter('ФОП_ВИС_Группирование').AsString() == '4. Трубопроводы':
            external_size = element.LookupParameter('Внешний диаметр').AsDouble() * 304.8
            internal_size = element.LookupParameter('Внутренний диаметр').AsDouble() * 304.8
            pipe_thickness = external_size - internal_size 
            

            Dy = element.LookupParameter('Диаметр').AsDouble() * 304.8

            New_Name = ADSK_Name + ' ' + 'Ду='+ str(Dy) + ' (Днар. х т.с. ' + str(external_size) + 'x' + str(pipe_thickness) + ')'
            
        if element.LookupParameter('ФОП_ВИС_Группирование').AsString() == '4. Воздуховоды':
            thickness = duct_thickness(element)
            New_Name = ADSK_Name + ' толщиной ' + thickness + ' мм'
            
        if element.LookupParameter('ФОП_ВИС_Группирование').AsString() == '6. Материалы трубопроводной изоляции':
            ADSK_Izm = ElemType.get_Parameter(Guid('4289cb19-9517-45de-9c02-5a74ebf5c86d')).AsString()
            if ADSK_Izm == 'м.п.' or ADSK_Izm == 'м.' or ADSK_Izm == 'мп' or ADSK_Izm == 'м' or ADSK_Izm == 'м.п':
                L = element.LookupParameter('Длина').AsDouble() * 304.8
                S = element.LookupParameter('Площадь').AsDouble() * 0.092903

                pipe = doc.GetElement(element.HostElementId)
                if pipe.LookupParameter('Внешний диаметр') != None:
                    d = pipe.LookupParameter('Внешний диаметр').AsDouble() * 304.8
                    New_Name = ADSK_Name + ' внутренним диаметром Ø' + str(d)
                    
        if element.LookupParameter('ФОП_ВИС_Группирование').AsString() == '5. Фасонные детали воздуховодов':
            
            New_Name = ADSK_Name + ' ' + element.LookupParameter('Размер').AsString()
            if str(element.MEPModel.PartType) == 'Elbow':
                thickness = duct_thickness(element)
                New_Name = ADSK_Name + ' толщиной ' + thickness + ' мм'

                    
                     
            
        Spec_Name.Set(New_Name)

def getConnectors(element):
    connectors = []
    try:
        a = element.ConnectorManager.Connectors.ForwardIterator()
        while a.MoveNext():
            connectors.append(a.Current)
    except:
        try:
            a = element.MEPModel.ConnectorManager.Connectors.ForwardIterator()
            while a.MoveNext():
                connectors.append(a.Current)
        except:
            a = element.MEPSystem.ConnectorManager.Connectors.ForwardIterator()
            while a.MoveNext():
                connectors.append(a.Current)
    return connectors




    
#в этом блоке получаем размеры воздуховодов и труб для наименования в спеке    
def getElementSize(element):
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

#этот блок для элементов с длиной или площадью(учесть что в единицах измерения проекта должны стоять милимметры для длины и м2 для площади)
def getCapacityParam(collection, position, length_reserve, area_reserve):
    SelectedLink  = __revit__.ActiveUIDocument.Document
    for element in collection:
            if element.LookupParameter('ФОП_ВИС_Группирование'):
                Pos = element.LookupParameter('ФОП_ВИС_Группирование')
                Pos.Set(position)
             
            ElemTypeId = element.GetTypeId()
            ElemType = SelectedLink.GetElement(ElemTypeId)
            
            ADSK_Izm = ElemType.get_Parameter(Guid('4289cb19-9517-45de-9c02-5a74ebf5c86d')).AsString()
            if ADSK_Izm == 'м.п.' or ADSK_Izm == 'м.' or ADSK_Izm == 'мп' or ADSK_Izm == 'м' or ADSK_Izm == 'м.п':
                param = 'Длина'
            elif ADSK_Izm == 'шт' or ADSK_Izm == 'шт.':
                amount = element.LookupParameter('ФОП_ВИС_Число')
                amount.Set(1)
                continue
            else:
                param = 'Площадь'

            if element.LookupParameter(param):
                if param == 'Длина':
                    CapacityParam = ((element.LookupParameter(param).AsDouble() * 304.8)/1000) * length_reserve 
                else:
                    CapacityParam = (element.LookupParameter(param).AsDouble() * 0.092903) * area_reserve
                    
                if element.LookupParameter('ФОП_ВИС_Число'):
                    if CapacityParam == None: continue
                    Spec_Length = element.LookupParameter('ФОП_ВИС_Число')
                    Spec_Length.Set(CapacityParam)

#этот блок для элементов которые идут поштучно 
def getNumericalParam(collection, position):
    for element in collection:
        try:
            if element.Location:
                if element.LookupParameter('ФОП_ВИС_Группирование'):
                    Pos = element.LookupParameter('ФОП_ВИС_Группирование')
                    Pos.Set(position)
                
                if element.LookupParameter('ФОП_ВИС_Число'):
                    amount = element.LookupParameter('ФОП_ВИС_Число')
                    amount.Set(1)
        except Exception:
            pass
        
getNumericalParam(colEquipment, '1.Оборудование')
getNumericalParam(colAccessory, '2. Арматура')
getNumericalParam(colTerminals, '3. Воздухораспределители')
getNumericalParam(colPipeAccessory, '2. Трубопроводная арматура')
getNumericalParam(colPipeFittings, '5. Фасонные детали трубопроводов')
getNumericalParam(colFittings, '5. Фасонные детали воздуховодов')

getCapacityParam(colCurves, '4. Воздуховоды', length_reserve, area_reserve)
getCapacityParam(colFlexCurves, '4. Гибкие воздуховоды', length_reserve, area_reserve)
getCapacityParam(colPipeCurves, '4. Трубопроводы', length_reserve, area_reserve)
getCapacityParam(colFlexPipeCurves, '4. Гибкие трубопроводы', length_reserve, area_reserve)
getCapacityParam(colPipeInsulations, '6. Материалы трубопроводной изоляции', length_reserve, area_reserve)
getCapacityParam(colInsulations, '6. Материалы изоляции воздуховодов', length_reserve, area_reserve)



make_new_name(colEquipment)
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

# duct_thickness(colCurves)


t.Commit()
