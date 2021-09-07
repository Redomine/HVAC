#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = '3.Назначить расходы'
__doc__ = "Перенос расчетного расхода между параметрами, с округлением в большую сторону"

import sys
import clr
import math
import WriteLog
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from itertools import compress
from rpw.ui.forms import Alert
from rpw.ui.forms import select_file
from rpw.ui.forms import SelectFromList
import System.Drawing
import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *


clr.AddReference('ProgressBar')
from AdnRme import ProgressForm

import WriteLog


doc = __revit__.ActiveUIDocument.Document

param_vent_names = ['ИОС_Расход воздуха приточный', 'ИОС_Расход воздуха вытяжной']

param_names = ['ИОС_Теплопотери', 'ИОС_Теплопритоки']

colHVAC_VENT_equipment = FilteredElementCollector(doc)\
                            .OfCategory(BuiltInCategory.OST_DuctTerminal)\
                            .WhereElementIsNotElementType()\
                            .ToElements()

colHVACequipment = FilteredElementCollector(doc)\
                            .OfCategory(BuiltInCategory.OST_MechanicalEquipment)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    
k = 101.940647731554 #Коэффициент для перевода в метры кубические.

k_w = 0.0929026687598116 #Коэффициент для перевода в ватты

t = Transaction(doc, 'Перенос расходов в оборудование')

t.Start()
try:
    for equipment in colHVAC_VENT_equipment:
        if equipment.Location:    
            if equipment.LookupParameter(param_vent_names[0]):
                calculated_supply_1 = equipment.LookupParameter(param_vent_names[0]).AsDouble()
            if equipment.LookupParameter(param_vent_names[1]):
                calculated_supply_2 = equipment.LookupParameter(param_vent_names[1]).AsDouble()
                
            if calculated_supply_1 > calculated_supply_2: calculated_supply = calculated_supply_1
            else: calculated_supply = calculated_supply_2
            
            calculated_supply = calculated_supply*k
            calculated_supply = math.ceil(calculated_supply)
            if calculated_supply!=0:
                if calculated_supply%10 == 5: continue 
                if calculated_supply%10 != 0:
                    if calculated_supply%10 < 5: calculated_supply+= (5 - calculated_supply%10)
                    if calculated_supply%10 > 5: calculated_supply+= (10 - calculated_supply%10)
                    continue

            if equipment.LookupParameter('Расход'):
                target_supply = equipment.LookupParameter('Расход')
                if round((target_supply.AsDouble())*k, 2) != round(calculated_supply, 2):
                    print 'Изменились расходы на оборудовании! ID:'
                    print equipment.Id
                    print 'Было:', round((target_supply.AsDouble())*k, 2), 'Стало:', round(calculated_supply, 2)
                target_supply.Set(calculated_supply/k)
except Exception:
    Alert('Ошибка при при работе с вентиляцией!', title= 'Ошибка', header = 'Проблемы с моделью')
    sys.exit()

try:
    for equipment in colHVACequipment:
        if equipment.Location:
            for param_name in param_names:
                if equipment.LookupParameter(param_name):
                    print param_name
                    calculated_supply = equipment.LookupParameter(param_name).AsDouble()
                    print calculated_supply

            
            calculated_supply = calculated_supply*k_w
            calculated_supply = math.ceil(calculated_supply)

            if calculated_supply!=0:
                if calculated_supply%10 == 5: continue 
                if calculated_supply%10 != 0:
                    if calculated_supply%10 < 5: calculated_supply+= (5 - calculated_supply%10)
                    if calculated_supply%10 > 5: calculated_supply+= (10 - calculated_supply%10)
                    continue
            
            if equipment.LookupParameter('Мощность'):
                target_supply = equipment.LookupParameter('Мощность')
                if target_supply.AsDouble() != calculated_supply/k_w:
                    print 'Изменились расходы на оборудовании! ID:'
                    print equipment.Id
                    print 'Было:', round(target_supply.AsDouble()*k_w), 'Стало:', calculated_supply
                target_supply.Set(calculated_supply/k_w)
except Exception:
    Alert('Ошибка при при работе с отоплением/кондиционированием!', title= 'Ошибка', header = 'Проблемы с моделью')
    sys.exit()
      
t.Commit()
    
WriteLog.SetLogFile("3.Назначить расходы", doc)