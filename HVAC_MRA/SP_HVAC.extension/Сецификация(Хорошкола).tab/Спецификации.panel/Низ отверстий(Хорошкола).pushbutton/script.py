#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Низ отверстий'
__doc__ = "-"


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
    
colModels = make_col(BuiltInCategory.OST_GenericModel)    


t = Transaction(doc, 'Обновление общей спеки')

t.Start()




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

    
    
for element in colModels:
    try:
        if element.LookupParameter('Уровень'):
            base = element.LookupParameter('Уровень').AsValueString()
            if base == 'Уровень кровли на отм. +13.650':
                base = 13650
            if base == 'Уровень кровли техн. этажа на отм. +17.400':
                base = 17400
            if base == 'Уровень этажа на отм. -4.500':
                base = -4500
            if base == 'Уровень этажа на отм. 0.000':
                base = 0
            if base == 'Уровень этажа на отм. +9.000':
                base = 9000
            if base == 'Уровень этажа на отм. +4.500':
                base = 4500
                
            removal = element.LookupParameter('Смещение').AsValueString()
            height = element.LookupParameter('О_Отверстие (Высота)').AsValueString()
            new_base = base - float(height)/2 + float(removal)
            
            target = element.LookupParameter('КР_Размер_Отметка расположения')
            target.Set(new_base/304.8)
    except Exception:
        pass

        
            



t.Commit()

