#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Решетки по пространствам'
__doc__ = "pass"


import clr
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




doc = __revit__.ActiveUIDocument.Document



            
colHVACequipment = FilteredElementCollector(doc)\
                            .OfCategory(BuiltInCategory.OST_DuctTerminal)\
                            .WhereElementIsNotElementType()\
                            .ToElements()

t = Transaction(doc, 'Решетки по пространствам')
t.Start()



for equipment in colHVACequipment:
    try:
        FamilyName = equipment.Symbol.Family.Name    
        phase = equipment.CreatedPhaseId #Следующие две строки - фаза. Без нее ревит не даст нам информации по элементу, просто пользуйтесь ей, чтоб получить объект.
        equipmentPhase = doc.GetElement(phase)
        ItemSpace = equipment.Space[equipmentPhase] #создаем объект простраства с которым можно работать
        SpaceName = Element.Name.__get__(ItemSpace).ToString()#проверка на адекватность пространств
        if SpaceName == 'IronPython.Runtime.Types.ReflectedProperty':
            continue
        SpaceNumber = SpatialElement.Number.__get__(ItemSpace).ToString()#забираем номер пространства для оборудования
        Space_Name = equipment.LookupParameter('ИОС_Номер пространства')
        Space_Name.Set(SpaceNumber)

        
        print SpaceNumber
    except Exception:
        pass
                            
                            
t.Commit()
    
