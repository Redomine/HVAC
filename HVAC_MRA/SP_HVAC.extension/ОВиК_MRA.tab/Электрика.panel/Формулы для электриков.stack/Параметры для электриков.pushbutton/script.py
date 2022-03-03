#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Добавление параметров'
__doc__ = "Добавление параметров в семейство для выдачи заданий электрикам"


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
from pyrevit import revit
from Autodesk.Revit.DB.Electrical import *


doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView



def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col 
    
connectorCol = make_col(BuiltInCategory.OST_ConnectorElem)
loadsCol = make_col(BuiltInCategory.OST_ElectricalLoadClassifications)

t = Transaction(doc, 'Добавление формул')

manager = doc.FamilyManager

def associate(param, famparam):
    manager.AssociateElementParameterToFamilyParameter(param, famparam)

spFile = doc.Application.OpenSharedParameterFile()

set = doc.FamilyManager.Parameters

paraNames = ['ADSK_Полная мощность', 'ADSK_Коэффициент мощности', 'ADSK_Количество фаз', 'ADSK_Напряжение',
             'ADSK_Классификация нагрузок', 'ADSK_Не нагреватель_Не шкаф', 'ADSK_Номинальная мощность']

notFormula = ['ADSK_Полная мощность', 'ADSK_Коэффициент мощности', 'ADSK_Количество фаз']


#проверяем тот ли файл общих параметров подгружен
spFileName = str(doc.Application.SharedParametersFilename)
spFileName = spFileName.split('\\')
spFileName = spFileName[-1]
if "ФОП_ADSK.txt" != spFileName:
    print 'Подгружен не тот файл общих параметров, переключитесь на ФОП_ADSK'
    sys.exit()



with revit.Transaction("Добавление параметров"):
    #удаляем формулы на элементах, к которым будут присвоены свои, чтоб избежать конфликтов
    for param in set:
        if str(param.Definition.Name) in notFormula:
            manager.SetFormula(param, "1")

    #проверяем наличие параметров из списка имен в проекте, присваиваем им параметры типа или экземпляра
    for param in set:
        if str(param.Definition.Name) in paraNames:
            #мощность и напряжение иногда уже проставлены и могут быть любого типа, просто их не трогаю
            if str(param.Definition.Name) == 'ADSK_Номинальная мощность' or str(param.Definition.Name) == 'ADSK_Напряжение':
               pass
            else:
                manager.MakeInstance(param)
            paraNames.remove(param.Definition.Name)

    #если в списке имен после проверки осталось что-то, добавляем параметры из списка
    if len(paraNames) > 0:
        addedNames = []
        for name in paraNames:
            for dG in spFile.Groups:
                group = "04 Обязательные ИНЖЕНЕРИЯ"
                if name == 'ADSK_Не нагреватель_Не шкаф':
                    group = "08 Необязательные ИНЖЕНЕРИЯ"
                if str(dG.Name) == group:
                    myDefinitions = dG.Definitions
                    eDef = myDefinitions.get_Item(name)
                    if name == 'ADSK_Номинальная мощность' or name == 'ADSK_Напряжение' or name == 'ADSK_Не нагреватель_Не шкаф':
                        manager.AddParameter(eDef, BuiltInParameterGroup.PG_ELECTRICAL_LOADS, False)
                    else:
                        manager.AddParameter(eDef, BuiltInParameterGroup.PG_ELECTRICAL_LOADS, True)
                    addedNames.append(name)

        #если все равно что-то осталось, сообщаем каких параметров не нашлось, но это вряд ли произойдет
        for name in addedNames:
            if name in paraNames:
                paraNames.remove(name)
        if len(paraNames) > 0:
            print 'Проблема при добавлении параметров'
            for name in paraNames:
                print name

    #если в семействе нет никаких типоразмеров, ревит почему-то не даст создать формулы. Создаем хотя бы один тип
    typeNumber = 0
    for type in doc.FamilyManager.Types:
        typeNumber = typeNumber + 1
    if typeNumber < 1:
        doc.FamilyManager.NewType('Стандарт')


