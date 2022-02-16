#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Разбиение уровня'
__doc__ = "Разбиение уровня"

import os.path as op

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
from itertools import groupby

from pyrevit import revit
from pyrevit import forms
from pyrevit import script
from pyrevit.forms import Reactive, reactive
from pyrevit.revit import selection, Transaction

doc = __revit__.ActiveUIDocument.Document  # type: Document
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

def getLevel(collection):
    for element in collection:
        if element.LookupParameter('Отметка посередине'):
            level = element.LookupParameter('Отметка посередине').AsDouble() * 304.8
            if element.LookupParameter('Базовый уровень').AsValueString() == '01 этаж':
                level = element.LookupParameter('Отметка посередине').AsDouble() * 304.8 + 4050
            if element.LookupParameter('Базовый уровень').AsValueString() == '02 этаж':
                level = element.LookupParameter('Отметка посередине').AsDouble() * 304.8 + 8850

        if element.LookupParameter('Отметка от уровня'):

            level = element.LookupParameter('Отметка от уровня').AsDouble() * 304.8
            if element.LookupParameter('Уровень').AsValueString() == '01 этаж':
                level = element.LookupParameter('Отметка от уровня').AsDouble() * 304.8 + 4050
            if element.LookupParameter('Уровень').AsValueString() == '02 этаж':
                level = element.LookupParameter('Отметка от уровня').AsDouble() * 304.8 + 8850



        element.LookupParameter('Уровень размещения абсолютный').Set(level)
    

with revit.Transaction("Обновление общей спеки"):
    getLevel(colPipeFittings)
    getLevel(colPipeAccessory)
    getLevel(colPipeCurves)
    getLevel(colEquipment)