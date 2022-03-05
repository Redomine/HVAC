#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Обновление ФОП ЭФ'
__doc__ = "Обновляет экономическую функцию"

import os.path as op

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

import sys
from Autodesk.Revit.DB import *
from System import Guid
from pyrevit import revit


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
colPlumbingFixtures= make_col(BuiltInCategory.OST_PlumbingFixtures)

collections = [colFittings, colPipeFittings, colCurves, colFlexCurves, colFlexPipeCurves, colTerminals, colAccessory,
               colPipeAccessory, colEquipment, colInsulations, colPipeInsulations, colPipeCurves, colPlumbingFixtures]


colDuctSystems = make_col(BuiltInCategory.OST_DuctSystem)
colPipeSystems = make_col(BuiltInCategory.OST_PipingSystem)



def getEFsystem(element):
    sys_name = element.LookupParameter('Имя системы').AsString()
    EF = None
    if sys_name != None:
        if element in colEquipment:
            sys_name = element.LookupParameter('Имя системы').AsString()
            sys_name = sys_name.split(',')
            sys_name = sys_name[0]

        if sys_name in ductDict:
            EF = ductDict[sys_name]
        if sys_name in pipeDict:
            EF = pipeDict[sys_name]

    if EF == None:
        EF = doc.ProjectInformation.LookupParameter('ФОП_Экономическая функция').AsString()
    return EF

def copyEF(collection):
    for element in collection:
        EF = getEFsystem(element)
        if EF != None:
            ElemTypeId = element.GetTypeId()
            ElemType = doc.GetElement(ElemTypeId)
            if ElemType.get_Parameter(Guid('23772cae-9eaa-4f96-99ba-b65a7f44f8cf')):
                if ElemType.get_Parameter(Guid('23772cae-9eaa-4f96-99ba-b65a7f44f8cf')).AsString() != None:
                    if ElemType.get_Parameter(Guid('23772cae-9eaa-4f96-99ba-b65a7f44f8cf')) != "":
                        EF = ElemType.get_Parameter(Guid('23772cae-9eaa-4f96-99ba-b65a7f44f8cf')).AsString()

            element.LookupParameter('ФОП_Экономическая функция').Set(EF)



def getDependent(collection):
    for element in collection:
        EF = element.LookupParameter('ФОП_Экономическая функция').AsString()
        dependent = element.GetSubComponentIds()

        for depend in dependent:
            depend.LookupParameter('ФОП_Экономическая функция').Set(EF)

def getSystemDict(collection):
    Dict = {}
    for system in collection:
        if system.Name not in Dict:
            ElemTypeId = system.GetTypeId()
            ElemType = doc.GetElement(ElemTypeId)
            if ElemType.get_Parameter(Guid('23772cae-9eaa-4f96-99ba-b65a7f44f8cf')):
                if ElemType.get_Parameter(Guid('23772cae-9eaa-4f96-99ba-b65a7f44f8cf')) != None:
                    if ElemType.get_Parameter(Guid('23772cae-9eaa-4f96-99ba-b65a7f44f8cf')) != "":
                        EF = ElemType.get_Parameter(Guid('23772cae-9eaa-4f96-99ba-b65a7f44f8cf')).AsString()

            if system.LookupParameter('ФОП_Экономическая функция').AsString() != None:
                if system.LookupParameter('ФОП_Экономическая функция').AsString() != "":
                    EF = system.LookupParameter('ФОП_Экономическая функция').AsString()
            Dict[system.Name] = EF
    return Dict


paraNames = ['ФОП_Экономическая функция', 'ФОП_ВИС_Экономическая функция']
#проверка на наличие нужных параметров
map = doc.ParameterBindings
it = map.ForwardIterator()
while it.MoveNext():
    newProjectParameterData = it.Key.Name
    if str(newProjectParameterData) in paraNames:
        paraNames.remove(str(newProjectParameterData))
if len(paraNames) > 0:
    print 'Необходимо добавить параметры'
    for name in paraNames:
        print name
    sys.exit()



try:
    if doc.ProjectInformation.LookupParameter('ФОП_Экономическая функция').AsString() == None:
        print 'ФОП_Экономическая функция не заполнен в сведениях о проекте'
        sys.exit()
except Exception:
    print 'Не найден параметр ФОП_Экономическая функция'
    sys.exit()

ductDict = getSystemDict(colDuctSystems)

pipeDict = getSystemDict(colPipeSystems)


with revit.Transaction("Обновление общей спеки"):
    for collection in collections:
        copyEF(collection)