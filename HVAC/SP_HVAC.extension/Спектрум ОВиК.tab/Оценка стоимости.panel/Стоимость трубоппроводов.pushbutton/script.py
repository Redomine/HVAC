#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Расчёт стоимости\nтруб и фасонины'
__doc__ = "Всем трубопроводам, размещённым в модели, присваиваются стоимости из ключевой спецификации ОВиК_Стоимость труб по ключам"

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

pipes = FilteredElementCollector(doc)\
                                    .OfCategory(BuiltInCategory.OST_PipeCurves)\
                                    .WhereElementIsNotElementType()\
                                    .ToElements()

ViewSchedules = FilteredElementCollector(doc)\
                                            .OfCategory(BuiltInCategory.OST_Schedules)\
                                            .WhereElementIsNotElementType()\
                                            .ToElements()

for i in ViewSchedules:
    if i.Name == "ОВиК_Стоимость труб по ключам":
        ViewScheduleId = i.Id

elements = FilteredElementCollector(doc, ViewScheduleId).ToElements()
params = []
for i in elements:
    params.append(i)

def SearchKeyParameterByName(ParameterName, KeyParameters):
    KeyParameterId = 0
    for i in KeyParameters:
        if (ParameterName == i.Name):
            KeyParameterId = i.Id
    return KeyParameterId

t = Transaction(doc, 'Расчёт стоимости труб и фасонины')
t.Start()

for i in pipes:
    PipeTypeId = i.GetTypeId()
    PipeType = doc.GetElement(PipeTypeId)
    SegmentName = i.LookupParameter("Описание сегмента").AsString()
    GOSTName = PipeType.LookupParameter("О_Обозначение").AsString()
    Size = i.LookupParameter("Размер").AsString()
    Size2 = Size.split(' мм')[0]
    if GOSTName is not "":
        Name = SegmentName + " " + Size2 + " " + GOSTName
    else:
        Name = SegmentName + " " + Size2 + ""
    CostKeyName = SearchKeyParameterByName(Name, params)
    CostKey = i.LookupParameter("Ключ для стоимости")
    CostKey.Set(CostKeyName)
    print Name

t.Commit()