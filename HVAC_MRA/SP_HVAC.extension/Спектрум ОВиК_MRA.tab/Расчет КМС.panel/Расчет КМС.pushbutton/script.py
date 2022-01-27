#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Пересчет КМС'
__doc__ = "Пересчитывает КМС соединительных деталей воздуховодов"


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
from Autodesk.Revit.DB.ExternalService import *
from Autodesk.Revit.DB.ExtensibleStorage import *
from System.Collections.Generic import List
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


t = Transaction(doc, 'Пересчет КМС')

t.Start()

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

def getDpElbow(element):
    a = getConnectors(element)
    try:
        S1 = a[0].Height*0.3048*a[0].Width*0.3048
    except:
        S1 = 3.14*0.3048*0.3048*a[0].Radius**2

    v1 = a[0].Flow*101.94/3600/S1
    
    angle = a[1].Angle
        
    if angle > 1:
        dp = 0.32*(v1**1.8)
    elif angle<=1 and angle>0.7:
        dp = 0.32*(v1**1.8)/2
    elif angle>0 and angle<=0.7: 
        dp = 0.32*(v1**1.8)/3
    return dp

def getDpTransition(element):
    a = getConnectors(element)
    try:
        S1 = a[0].Height*0.3048*a[0].Width*0.3048
    except:
        S1 = 3.14*0.3048*0.3048*a[0].Radius**2
    try:
        S2 = a[1].Height*0.3048*a[1].Width*0.3048
    except:
        S2 = 3.14*0.3048*0.3048*a[1].Radius**2
        
    v1 = a[0].Flow*101.94/3600/S1
    v2 = a[1].Flow*101.94/3600/S2
    
    #проверяем в какую сторону дует воздух чтоб выяснить расширение это или заужение
    if str(a[0].Direction) == "In":
        v_1 = v1
        v_2 = v2
    if str(a[0].Direction) == "Out":
        v_1 = v2
        v_2 = v1
    
    if v_1 > v_2:
        dp = 0.864*(v_1 - v_2)**1.8
    if v_2 > v_1:
        dp = 0.146*(v_2 - v_1)**1.9
            

    return dp
    

        
def getLossMethods(serviceId):
    lc=[]
    service = ExternalServiceRegistry.GetService(serviceId)
    serverIds = service.GetRegisteredServerIds()
    list=List[ElementId]()
    for serverId in serverIds:
        server = getServerById(serverId, serviceId)
        id=serverId
        name=server.GetName()
        lc.append(id)
        lc.append(name)
        lc.append(server)
    return lc

def getServerById(serverGUID, serviceId):
    service = ExternalServiceRegistry.GetService(serviceId)
    if service != "null" and serverGUID != "null":
        server = service.GetServer(serverGUID)
        if server != "null":
            return server
    return null

elems=colFittings

for el in elems:
    try:
        dp = 3.3
        if str(el.MEPModel.PartType) == 'Elbow':
            dp = getDpElbow(el)
            
        if str(el.MEPModel.PartType) == 'Transition':
            dp = getDpTransition(el)
        
        eleId = el.Id
        fitting = doc.GetElement(eleId)
        param = fitting.get_Parameter(BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM)
        lc = getLossMethods(ExternalServices.BuiltInExternalServices.DuctFittingAndAccessoryPressureDropService)
        param.Set(lc[9].ToString()) # установка метода потерь

        schema = lc[11].GetDataSchema()
        field = schema.GetField("PressureLoss")
        entity=fitting.GetEntity(schema)
        entity.Set[field.ValueType](field, str(int(math.ceil(dp/3.3))))
        fitting.SetEntity(entity)
        
        

    except Exception:
        pass

t.Commit()


t_2 = Transaction(doc, 'Выключение систем')
t_2.Start()
colSystems = make_col(BuiltInCategory.OST_DuctSystem)
for el in colSystems:
    sysType = doc.GetElement(el.GetTypeId())
    sysType.CalculationLevel = sysType.CalculationLevel.None 
t_2.Commit()

t_3 = Transaction(doc, 'Включение систем')
t_3.Start()
colSystems = make_col(BuiltInCategory.OST_DuctSystem)
for el in colSystems:
    sysType = doc.GetElement(el.GetTypeId())

    sysType.CalculationLevel = sysType.CalculationLevel.All 
t_3.Commit()
