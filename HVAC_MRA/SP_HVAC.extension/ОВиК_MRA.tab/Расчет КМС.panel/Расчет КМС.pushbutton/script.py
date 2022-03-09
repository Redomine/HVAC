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
from pyrevit import revit

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView



def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col 
    
colFittings = make_col(BuiltInCategory.OST_DuctFitting)    




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


def getConCoords(connector):
    a0 = connector.Origin.ToString()
    a0 = a0.replace("(", "")
    a0 = a0.replace(")", "")
    a0 = a0.split(",")
    for x in a0:
        x = float(x)
    return a0

def getTeeOrient(element):
    a = getConnectors(element)
    orient = 1
    #первым делом ищем центр масс треугольника из коннекторов тройника
    XYZ1 = getConCoords(a[0])
    XYZ2 = getConCoords(a[1])
    XYZ3 = getConCoords(a[2])

    print XYZ1
    xm = (float(XYZ1[0]) + float(XYZ2[0]) + float(XYZ3[0]))/3
    ym = (float(XYZ1[1]) + float(XYZ2[1]) + float(XYZ3[1]))/3
    zm = (float(XYZ1[2]) + float(XYZ2[2]) + float(XYZ3[2]))/3
    XYZm = (xm, ym, zm)

    #вычисляем стартовую точку прямой по максимальному потоку
    if max(a[0].Flow, a[1].Flow, a[2].Flow) == a[0].Flow:
        XYZs = getConCoords(a[0])
    if max(a[0].Flow, a[1].Flow, a[2].Flow) == a[1].Flow:
        XYZs = getConCoords(a[1])
    if max(a[0].Flow, a[1].Flow, a[2].Flow) == a[2].Flow:
        XYZs = getConCoords(a[2])

    #вычисляем конечную точку прямой по минимальному потоку
    if min(a[0].Flow, a[1].Flow, a[2].Flow) == a[0].Flow:
        XYZe = getConCoords(a[0])
    if min(a[0].Flow, a[1].Flow, a[2].Flow) == a[1].Flow:
        XYZe = getConCoords(a[1])
    if min(a[0].Flow, a[1].Flow, a[2].Flow) == a[2].Flow:
        XYZe = getConCoords(a[2])

    #составим уравнение прямой в каноническом виде по трем точкам
    xa = float(XYZs[0])
    ya = float(XYZs[1])
    za = float(XYZs[2])

    xb = float(XYZe[0])
    yb = float(XYZe[1])
    zb = float(XYZe[2])

    (xm - xa) / (xb - xa)
    (ym - ya) / (yb - ya)
    (zm - za) / (zb - za)


    return orient


def getDpTee(element):
    a = getConnectors(element)

    getTeeOrient(element)
    dp = 10
    # try:
    #     S1 = a[1].Height*0.3048*a[1].Width*0.3048
    # except:
    #     S1 = 3.14*0.3048*0.3048*a1.Radius**2
    # try:
    #     S2 = a2.Height*0.3048*a2.Width*0.3048
    # except:
    #     S2 = 3.14*0.3048*0.3048*a2.Radius**2
    # try:
    #     S3 = a3.Height*0.3048*a3.Width*0.3048
    # except:
    #     S3 = 3.14*0.3048*0.3048*a3.Radius**2
    #
    # v1 = a1.Flow*101.94/3600/S1
    # v2 = a2.Flow*101.94/3600/S2
    # v3 = a3.Flow*101.94/3600/S3
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

with revit.Transaction("Пересчет потерь напора"):
    for el in elems:
        dp = 3.3
        if str(el.MEPModel.PartType) == 'Elbow':
            dp = getDpElbow(el)

        if str(el.MEPModel.PartType) == 'Transition':
            dp = getDpTransition(el)

        if str(el.MEPModel.PartType) == 'Tee':
            dp = getDpTee(el)

        eleId = el.Id
        fitting = doc.GetElement(eleId)
        param = fitting.get_Parameter(BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM)
        lc = getLossMethods(ExternalServices.BuiltInExternalServices.DuctFittingAndAccessoryPressureDropService)
        param.Set(lc[9].ToString()) # установка метода потерь
        schema = lc[11].GetDataSchema()
        field = schema.GetField("PressureLoss")
        entity=fitting.GetEntity(schema)
        try:
            entity.Set[field.ValueType](field, str(int(math.ceil(dp/3.3))))
            fitting.SetEntity(entity)
        except Exception:
            pass



with revit.Transaction("Выключение систем"):
    colSystems = make_col(BuiltInCategory.OST_DuctSystem)
    for el in colSystems:
        sysType = doc.GetElement(el.GetTypeId())
        sysType.CalculationLevel = sysType.CalculationLevel.None

with revit.Transaction("Включение систем"):
    colSystems = make_col(BuiltInCategory.OST_DuctSystem)
    for el in colSystems:
        sysType = doc.GetElement(el.GetTypeId())
        sysType.CalculationLevel = sysType.CalculationLevel.All

