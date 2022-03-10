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
    connectors = getConnectors(element)
    orient = 1
    connector_out = []
    for connector in connectors:
        if str(connector.Direction) == "Out":
            connector_out.append(connector)

        #для входа в тройник ищем координаты начала входящего воздуховода чтоб построить прямую через эти две точки
        connector_p = []
        if str(connector.Direction) == "In":
            inTee = getConCoords(connector)
            a = connector.AllRefs.ForwardIterator()
            while a.MoveNext():
                connector_p.append(a.Current)
            duct = connector_p[0].Owner
            duct_connectors = getConnectors(duct)
            for duct_connector in duct_connectors:
                if getConCoords(duct_connector) != inTee:
                    inDuct = getConCoords(duct_connector)


    # среди выходящих коннекторов ищем диктующий по большему расходу
    try:
        if connector_out[0].Flow > connector_out[1].Flow:
            connector_o = connector_out[0]
        else:
            connector_o = connector_out[1]
    except Exception:
        print element.Id
        print connector_out
    connector_o = getConCoords(connector_o)

    #найдем вектор по координатам точек AB = {Bx - Ax; By - Ay; Bz - Az}
    Tee_Duct = [(float(inDuct[0]) - float(inTee[0])), (float(inDuct[1]) - float(inTee[1])), (float(inDuct[2]) - float(inTee[2]))]
    Tee_Out = [(float(inTee[0]) - float(connector_o[0])), (float(inTee[1]) - float(connector_o[1])), (float(inTee[2]) - float(connector_o[2]))]

    #найдем скалярное произведение векторов AB · CD = ABx · CDx + ABy · CDy + ABz · CDz
    Tee_Out_Tee_Duct = Tee_Duct[0]*Tee_Out[0] + Tee_Duct[1]*Tee_Out[1] + Tee_Duct[2]*Tee_Out[2]

    #найдем длины векторов
    len_Tee_Duct = ((Tee_Duct[0])**2 + (Tee_Duct[1])**2 + (Tee_Duct[2])**2)**0.5
    len_Tee_Out = ((Tee_Out[0])**2 + (Tee_Out[1])**2 + (Tee_Out[2])**2)**0.5

    #найдем косинус
    cos = (Tee_Out_Tee_Duct)/(len_Tee_Duct * len_Tee_Out)

    #Если угол расхождения между вектором входа воздуха и выхода больше 10 градусов(цифра с потолка) то считаем что идет буквой L
    #Если нет, то считаем что идет по прямой буквой I
    if math.acos(cos) > 0.10:
        orient = 'L'
    else:
        orient = 'I'
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

