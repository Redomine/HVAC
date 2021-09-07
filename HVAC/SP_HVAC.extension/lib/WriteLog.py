#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

import datetime
from System import DateTime
import getpass

def SetLogFile(content, docum):
    
    user = getpass.getuser()
    date_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    data = str(date_time) + "   " + user + "   " + content + "   " + docum.Title + "\n" 
    
    with open("N:\BIM\КУ_Ув\Logs.txt", "a") as file:
        file.write(data.encode('utf8'))