#! /usr/bin/env python
# -*- coding: utf-8 -*-
import clr
clr.AddReference("Microsoft.Office.Interop.Outlook")
from System.Runtime.InteropServices import Marshal

import WriteLog

doc = __revit__.ActiveUIDocument.Document

mail= Marshal.GetActiveObject("Outlook.Application").CreateItem(0)
mail.Recipients.Add("gumin@spgr.ru; BIMSupport@spectrum-group.ru")
mail.Subject = "HELP. Спектрум-Электрика"
mail.Body = ""
mail.Display(True)

WriteLog.SetLogFile("ЭОМ_Справка", doc)