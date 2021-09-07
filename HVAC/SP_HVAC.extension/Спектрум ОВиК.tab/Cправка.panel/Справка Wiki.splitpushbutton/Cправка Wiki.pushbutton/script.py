#! /usr/bin/env python
# -*- coding: utf-8 -*-


"""Открыть справку Spectrum Wiki"""
from pyrevit import script

import WriteLog

__context__ = 'zerodoc'

doc = __revit__.ActiveUIDocument.Document

url = 'http://wiki.spgr.ru/index.php/%D0%9F%D0%BB%D0%B0%D0%B3%D0%B8%D0%BD_HVAC'
script.open_url(url)


WriteLog.SetLogFile("ЭОМ_Справка", doc)