#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Пакетное\nсоздание схем'
__doc__ = "Разбить активный вид на схемы систем"


import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

import sys
import System
import WriteLog
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from rpw.ui.forms import select_file
from rpw.ui.forms import TextInput
from rpw.ui.forms import SelectFromList
from rpw.ui.forms import Alert
from System.Collections.Generic import List
from rpw.ui.forms import SelectFromList
import System.Drawing
import System.Windows.Forms
from System.Drawing import *
from System.Windows.Forms import *
    
doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView


class MainForm_a(Form):
    
    def __init__(self, levels):

        LEN_SYMBOL = 40

        xOffset = 0
        yOffset = 0
        isMouseDown = False

        self.xOffset = xOffset
        self.yOffset = yOffset
        self.isMouseDown = isMouseDown
        mouseOffset = Point(0, 0)
        self.mouseOffset = mouseOffset

        params = FilteredElementCollector(doc).OfClass(SharedParameterElement)


        self.name_param = levels
        self.values = []

        self._button2 = Button()
        self._button3 = Button()
        self._label1 = Label()
        self._button4 = Button()
        self._checkedListBox1 = CheckedListBox()
        self._textBox1 = TextBox()
        self._label2 = Label()
        self._checkBox1 = CheckBox()
        self.SuspendLayout()
        # 

        # 
        # button2
        # 
        self._button2.BackColor = Color.FromArgb(15, 27, 40)
        self._button2.FlatAppearance.BorderColor = Color.FromArgb(95, 107, 120)
        self._button2.FlatAppearance.MouseOverBackColor = Color.Gray
        self._button2.FlatStyle = FlatStyle.Flat
        self._button2.Font = Font("Segoe UI", 10, FontStyle.Regular, GraphicsUnit.Point, 204)
        self._button2.ForeColor = Color.FromArgb(230, 230, 230)
        self._button2.Location = Point(620, 437)
        self._button2.Name = "button2"
        self._button2.Size = Size(197, 47)
        self._button2.TabIndex = 1
        self._button2.Text = "Отмена"
        self._button2.UseVisualStyleBackColor = False
        self._button2.Click += self.onExit
        # 
        # button3
        # 
        self._button3.BackColor = Color.FromArgb(15, 27, 40)
        self._button3.FlatAppearance.BorderColor = Color.FromArgb(95, 107, 120)
        self._button3.FlatAppearance.MouseOverBackColor = SystemColors.GrayText
        self._button3.FlatStyle = FlatStyle.Flat
        self._button3.Font = Font("Segoe UI", 10, FontStyle.Regular, GraphicsUnit.Point, 204)
        self._button3.ForeColor = Color.FromArgb(190, 190, 190)
        self._button3.Location = Point(621, 331)
        self._button3.Name = "button3"
        self._button3.Size = Size(196, 47)
        self._button3.TabIndex = 1
        self._button3.Text = "Готово"
        self._button3.UseVisualStyleBackColor = False
        self._button3.Click += self.ButtonClicked
        # 
        # label1
        # 
        self._label1.Font = Font("Segoe UI", 13, FontStyle.Regular, GraphicsUnit.Point, 204)
        self._label1.ForeColor = Color.FromArgb(190, 190, 190)
        self._label1.Location = Point(49, 33)
        self._label1.Name = "label1"
        self._label1.Size = Size(378, 33)
        self._label1.TabIndex = 2
        self._label1.Text = "Выберите рабочие элементы"
        # 
        # button4
        # 
        self._button4.BackColor = Color.FromArgb(65, 77, 90)
        self._button4.FlatAppearance.BorderColor = Color.FromArgb(100, 100, 100)
        self._button4.FlatAppearance.MouseOverBackColor = Color.Gray
        self._button4.FlatStyle = FlatStyle.Flat
        self._button4.Font = Font("Segoe UI Light", 11, FontStyle.Regular, GraphicsUnit.Point, 204)
        self._button4.ForeColor = Color.FromArgb(230, 230, 230)
        self._button4.Location = Point(809, 2)
        self._button4.Name = "button4"
        self._button4.Size = Size(37, 31)
        self._button4.TabIndex = 1
        self._button4.Text = "X"
        self._button4.UseVisualStyleBackColor = False
        self._button4.Click += self.onExit
        # 
        # checkedListBox1
        #
        
        for i in self.name_param:
            if i == None: continue
            self._checkedListBox1.Items.Add(i)

        self._checkedListBox1.BackColor = Color.FromArgb(25, 37, 50)
        self._checkedListBox1.BorderStyle = BorderStyle.None
        self._checkedListBox1.CheckOnClick = True
        self._checkedListBox1.Font = Font("Segoe UI Emoji", 8, FontStyle.Regular, GraphicsUnit.Point, 204)
        self._checkedListBox1.ForeColor = Color.FromArgb(190, 190, 190)
        self._checkedListBox1.FormattingEnabled = True
        self._checkedListBox1.Location = Point(26, 88)
        self._checkedListBox1.Name = "checkedListBox1"
        self._checkedListBox1.Size = Size(561, 396)
        self._checkedListBox1.TabIndex = 3
        # 
        # textBox1
        # 
        self._textBox1.BackColor = Color.FromArgb(25, 37, 50)
        self._textBox1.BorderStyle = BorderStyle.None
        self._textBox1.Font = Font("Segoe UI", 10, FontStyle.Regular, GraphicsUnit.Point, 204)
        self._textBox1.ForeColor = Color.FromArgb(224, 224, 224)
        self._textBox1.Location = Point(620, 119)
        self._textBox1.Name = "textBox1"
        self._textBox1.Size = Size(182, 16)
        self._textBox1.TabIndex = 4
        self._textBox1.TextChanged += self.RemoveListCheckBox
        self._textBox1.Enter += self.RemoveListCheckBox
        # 
        # label2
        # 
        self._label2.Font = Font("Segoe UI", 10, FontStyle.Regular, GraphicsUnit.Point, 204)
        self._label2.ForeColor = Color.FromArgb(190, 190, 190)
        self._label2.Location = Point(621, 88)
        self._label2.Name = "label2"
        self._label2.Size = Size(182, 23)
        self._label2.TabIndex = 5
        self._label2.Text = "Фильтр"
        # 
        # CheckBOX1
        # 
        self._checkBox1.FlatAppearance.BorderColor = Color.FromArgb(15, 27, 40)
        self._checkBox1.FlatAppearance.CheckedBackColor = Color.FromArgb(15, 27, 40)
        self._checkBox1.FlatAppearance.MouseDownBackColor = Color.FromArgb(15, 27, 40)
        self._checkBox1.FlatAppearance.MouseOverBackColor = Color.FromArgb(15, 27, 40)
        self._checkBox1.FlatStyle = FlatStyle.Flat
        self._checkBox1.Font = Font("Segoe UI", 10, FontStyle.Regular, GraphicsUnit.Point, 204)
        self._checkBox1.ForeColor = Color.FromArgb(190, 190, 190)
        self._checkBox1.Location = Point(621, 154)
        self._checkBox1.Name = "checkBox1"
        self._checkBox1.Size = Size(196, 24)
        self._checkBox1.TabIndex = 6
        self._checkBox1.Text = "Выбрать все"
        self._checkBox1.UseVisualStyleBackColor = True
        self._checkBox1.CheckedChanged += self.SelectedAll
        # 
        # MainForm
        # 
        self.BackColor = Color.FromArgb(15, 27, 40)
        self.BackgroundImageLayout = ImageLayout.None
        self.ClientSize = Size(848, 508)
        self.Controls.Add(self._label2)
        self.Controls.Add(self._textBox1)
        self.Controls.Add(self._checkedListBox1)
        self.Controls.Add(self._label1)
        self.Controls.Add(self._button2)
        self.Controls.Add(self._button4)
        self.Controls.Add(self._button3)
        self.Controls.Add(self._checkBox1)
        self.Font = Font("Segoe UI", 8.25, FontStyle.Regular, GraphicsUnit.Point, 204)
        self.ForeColor = Color.FromArgb(224, 224, 224)
        self.FormBorderStyle = FormBorderStyle.None
        self.Name = "MainForm"
        self.StartPosition = FormStartPosition.CenterScreen
        self.Text = "тест приложения"
        self.MouseDown += self.Form1_MouseDown
        self.MouseMove += self.Form1_MouseMove
        self.MouseUp += self.Form1_MouseUp
        self.ResumeLayout(False)
        self.PerformLayout()
        
    def onExit(self, sender, e):
        self.Close()

    def Form1_MouseDown(self, sender, e):

        self.xOffset = -e.X - SystemInformation.FrameBorderSize.Width
        self.yOffset = -e.Y - SystemInformation.CaptionHeight - SystemInformation.FrameBorderSize.Height
        self.mouseOffset = Point(self.xOffset, self.yOffset)
        self.isMouseDown = True
        
    def Form1_MouseMove(self, sender, e):
        
        if self.isMouseDown:
            mousePos = Control.MousePosition
            mousePos.Offset(self.mouseOffset.X, self.mouseOffset.Y)
            Location = mousePos
            self.Left = Location.X
            self.Top = Location.Y
            
    def Form1_MouseUp(self, sender, e):
        if e.Button == MouseButtons.Left:
            self.isMouseDown = False

    def ButtonClicked(self, sender, args):
        if sender.Click:
            for i in self._checkedListBox1.CheckedItems:
                self.values.append(i)
                self.Close()

    def RemoveListCheckBox (self, sender, args):

        text = self._textBox1.Text

        self._checkedListBox1.Items.Clear()

        if len(text) == 0:
            for i in self.name_param:
                if i == None:
                    continue
                self._checkedListBox1.Items.Add(i)
                if i == None:
                    continue

        if len(text) > 0:
            for i in self.name_param:
                a = type(i)
                if a is not str:
                    continue
                if text in i:
                    self._checkedListBox1.Items.Add(i)

    def SelectedAll(self, sender, e):

        if sender.Checked:

            for i in range(0,len(self._checkedListBox1.Items)):
                if i == None:
                    continue
                self._checkedListBox1.SetItemChecked(i, True)

        else:
            for i in range(0,len(self._checkedListBox1.Items)):
                self._checkedListBox1.SetItemChecked(i, False)
                
                
def cheklist(input_list):
    form = MainForm_a(input_list)#выводим форму вывода семейств, назначем классу объект для возможности вывода
    form.ShowDialog()#???
    input_list = form.values#выводим список семейств для обработки
    return input_list

#для вентиляции отфильтруем приточно-вытяжные установки, чтобы назначить их фильтр


def create_filter_view(project_todo, element, systems, master_view, filter_name):
    rules = []
    
    for rule in systems:
        if project_todo == 'Вентиляция':
            rules.append(ParameterFilterRuleFactory.CreateNotEqualsRule(ElementId(BuiltInParameter.RBS_SYSTEM_NAME_PARAM), rule, rule))
        else:
            rules.append(ParameterFilterRuleFactory.CreateNotContainsRule(ElementId(BuiltInParameter.RBS_SYSTEM_NAME_PARAM), rule, rule))
    filter_name = '_скрипт' + filter_name
    if (ParameterFilterElement.IsNameUnique(doc, filter_name)):
        try:
            filter = ParameterFilterElement.Create(doc, filter_name, categories, rules)
        except Exception:
            filter = ParameterFilterElement.Create(doc, filter_name, categories)
            elemFilters = []
            for rule in rules:
                elemParamFilter = ElementParameterFilter(rule)
                elemFilters.append(elemParamFilter)
            elemFilter = LogicalAndFilter(elemFilters)
            filter.SetElementFilter(elemFilter)
        copy_view_eid = master_view.Duplicate(ViewDuplicateOption.WithDetailing)
        copy_view = doc.GetElement(copy_view_eid)
        copy_view.Name = filter_name
        copy_view.AddFilter(filter.Id)
        copy_view.SetFilterVisibility(filter.Id,False)
        rules = []
    else:
        print 'Для системы', element, 'уже создан фильтр. Удалите его или проверьте правильность'


    

project_todo = SelectFromList('Выберите тип схемы', ['Вентиляция','Трубы'])


colEquipment = FilteredElementCollector(doc)\
                            .OfCategory(BuiltInCategory.OST_MechanicalEquipment)\
                            .WhereElementIsNotElementType()\
                            .ToElements()



Equipment = [] #сюда поместим оборудование, с системами которые его обслуживают
for fan in colEquipment:
    if fan.LookupParameter("Имя системы"):
        Equipment.append([])
        Equipment[-1].append(fan)
        fan_system = fan.LookupParameter("Имя системы")
        Equipment[-1].append(fan_system.AsString())

systemfilter = ElementClassFilter(MEPSystem)

if project_todo == 'Вентиляция':
    e_filter = ElementCategoryFilter(BuiltInCategory.OST_DuctSystem)
else:
    e_filter = ElementCategoryFilter(BuiltInCategory.OST_PipingSystem)

ductsystemfilter = LogicalAndFilter(systemfilter, e_filter)

collector = FilteredElementCollector(doc)

system = collector.WherePasses(ductsystemfilter).ToElements()


categories = []

categories = List[ElementId](categories)


categories.Add(ElementId(BuiltInCategory.OST_DuctCurves))
categories.Add(ElementId(BuiltInCategory.OST_DuctFitting))
categories.Add(ElementId(BuiltInCategory.OST_DuctAccessory))
categories.Add(ElementId(BuiltInCategory.OST_DuctTerminal))
categories.Add(ElementId(BuiltInCategory.OST_MechanicalEquipment))
categories.Add(ElementId(BuiltInCategory.OST_FlexDuctCurves))
categories.Add(ElementId(BuiltInCategory.OST_DuctInsulations))
categories.Add(ElementId(BuiltInCategory.OST_PipeCurves))
categories.Add(ElementId(BuiltInCategory.OST_FlexPipeCurves))
categories.Add(ElementId(BuiltInCategory.OST_PipeFitting))
categories.Add(ElementId(BuiltInCategory.OST_PipeAccessory))
categories.Add(ElementId(BuiltInCategory.OST_MechanicalEquipment))
categories.Add(ElementId(BuiltInCategory.OST_PipeInsulations))    




master_view = doc.GetElement(view.Id) 
t = Transaction(doc, 'Формирование схем')

t.Start()

#нужная логика - есть список систем, по которым нужно составить список, составляем схему по системе, проверяем есть ли среди оборудования что-то с такой системой в свойствах,
#если есть, добавляем оставшиеся системы из свойств оборудования в схему
#Так же полученные системы перед добавлением в схему надо проверить на соответствие выбранному фильтру, чтобы не добавлялись трубопроводные в схемы вентиляции и наоборот
#

checked_systems = []
for element in system:
    if element.Name not in checked_systems: checked_systems.append(element.Name)
    
checked_systems = cheklist(checked_systems)    

secondary_systems = [] #этот список наполняется вторичными системами от оборудования которые уже были построены на схемах. Если проверяемая система в списке то пропускаем цикл.

for element in checked_systems:
    first_mark = element
    
    if first_mark in secondary_systems:
       continue
    equipment_systems = [first_mark]
    for equipment in Equipment:
        if equipment[1] == None:
            continue
        if first_mark in equipment[1].split(','):
            system_to_confirm = equipment[1].split(',')

            
            for confirm in system:
                try:
                    if confirm.Name in system_to_confirm:
                        if confirm.Name != first_mark:
                            equipment_systems.append(confirm.Name)
                except Exception:
                    pass
            for equipment_system in equipment_systems:
                secondary_systems.append(equipment_system)
            if project_todo == 'Вентиляция':
                equipment_systems.append(equipment[1])
            break

    filter_name = ''
    for name in equipment_systems:
        filter_name = filter_name + name
    create_filter_view(project_todo, element, equipment_systems, master_view, filter_name)
t.Commit()

    
