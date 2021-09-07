#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = '2.Распределить расходы'
__doc__ = "Распредление вынесенных из таблиц нагрузок по оборудованию в пространствах"


import clr
import WriteLog
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from itertools import compress
from rpw.ui.forms import Alert
from rpw.ui.forms import select_file
from rpw.ui.forms import SelectFromList
import System.Drawing
import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *



import WriteLog

def parameter_names(project_todo):
    names = []
    if project_todo == 'Вентиляция':
        names.append('ИОС_Расход воздуха приточный')
        names.append('ИОС_Расход воздуха вытяжной')
    if project_todo == 'Отопление':
        names.append('ИОС_Теплопотери')
        names.append('-')
    if project_todo == 'Кондиционирование':
        names.append('ИОС_Теплопритоки')
        names.append('-')
    return names

def checkbox_header(project_todo):
    names = []
    if project_todo == 'Вентиляция':
        names.append('Воздухораспределители приточные')
        names.append('Воздухораспределители вытяжные')
    if project_todo == 'Отопление':
        names.append('Выберите отопительные приборы')
        names.append('-')
    if project_todo == 'Кондиционирование':
        names.append('Выберите внутренние блоки')
        names.append('-')
    return names

#форма сформирована через приложение Sharpdevelop, через нее мы выберем семейства которым будем назначать расход. На вход даем заголовок для нее и список оборудования в проекте.
class MainForm(Form):
    
    def __init__(self, form_header, colHVACequipment):

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


        self.values = []
        self.FamilyNames = []
        for equipment in colHVACequipment:
            FamilyName = equipment.Symbol.Family.Name
            if FamilyName not in self.FamilyNames:
                self.FamilyNames.append(FamilyName)
        
        

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
        for name in self.FamilyNames:
            self._checkedListBox1.Items.Add(name)
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
        self._label1.Text = form_header
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
        
        #for i in self.name_param:
            #if i == None: continue
            #self._checkedListBox1.Items.Add(i)

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
            for i in self.FamilyNames:
                if i == None:
                    continue
                self._checkedListBox1.Items.Add(i)
                if i == None:
                    continue

        if len(text) > 0:
            for i in self.FamilyNames:
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
        
def get_HVAC_equipment_in_space(colHVACequipment, input_spaces):#функция для возврата списка с вложенными списками пространств с оборудованием в них
    spaces = []#список пространств для того чтоб не пользоваться словарями. Используем для проверки встречалось пространство или нет.
    HVAC_spaces = []#сюда закидываем вложеные списки с номером пространства и всем оборудованием в нем,
    #формата [номер пространства, оборудование_1...оборудвание_n]
    for equipment in colHVACequipment:
        FamilyName = equipment.Symbol.Family.Name    
        for input in input_spaces:
            if input == FamilyName:  
                phase = equipment.CreatedPhaseId #Следующие две строки - фаза. Без нее ревит не даст нам информации по элементу, просто пользуйтесь ей, чтоб получить объект.
                equipmentPhase = doc.GetElement(phase)
                ItemSpace = equipment.Space[equipmentPhase] #создаем объект простраства с которым можно работать
                SpaceName = Element.Name.__get__(ItemSpace).ToString()#проверка на адекватность пространств
                if SpaceName == 'IronPython.Runtime.Types.ReflectedProperty':
                    continue
     
                SpaceNumber = SpatialElement.Number.__get__(ItemSpace).ToString()#забираем номер пространства для оборудования
    
           #дальше логика такая: если номер пространства мы встречаем в первый раз, добавляем новый вложенный список в HVAC_spaces, и номер пространства в
           #в spaces, если не в первый раз - смотрим по spaces его позицию в списке и добавляем по этому номеру в список внутри HVAC_spaces оборудование в нужное
           #пространство
                if equipment.LookupParameter('ИОС_Номер пространства'):
                    SpaceNumberInEquipment = equipment.LookupParameter('ИОС_Номер пространства')
                    SpaceNumberInEquipment.Set(SpaceNumber)
                if SpaceNumber not in spaces:
                    HVAC_spaces.append([])
                    spaces.append(SpaceNumber)
                    HVAC_spaces[-1].append(SpaceNumber)
                    HVAC_spaces[-1].append(equipment)
                else:
                    HVAC_spaces[spaces.index(SpaceNumber)].append(equipment)
            
    return HVAC_spaces


project_todo = SelectFromList('Выберите тип таблицы', ['Вентиляция','Отопление','Кондиционирование'])

doc = __revit__.ActiveUIDocument.Document


colSpaces = FilteredElementCollector(doc)\
                            .OfCategory(BuiltInCategory.OST_MEPSpaces)\
                            .WhereElementIsNotElementType()\
                            .ToElements()

colSpacesNames = []

for Space in colSpaces:
    SpaceName = Element.Name.__get__(Space)
    colSpacesNames.append(SpaceName)        
    
space_in_model = []

#формируем список пространств в модели

for Space in colSpaces:
    if Space.Location:
        if Space.LookupParameter("Номер"):
            SpaceNumber = Space.LookupParameter("Номер")
            SpaceNumber = SpaceNumber.AsString()
            space_in_model.append(SpaceNumber)
            
#в зависимости от выбранного проекта выбираем фильтр по воздухораспределителям или по оборудованию
            
if project_todo == 'Вентиляция':
    colHVACequipment = FilteredElementCollector(doc)\
                                .OfCategory(BuiltInCategory.OST_DuctTerminal)\
                                .WhereElementIsNotElementType()\
                                .ToElements()
else:
    colHVACequipment = FilteredElementCollector(doc)\
                                .OfCategory(BuiltInCategory.OST_MechanicalEquipment)\
                                .WhereElementIsNotElementType()\
                                .ToElements()
    
header = checkbox_header(project_todo)#выбираем заголовки для формы выбора семейств 
form_supply = MainForm(header[0], colHVACequipment)#выводим форму вывода семейств, назначем классу объект для возможности вывода
form_supply.ShowDialog()#???
input_spaces_supply = form_supply.values#выводим список семейств для обработки

if project_todo == 'Вентиляция':#то же что и выше, но для вытяжки
    form_extract = MainForm(header[1], colHVACequipment)
    form_extract.ShowDialog()    
    input_spaces_extract = form_extract.values

t = Transaction(doc, 'Перенос расходов в оборудование')

t.Start()

count = 0
        
HVAC_spaces_supply = get_HVAC_equipment_in_space(colHVACequipment, input_spaces_supply)
if project_todo == 'Вентиляция':
    HVAC_spaces_extract = get_HVAC_equipment_in_space(colHVACequipment, input_spaces_extract)

names = parameter_names(project_todo)
#try:
#    for Space in colSpaces:
#        if Space.Location:
#            if Space.LookupParameter("Номер"):
#                SpaceNumber = Space.LookupParameter("Номер")
#                SpaceNumber = SpaceNumber.AsString()
#            if Space.LookupParameter(names[0]):
#                SpaceSupply = Space.LookupParameter(names[0])             
#                SpaceSupply = SpaceSupply.AsDouble()
#            if project_todo == 'Вентиляция':
#                if Space.LookupParameter(names[1]):
##                    SpaceExtract = Space.LookupParameter(names[1])
 #                   SpaceExtract = SpaceExtract.AsDouble()
 #           for number in HVAC_spaces_supply: 
 #               if number[0] == SpaceNumber:                   
  #                  amount_of_equipment = len(number) - 1 #из длины списка вычитаю номер помещения
   #                 for equipment in number[1:]:
    #                    if Space.LookupParameter(names[0]):
     #                       SpaceSupply = Space.LookupParameter(names[0]).AsDouble()
      #                  if equipment.LookupParameter(names[0]):
       #                     EquipmentSupply = equipment.LookupParameter(names[0])
        #                    EquipmentSupply.Set(SpaceSupply/amount_of_equipment)
                            
 #           if project_todo == 'Вентиляция':
  #              for number in HVAC_spaces_extract:
   #                 if number[0] == SpaceNumber:
    #                    amount_of_equipment = len(number) - 1
     #                   for equipment in number[1:]:
      #                      if equipment.LookupParameter('Тип системы'):
       #                         system_type = equipment.LookupParameter('Тип системы').AsValueString()
        #                        if system_type == 'СП_Вентиляция_Аварийная' or system_type == 'СП_Вентиляция_Местные отсосы' or system_type == 'СП_ПДЗ_Дымоудаление':
         #                           amount_of_equipment-=1
                            
          #              for equipment in number[1:]:
           #                 if equipment.LookupParameter('Тип системы'):
            #                    system_type = equipment.LookupParameter('Тип системы').AsValueString()
             #                   if system_type == 'СП_Вентиляция_Аварийная' or system_type == 'СП_Вентиляция_Местные отсосы' or system_type == 'СП_ПДЗ_Дымоудаление':
              #                     continue
               #             if Space.LookupParameter(names[1]):
                #                SpaceExtract = Space.LookupParameter(names[1]).AsDouble()
                 #           if equipment.LookupParameter(names[1]):
                  #              EquipmentExtract = equipment.LookupParameter(names[1])
                   #             EquipmentExtract.Set(SpaceExtract/amount_of_equipment)                        
                    #        
#except Exception:
#    Alert('Ошибка при распределении расходов!', title= 'Ошибка', header = 'Проблемы с моделью')
#    sys.exit()        


t.Commit()
    
WriteLog.SetLogFile("2.Распределить расходы", doc)