Attribute VB_Name = "Module1"
Dim Column_name As Integer
Dim Column_m As Integer
Dim Column_material As Integer
Dim Column_o As Integer
Dim Column_size As Integer
Dim Column_maker As Integer
Dim Column_dimension As Integer
Dim Column_number As Integer
Dim Column_note As Integer
Dim Column_articul As Integer
Dim Column_Fnumber As Integer


Sub VTBR_specchec()
Dim firstBook As Workbook
Dim secondBook As Workbook
Set firstBook = ThisWorkbook
Path = Application.GetOpenFilename(, , "Выбери следующую ревизию")
Set secondBook = Workbooks.Open(Path)

k = 1

Column_name = 3
Column_m = 4
Column_material = 5
Column_o = 7
Column_size = 6
Column_articul = 8
Column_maker = 9
Column_dimension = 10
Column_number = 11
Column_note = 12
Column_Fnumber = 14

Line = 1
lastrow = Cells(firstBook.Sheets("Лист1").Cells.Rows.Count, "C").End(xlUp).row

work_status = "+"
For i = 1 To lastrow
    o_cell = firstBook.Sheets("Лист1").Cells(Line, Column_o) 'обозначение
    m_cell = firstBook.Sheets("Лист1").Cells(Line, Column_m) 'Номер системы
    material_Cell = firstBook.Sheets("Лист1").Cells(Line, Column_material) 'Материал
    Size_cell = firstBook.Sheets("Лист1").Cells(Line, Column_size) 'Размер
    number_cell = firstBook.Sheets("Лист1").Cells(Line, Column_number) 'число
    Fnumber_cell = firstBook.Sheets("Лист1").Cells(Line, Column_Fnumber) 'фактическое число
    name_cell = firstBook.Sheets("Лист1").Cells(Line, Column_name) 'Наименование
    Note_cell = firstBook.Sheets("Лист1").Cells(Line, Column_note) 'Примечание
    Maker_cell = firstBook.Sheets("Лист1").Cells(Line, Column_maker) 'Производитель
    Dimension_cell = firstBook.Sheets("Лист1").Cells(Line, Column_dimension) 'Размерность
    Articul_cell = firstBook.Sheets("Лист1").Cells(Line, Column_articul) 'Артикул
    Find_Changes o_cell, m_cell, name_cell, Size_cell, number_cell, Fnumber_cell, Maker_cell, Note_cell, Dimension_cell, material_Cell, Articul_cell, secondBook
    find_deleted o_cell, m_cell, name_cell, Size_cell, number_cell, Fnumber_cell, Maker_cell, Note_cell, Dimension_cell, material_Cell, Articul_cell, secondBook
    Line = Line + 1
Next

find_new firstBook, secondBook

'secondBook.Close (SaveChanges)
End Sub

Sub Find_Changes(symbol, mark, Name, Size, Number, Fnumber, Maker, Note, dimension, material, articul, book)
Line = 1

lastrow = Cells(Rows.Count, "C").End(xlUp).row

For i = 1 To lastrow
    o_cell = book.Sheets("Лист1").Cells(Line, Column_o) 'обозначение
    m_cell = book.Sheets("Лист1").Cells(Line, Column_m) 'Номер системы
    material_Cell = book.Sheets("Лист1").Cells(Line, Column_material) 'Материал
    number_cell = book.Sheets("Лист1").Cells(Line, Column_number) 'число
    Fnumber_cell = book.Sheets("Лист1").Cells(Line, Column_Fnumber) 'Фактическое число
    name_cell = book.Sheets("Лист1").Cells(Line, Column_name) 'Наименование
    Size_cell = book.Sheets("Лист1").Cells(Line, Column_size) 'Размер
    Note_cell = book.Sheets("Лист1").Cells(Line, Column_note) 'Примечание
    Maker_cell = book.Sheets("Лист1").Cells(Line, Column_maker) 'Примечание
    Dimension_cell = book.Sheets("Лист1").Cells(Line, Column_dimension) 'Размерность
    Articul_cell = book.Sheets("Лист1").Cells(Line, Column_articul) 'Артикул
    
    If symbol = "" And mark = "" And Size = "" And articul = "" And material = "" Then
        If symbol = "" And mark = "" And Size = "" And Name = name_cell Then
            compare_cells Number, number_cell, Line, Column_number
            compare_cells Maker, Maker_cell, Line, Column_maker
            compare_cells Note, Note_cell, Line, Column_note
        End If
    End If
    
    If name_cell Like "*Воздуховод*" Or name_cell Like "*трубка*" Or name_cell Like "*трубы*" Or name_cell Like "*труба*" Or name_cell Like "*воздуховоды*" Then
        If Size = Size_cell And articul = Articul_cell And material = material_Cell And Maker = Maker_cell And Name = name_cell And mark = m_cell And symbol = o_cell Then
            compare_cells Number, number_cell, Line, Column_number
            compare_cells Fnumber, Fnumber_cell, Line, Column_Fnumber
            compare_cells Note, Note_cell, Line, Column_note
            compare_cells dimension, Dimension_cell, Line, Column_dimension
        End If
    Else
        If symbol <> "" Or mark <> "" Or Size <> "" Or articul <> "" Or material <> "" Then
            If symbol = o_cell And mark = m_cell And Size = Size_cell And articul = Articul_cell And material = material_Cell And Maker = Maker_cell Then
                compare_cells Number, number_cell, Line, Column_number
                compare_cells Fnumber, Fnumber_cell, Line, Column_Fnumber
                compare_cells Name, name_cell, Line, Column_name
                compare_cells Note, Note_cell, Line, Column_note
                compare_cells dimension, Dimension_cell, Line, Column_dimension
            End If
        End If
    End If
    Line = Line + 1
Next


End Sub


Sub compare_cells(symbol, mark, Line, Column)
    If symbol <> mark Then
        Cells(Line, Column).Interior.color = vbYellow
    End If
End Sub

Sub find_deleted(symbol, mark, Name, Size, Number, Fnumber, Maker, Note, dimension, material, articul, book)
Line = 1

lastrow = Cells(book.Sheets("Лист1").Cells.Rows.Count, "C").End(xlUp).row
deletedrow = lastrow + 1
deleted = "TRUE"
For i = 1 To lastrow
    o_cell = book.Sheets("Лист1").Cells(Line, Column_o) 'обозначение
    m_cell = book.Sheets("Лист1").Cells(Line, Column_m) 'Номер системы
    number_cell = book.Sheets("Лист1").Cells(Line, Column_number) 'число
    material_Cell = book.Sheets("Лист1").Cells(Line, Column_material) 'Материал
    Fnumber_cell = book.Sheets("Лист1").Cells(Line, Column_Fnumber) 'Фактическое число
    name_cell = book.Sheets("Лист1").Cells(Line, Column_name) 'Наименование
    Size_cell = book.Sheets("Лист1").Cells(Line, Column_size) 'Размер
    Note_cell = book.Sheets("Лист1").Cells(Line, Column_note) 'Примечание
    Maker_cell = book.Sheets("Лист1").Cells(Line, Column_maker) 'Производитель
    Articul_cell = book.Sheets("Лист1").Cells(Line, Column_articul) 'Артикул
    
    
    If symbol = "" And mark = "" And Size = "" And articul = "" And material = "" Then
        If symbol = "" And mark = "" And Size = "" And Name = name_cell Then
            deleted = "FALSE"
        End If
    End If
    
    If name_cell Like "*Воздуховод*" Or name_cell Like "*трубка*" Or name_cell Like "*трубы*" Or name_cell Like "*труба*" Or name_cell Like "*воздуховоды*" Then
        If Size = Size_cell And articul = Articul_cell And material = material_Cell And Maker = Maker_cell And Name = name_cell And mark = m_cell And symbol = o_cell Then
            deleted = "FALSE"
        End If
    Else
        If symbol <> "" Or mark <> "" Or Size <> "" Or articul <> "" Or material <> "" Then
            If symbol = o_cell And mark = m_cell And Size = Size_cell And articul = Articul_cell And material = material_Cell And Maker = Maker_cell Then
                deleted = "FALSE"
                'Debug.Print o_cell
            End If
        End If
    End If
    Line = Line + 1
Next

If deleted = "TRUE" Then
    book.Sheets("Лист1").Cells(deletedrow, Column_o) = symbol 'обозначение
    book.Sheets("Лист1").Cells(deletedrow, Column_m) = mark 'Номер системы
    book.Sheets("Лист1").Cells(deletedrow, Column_number) = Number 'число
    book.Sheets("Лист1").Cells(Line, Column_material) = material 'Материал
    book.Sheets("Лист1").Cells(deletedrow, Column_Fnumber) = Fnumber 'Фактическое число
    book.Sheets("Лист1").Cells(deletedrow, Column_name) = Name 'Наименование
    book.Sheets("Лист1").Cells(deletedrow, Column_size) = Size 'Размер
    book.Sheets("Лист1").Cells(deletedrow, Column_note) = Note 'Примечание
    book.Sheets("Лист1").Cells(deletedrow, Column_maker) = Maker 'Производитель
    book.Sheets("Лист1").Cells(deletedrow, Column_dimension) = dimension 'Размерность
    color deletedrow, vbRed
End If

End Sub


Sub color(row, color)

Cells(row, Column_name).Interior.color = color
Cells(row, Column_m).Interior.color = color
Cells(row, Column_o).Interior.color = color
Cells(row, Column_size).Interior.color = color
Cells(row, Column_maker).Interior.color = color
Cells(row, Column_number).Interior.color = color
Cells(row, Column_note).Interior.color = color
Cells(row, Column_Fnumber).Interior.color = color
Cells(row, Column_dimension).Interior.color = color
Cells(row, Column_dimension).Interior.color = color
Cells(row, Column_material).Interior.color = color
Cells(row, Column_articul).Interior.color = color
Cells(row, 13).Interior.color = color
End Sub

Sub find_new(firstBook, secondBook)
Line = 1
lastrow = Cells(secondBook.Sheets("Лист1").Cells.Rows.Count, "C").End(xlUp).row

work_status = "+"

For i = 1 To lastrow
    o_cell = secondBook.Sheets("Лист1").Cells(Line, Column_o) 'обозначение
    m_cell = secondBook.Sheets("Лист1").Cells(Line, Column_m) 'Номер системы
    material_Cell = secondBook.Sheets("Лист1").Cells(Line, Column_material) 'Материал
    Size_cell = secondBook.Sheets("Лист1").Cells(Line, Column_size) 'Размер
    number_cell = secondBook.Sheets("Лист1").Cells(Line, Column_number) 'число
    Fnumber_cell = secondBook.Sheets("Лист1").Cells(Line, Column_Fnumber) 'фактическое число
    name_cell = secondBook.Sheets("Лист1").Cells(Line, Column_name) 'Наименование
    Note_cell = secondBook.Sheets("Лист1").Cells(Line, Column_note) 'Примечание
    Maker_cell = secondBook.Sheets("Лист1").Cells(Line, Column_maker) 'Производитель
    Dimension_cell = secondBook.Sheets("Лист1").Cells(Line, Column_dimension) 'Размерность
    Articul_cell = secondBook.Sheets("Лист1").Cells(Line, Column_articul) 'Артикул
    color_new o_cell, m_cell, name_cell, Size_cell, number_cell, Fnumber_cell, Maker_cell, Note_cell, Dimension_cell, material_Cell, Articul_cell, firstBook, secondBook, Line
    Line = Line + 1
Next

End Sub

Sub color_new(symbol, mark, Name, Size, Number, Fnumber, Maker, Note, dimension, material, articul, firstBook, secondBook, newline)
Line = 1

lastrow = Cells(firstBook.Sheets("Лист1").Cells.Rows.Count, "C").End(xlUp).row
deletedrow = lastrow + 1
newpos = "TRUE"
For i = 1 To lastrow
    o_cell = firstBook.Sheets("Лист1").Cells(Line, Column_o) 'обозначение
    m_cell = firstBook.Sheets("Лист1").Cells(Line, Column_m) 'Номер системы
    number_cell = firstBook.Sheets("Лист1").Cells(Line, Column_number) 'число
    material_Cell = firstBook.Sheets("Лист1").Cells(Line, Column_material) 'Материал
    Fnumber_cell = firstBook.Sheets("Лист1").Cells(Line, Column_Fnumber) 'Фактическое число
    name_cell = firstBook.Sheets("Лист1").Cells(Line, Column_name) 'Наименование
    Size_cell = firstBook.Sheets("Лист1").Cells(Line, Column_size) 'Размер
    Note_cell = firstBook.Sheets("Лист1").Cells(Line, Column_note) 'Примечание
    Maker_cell = firstBook.Sheets("Лист1").Cells(Line, Column_maker) 'Производитель
    Articul_cell = firstBook.Sheets("Лист1").Cells(Line, Column_articul) 'Артикул
    
    If symbol = "" And mark = "" And Size = "" And articul = "" And material = "" Then
        If symbol = "" And mark = "" And Size = "" And Name = name_cell Then
            newpos = "FALSE"
        End If
    End If
    
    If name_cell Like "*Воздуховод*" Or name_cell Like "*трубка*" Or name_cell Like "*трубы*" Or name_cell Like "*труба*" Or name_cell Like "*воздуховоды*" Then
        If Size = Size_cell And articul = Articul_cell And material = material_Cell And Maker = Maker_cell And Name = name_cell And mark = m_cell And symbol = o_cell Then
            newpos = "FALSE"
        End If
    Else
        If symbol <> "" Or mark <> "" Or Size <> "" Or articul <> "" Or material <> "" Then
            If symbol = o_cell And mark = m_cell And Size = Size_cell And articul = Articul_cell And material = material_Cell And Maker = Maker_cell Then
                newpos = "FALSE"
                'Debug.Print o_cell
            End If
        End If
    End If
    Line = Line + 1
Next

If newpos = "TRUE" Then
    color newline, vbGreen
End If

End Sub



