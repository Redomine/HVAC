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
Path = Application.GetOpenFilename(, , "������ ��������� �������")
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
lastrow = Cells(firstBook.Sheets("����1").Cells.Rows.Count, "C").End(xlUp).row

work_status = "+"
For i = 1 To lastrow
    o_cell = firstBook.Sheets("����1").Cells(Line, Column_o) '�����������
    m_cell = firstBook.Sheets("����1").Cells(Line, Column_m) '����� �������
    material_Cell = firstBook.Sheets("����1").Cells(Line, Column_material) '��������
    Size_cell = firstBook.Sheets("����1").Cells(Line, Column_size) '������
    number_cell = firstBook.Sheets("����1").Cells(Line, Column_number) '�����
    Fnumber_cell = firstBook.Sheets("����1").Cells(Line, Column_Fnumber) '����������� �����
    name_cell = firstBook.Sheets("����1").Cells(Line, Column_name) '������������
    Note_cell = firstBook.Sheets("����1").Cells(Line, Column_note) '����������
    Maker_cell = firstBook.Sheets("����1").Cells(Line, Column_maker) '�������������
    Dimension_cell = firstBook.Sheets("����1").Cells(Line, Column_dimension) '�����������
    Articul_cell = firstBook.Sheets("����1").Cells(Line, Column_articul) '�������
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
    o_cell = book.Sheets("����1").Cells(Line, Column_o) '�����������
    m_cell = book.Sheets("����1").Cells(Line, Column_m) '����� �������
    material_Cell = book.Sheets("����1").Cells(Line, Column_material) '��������
    number_cell = book.Sheets("����1").Cells(Line, Column_number) '�����
    Fnumber_cell = book.Sheets("����1").Cells(Line, Column_Fnumber) '����������� �����
    name_cell = book.Sheets("����1").Cells(Line, Column_name) '������������
    Size_cell = book.Sheets("����1").Cells(Line, Column_size) '������
    Note_cell = book.Sheets("����1").Cells(Line, Column_note) '����������
    Maker_cell = book.Sheets("����1").Cells(Line, Column_maker) '����������
    Dimension_cell = book.Sheets("����1").Cells(Line, Column_dimension) '�����������
    Articul_cell = book.Sheets("����1").Cells(Line, Column_articul) '�������
    
    If symbol = "" And mark = "" And Size = "" And articul = "" And material = "" Then
        If symbol = "" And mark = "" And Size = "" And Name = name_cell Then
            compare_cells Number, number_cell, Line, Column_number
            compare_cells Maker, Maker_cell, Line, Column_maker
            compare_cells Note, Note_cell, Line, Column_note
        End If
    End If
    
    If name_cell Like "*����������*" Or name_cell Like "*������*" Or name_cell Like "*�����*" Or name_cell Like "*�����*" Or name_cell Like "*�����������*" Then
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

lastrow = Cells(book.Sheets("����1").Cells.Rows.Count, "C").End(xlUp).row
deletedrow = lastrow + 1
deleted = "TRUE"
For i = 1 To lastrow
    o_cell = book.Sheets("����1").Cells(Line, Column_o) '�����������
    m_cell = book.Sheets("����1").Cells(Line, Column_m) '����� �������
    number_cell = book.Sheets("����1").Cells(Line, Column_number) '�����
    material_Cell = book.Sheets("����1").Cells(Line, Column_material) '��������
    Fnumber_cell = book.Sheets("����1").Cells(Line, Column_Fnumber) '����������� �����
    name_cell = book.Sheets("����1").Cells(Line, Column_name) '������������
    Size_cell = book.Sheets("����1").Cells(Line, Column_size) '������
    Note_cell = book.Sheets("����1").Cells(Line, Column_note) '����������
    Maker_cell = book.Sheets("����1").Cells(Line, Column_maker) '�������������
    Articul_cell = book.Sheets("����1").Cells(Line, Column_articul) '�������
    
    
    If symbol = "" And mark = "" And Size = "" And articul = "" And material = "" Then
        If symbol = "" And mark = "" And Size = "" And Name = name_cell Then
            deleted = "FALSE"
        End If
    End If
    
    If name_cell Like "*����������*" Or name_cell Like "*������*" Or name_cell Like "*�����*" Or name_cell Like "*�����*" Or name_cell Like "*�����������*" Then
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
    book.Sheets("����1").Cells(deletedrow, Column_o) = symbol '�����������
    book.Sheets("����1").Cells(deletedrow, Column_m) = mark '����� �������
    book.Sheets("����1").Cells(deletedrow, Column_number) = Number '�����
    book.Sheets("����1").Cells(Line, Column_material) = material '��������
    book.Sheets("����1").Cells(deletedrow, Column_Fnumber) = Fnumber '����������� �����
    book.Sheets("����1").Cells(deletedrow, Column_name) = Name '������������
    book.Sheets("����1").Cells(deletedrow, Column_size) = Size '������
    book.Sheets("����1").Cells(deletedrow, Column_note) = Note '����������
    book.Sheets("����1").Cells(deletedrow, Column_maker) = Maker '�������������
    book.Sheets("����1").Cells(deletedrow, Column_dimension) = dimension '�����������
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
lastrow = Cells(secondBook.Sheets("����1").Cells.Rows.Count, "C").End(xlUp).row

work_status = "+"

For i = 1 To lastrow
    o_cell = secondBook.Sheets("����1").Cells(Line, Column_o) '�����������
    m_cell = secondBook.Sheets("����1").Cells(Line, Column_m) '����� �������
    material_Cell = secondBook.Sheets("����1").Cells(Line, Column_material) '��������
    Size_cell = secondBook.Sheets("����1").Cells(Line, Column_size) '������
    number_cell = secondBook.Sheets("����1").Cells(Line, Column_number) '�����
    Fnumber_cell = secondBook.Sheets("����1").Cells(Line, Column_Fnumber) '����������� �����
    name_cell = secondBook.Sheets("����1").Cells(Line, Column_name) '������������
    Note_cell = secondBook.Sheets("����1").Cells(Line, Column_note) '����������
    Maker_cell = secondBook.Sheets("����1").Cells(Line, Column_maker) '�������������
    Dimension_cell = secondBook.Sheets("����1").Cells(Line, Column_dimension) '�����������
    Articul_cell = secondBook.Sheets("����1").Cells(Line, Column_articul) '�������
    color_new o_cell, m_cell, name_cell, Size_cell, number_cell, Fnumber_cell, Maker_cell, Note_cell, Dimension_cell, material_Cell, Articul_cell, firstBook, secondBook, Line
    Line = Line + 1
Next

End Sub

Sub color_new(symbol, mark, Name, Size, Number, Fnumber, Maker, Note, dimension, material, articul, firstBook, secondBook, newline)
Line = 1

lastrow = Cells(firstBook.Sheets("����1").Cells.Rows.Count, "C").End(xlUp).row
deletedrow = lastrow + 1
newpos = "TRUE"
For i = 1 To lastrow
    o_cell = firstBook.Sheets("����1").Cells(Line, Column_o) '�����������
    m_cell = firstBook.Sheets("����1").Cells(Line, Column_m) '����� �������
    number_cell = firstBook.Sheets("����1").Cells(Line, Column_number) '�����
    material_Cell = firstBook.Sheets("����1").Cells(Line, Column_material) '��������
    Fnumber_cell = firstBook.Sheets("����1").Cells(Line, Column_Fnumber) '����������� �����
    name_cell = firstBook.Sheets("����1").Cells(Line, Column_name) '������������
    Size_cell = firstBook.Sheets("����1").Cells(Line, Column_size) '������
    Note_cell = firstBook.Sheets("����1").Cells(Line, Column_note) '����������
    Maker_cell = firstBook.Sheets("����1").Cells(Line, Column_maker) '�������������
    Articul_cell = firstBook.Sheets("����1").Cells(Line, Column_articul) '�������
    
    If symbol = "" And mark = "" And Size = "" And articul = "" And material = "" Then
        If symbol = "" And mark = "" And Size = "" And Name = name_cell Then
            newpos = "FALSE"
        End If
    End If
    
    If name_cell Like "*����������*" Or name_cell Like "*������*" Or name_cell Like "*�����*" Or name_cell Like "*�����*" Or name_cell Like "*�����������*" Then
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



