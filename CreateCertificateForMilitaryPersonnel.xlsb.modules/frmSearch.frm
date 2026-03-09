VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearch 
   Caption         =   "Поиск ..."
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8475.001
   OleObjectBlob   =   "frmSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox1_Change()
    Dim searchValue As String
    searchValue = TextBox1.Value

    ' Активировать поиск, если введено три или более символов
    If Len(searchValue) >= 3 Then
        PerformSearch searchValue
    Else
        ListBox1.Clear ' Очистить результаты поиска, если символов меньше трех
    End If
End Sub

'Private Sub PerformSearch(ByVal searchValue As String)
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets("Выгрузка")
'    Dim lastRow As Long
'    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
'
'    ListBox1.Clear ' Очистка ListBox перед поиском
'
'    Dim i As Long
'    For i = 1 To lastRow
'        If LCase(Trim(ws.Cells(i, 3).Value)) Like "*" & LCase(Trim(searchValue)) & "*" Then
'            ListBox1.AddItem ws.Cells(i, 1).Value & " - " & ws.Cells(i, 3).Value & " - " & ws.Cells(i, 4).Value
'        End If
'    Next i
'End Sub



'Private Sub PerformSearch(ByVal searchValue As String)
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets("Выгрузка")
'
'    Dim lastRow As Long
'    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
'
'    ListBox1.Clear ' Очищаем список перед новым поиском
'
'    Dim i As Long, j As Long
'    For i = 2 To lastRow ' Предполагаем, что первая строка содержит заголовки
'        For j = 1 To ws.Columns.Count ' Перебираем все столбцы
'            If InStr(1, ws.Cells(i, j).Value, searchValue, vbTextCompare) > 0 Then
'                ' Добавление результатов в ListBox, включая значение из 21-го столбца
'                ListBox1.AddItem ws.Cells(i, 2).Value & " - " & ws.Cells(i, 3).Value & " - " & ws.Cells(i, 4).Value & " - " & ws.Cells(i, 21).Value
'                Exit For ' Прекращаем поиск в текущей строке после первого совпадения
'            End If
'        Next j
'    Next i
'End Sub


Private Sub PerformSearch(ByVal searchValue As String)
    ' Запускаем поиск только если в строке поиска есть минимум 3 символа
    If Len(searchValue) < 3 Then
        ListBox1.Clear
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Выгрузка")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ListBox1.Clear ' Очищаем список перед новым поиском

    Dim i As Long, j As Long
    For i = 2 To lastRow ' Предполагаем, что первая строка содержит заголовки
        For j = 2 To 5 ' Ограничиваем поиск столбцами 2, 3, 4 и 5
            If InStr(1, ws.Cells(i, j).Value, searchValue, vbTextCompare) > 0 Then
                ' Добавляем результаты в ListBox
                ListBox1.AddItem ws.Cells(i, 2).Value & " - " & ws.Cells(i, 3).Value & " - " & ws.Cells(i, 4).Value & " - " & ws.Cells(i, 5).Value & " - " & ws.Cells(i, 21).Value
                Exit For ' Прекращаем поиск в текущей строке после первого совпадения
            End If
        Next j
    Next i
End Sub













Private Sub ListBox1_Click()
    Dim selectedValue As String
    selectedValue = ListBox1.Value

    ' Извлекаем только значение из третьего столбца из выбранной строки
    Dim valueParts As Variant
    valueParts = Split(selectedValue, " - ")
    If UBound(valueParts) >= 2 Then
        selectedValue = valueParts(1) ' значение из третьего столбца
    Else
        Exit Sub ' Если формат строки не соответствует ожидаемому, выходим из процедуры
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Выгрузка")

    ' Находим строку в листе "Выгрузка", соответствующую выбранному значению
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row

    For i = 1 To lastRow
        If ws.Cells(i, 3).Value = valueParts(1) Then
            ' Вставляем значения в лист "data"
            Dim targetSheet As Worksheet
            Set targetSheet = ThisWorkbook.Sheets("data")
            With targetSheet
                .Cells(Application.ActiveCell.Row, 4).Value = selectedValue ' Вставляем выбранное значение в 4-й столбец
                .Cells(Application.ActiveCell.Row, 5).Value = ws.Cells(i, 2).Value ' значение из второго столбца
                .Cells(Application.ActiveCell.Row, 6).Value = ws.Cells(i, 4).Value ' значение из четвертого столбца
                .Cells(Application.ActiveCell.Row, 7).Value = "'" & ExtractNumbers(ws.Cells(i, 21).Value) ' значение из двадцать первого столбца
            End With
            Exit For
        End If
    Next i

    Unload Me ' Закрытие формы
End Sub




