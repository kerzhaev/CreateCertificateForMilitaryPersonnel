VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearch 
   Caption         =   "Search..."
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






Option Explicit

' Version: 0.7.2

' Updated: 2026-03-09

Private Const MIN_SEARCH_LENGTH As Long = 3

Private Const DISPLAY_SEPARATOR As String = " - "
Private mImportedColumnMap As Object

Private Sub UserForm_Initialize()

    ListBox1.ColumnCount = 2

    ListBox1.ColumnWidths = "0 pt;340 pt"

    Set mImportedColumnMap = GetImportedColumnMap(GetImportedDataWorksheet())

End Sub

Private Sub TextBox1_Change()

    Dim searchValue As String

    searchValue = Trim$(TextBox1.Value)

    If Len(searchValue) >= MIN_SEARCH_LENGTH Then

        PerformSearch searchValue

    Else

        ListBox1.Clear

    End If

End Sub

Private Sub PerformSearch(ByVal searchValue As String)

    Dim ws As Worksheet

    Dim lastRow As Long

    Dim rowIndex As Long

    Set ws = GetImportedDataWorksheet()

    If mImportedColumnMap Is Nothing Then Set mImportedColumnMap = GetImportedColumnMap(ws)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ListBox1.Clear

    For rowIndex = 2 To lastRow

        If RowMatchesSearch(ws, rowIndex, searchValue) Then AddSearchItem rowIndex, BuildDisplayValue(ws, rowIndex)

    Next rowIndex

End Sub

Private Sub AddSearchItem(ByVal sourceRow As Long, ByVal displayValue As String)

    Dim itemIndex As Long

    ListBox1.AddItem CStr(sourceRow)

    itemIndex = ListBox1.ListCount - 1

    ListBox1.List(itemIndex, 1) = displayValue

End Sub

Private Function BuildDisplayValue(ByVal ws As Worksheet, ByVal rowIndex As Long) As String

    BuildDisplayValue = JoinDisplayValues(Array(GetImportedCellText(ws, rowIndex, "personal_number"), GetImportedFullNameText(ws, rowIndex), GetImportedBirthDateText(ws, rowIndex), GetImportedCellText(ws, rowIndex, "military_unit")))

End Function

Private Function JoinDisplayValues(ByVal values As Variant) As String

    Dim item As Variant

    Dim normalizedValue As String

    For Each item In values

        normalizedValue = Trim$(CStr(item))

        If Len(normalizedValue) > 0 Then

            If Len(JoinDisplayValues) > 0 Then

                JoinDisplayValues = JoinDisplayValues & DISPLAY_SEPARATOR

            End If

            JoinDisplayValues = JoinDisplayValues & normalizedValue

        End If

    Next item

End Function

Private Sub ListBox1_Click()

    Dim sourceRow As Long

    On Error GoTo HandleError

    If ListBox1.listIndex < 0 Then Exit Sub

    sourceRow = CLng(ListBox1.List(ListBox1.listIndex, 0))

    CopySelectionToDataSheet sourceRow

    Unload Me

    Exit Sub

HandleError:

    MsgBox "Unable to copy the selected row: " & Err.Description, vbCritical, "Search error"

End Sub

Private Sub CopySelectionToDataSheet(ByVal sourceRow As Long)

    Dim sourceSheet As Worksheet

    Dim targetSheet As Worksheet

    Dim targetRow As Long

    Set sourceSheet = GetImportedDataWorksheet()

    If mImportedColumnMap Is Nothing Then Set mImportedColumnMap = GetImportedColumnMap(sourceSheet)

    Set targetSheet = GetDataWorksheet()

    targetRow = GetActiveDataRow(targetSheet)

    With targetSheet

        .Cells(targetRow, 4).Value = GetImportedFullNameText(sourceSheet, sourceRow)

        .Cells(targetRow, 5).Value = GetImportedCellText(sourceSheet, sourceRow, "personal_number")

        .Cells(targetRow, 6).Value = GetImportedBirthDateText(sourceSheet, sourceRow)

        .Cells(targetRow, 7).Value = NormalizeUnitValue(GetImportedCellText(sourceSheet, sourceRow, "military_unit"))

    End With

End Sub

Private Function RowMatchesSearch(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal searchValue As String) As Boolean
    If InStr(1, GetImportedCellText(ws, rowIndex, "personal_number"), searchValue, vbTextCompare) > 0 Then RowMatchesSearch = True: Exit Function
    If InStr(1, GetImportedFullNameText(ws, rowIndex), searchValue, vbTextCompare) > 0 Then RowMatchesSearch = True: Exit Function
    If InStr(1, GetImportedBirthDateText(ws, rowIndex), searchValue, vbTextCompare) > 0 Then RowMatchesSearch = True: Exit Function
    If InStr(1, GetImportedCellText(ws, rowIndex, "military_unit"), searchValue, vbTextCompare) > 0 Then RowMatchesSearch = True
End Function

Private Function GetImportedFullNameText(ByVal ws As Worksheet, ByVal rowIndex As Long) As String
    Dim fullNameColumn As Long
    Dim surnameColumn As Long
    Dim givenNameColumn As Long
    Dim patronymicColumn As Long
    If mImportedColumnMap Is Nothing Then Set mImportedColumnMap = GetImportedColumnMap(ws)
    fullNameColumn = CLng(mImportedColumnMap("full_name"))
    If fullNameColumn > 0 Then
        GetImportedFullNameText = Trim$(CStr(ws.Cells(rowIndex, fullNameColumn).Value))
        Exit Function
    End If
    surnameColumn = CLng(mImportedColumnMap("surname"))
    givenNameColumn = CLng(mImportedColumnMap("given_name"))
    patronymicColumn = CLng(mImportedColumnMap("patronymic"))
    GetImportedFullNameText = Trim$(CStr(ws.Cells(rowIndex, surnameColumn).Value) & " " & CStr(ws.Cells(rowIndex, givenNameColumn).Value) & " " & CStr(ws.Cells(rowIndex, patronymicColumn).Value))
End Function

Private Function GetImportedBirthDateText(ByVal ws As Worksheet, ByVal rowIndex As Long) As String
    Dim birthDateColumn As Long
    Dim sourceValue As Variant
    If mImportedColumnMap Is Nothing Then Set mImportedColumnMap = GetImportedColumnMap(ws)
    birthDateColumn = CLng(mImportedColumnMap("birth_date"))
    If birthDateColumn <= 0 Then Exit Function
    sourceValue = ws.Cells(rowIndex, birthDateColumn).Value
    If IsEmpty(sourceValue) Or IsNull(sourceValue) Then Exit Function
    If IsDate(sourceValue) Then
        GetImportedBirthDateText = Format$(CDate(sourceValue), "dd.mm.yyyy")
    Else
        GetImportedBirthDateText = Trim$(CStr(sourceValue))
    End If
End Function

Private Function GetImportedCellText(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal mapKey As String) As String
    Dim columnIndex As Long
    If mImportedColumnMap Is Nothing Then Set mImportedColumnMap = GetImportedColumnMap(ws)
    columnIndex = CLng(mImportedColumnMap(mapKey))
    If columnIndex <= 0 Then Exit Function
    GetImportedCellText = Trim$(CStr(ws.Cells(rowIndex, columnIndex).Value))
End Function

Private Function GetActiveDataRow(ByVal targetSheet As Worksheet) As Long

    If TypeName(Selection) <> "Range" Then

        Err.Raise vbObjectError + 2000, "GetActiveDataRow", "Select a target row on the 'data' sheet first."

    End If

    If Not Application.ActiveCell.Parent Is targetSheet Then

        Err.Raise vbObjectError + 2001, "GetActiveDataRow", "Select a target row on the 'data' sheet first."

    End If

    If Application.ActiveCell.Row < 2 Then

        Err.Raise vbObjectError + 2002, "GetActiveDataRow", "Select a data row below the header."

    End If

    GetActiveDataRow = Application.ActiveCell.Row

End Function

