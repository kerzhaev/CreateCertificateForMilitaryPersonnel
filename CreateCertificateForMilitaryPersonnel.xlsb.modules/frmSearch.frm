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
' Version: 0.4.1
' Updated: 2026-03-09

Private Const MIN_SEARCH_LENGTH As Long = 3
Private Const DISPLAY_SEPARATOR As String = " - "

Private Sub UserForm_Initialize()
    ListBox1.ColumnCount = 2
    ListBox1.ColumnWidths = "0 pt;340 pt"
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
    Dim columnIndex As Long

    Set ws = GetImportedDataWorksheet()
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ListBox1.Clear

    For rowIndex = 2 To lastRow
        For columnIndex = 2 To 5
            If InStr(1, CStr(ws.Cells(rowIndex, columnIndex).Value), searchValue, vbTextCompare) > 0 Then
                AddSearchItem rowIndex, BuildDisplayValue(ws, rowIndex)
                Exit For
            End If
        Next columnIndex
    Next rowIndex
End Sub

Private Sub AddSearchItem(ByVal sourceRow As Long, ByVal displayValue As String)
    Dim itemIndex As Long

    ListBox1.AddItem CStr(sourceRow)
    itemIndex = ListBox1.ListCount - 1
    ListBox1.List(itemIndex, 1) = displayValue
End Sub

Private Function BuildDisplayValue(ByVal ws As Worksheet, ByVal rowIndex As Long) As String
    BuildDisplayValue = CStr(ws.Cells(rowIndex, 2).Value) & DISPLAY_SEPARATOR & _
        CStr(ws.Cells(rowIndex, 3).Value) & DISPLAY_SEPARATOR & _
        CStr(ws.Cells(rowIndex, 4).Value) & DISPLAY_SEPARATOR & _
        CStr(ws.Cells(rowIndex, 5).Value) & DISPLAY_SEPARATOR & _
        CStr(ws.Cells(rowIndex, 21).Value)
End Function

Private Sub ListBox1_Click()
    Dim sourceRow As Long

    On Error GoTo HandleError

    If ListBox1.ListIndex < 0 Then Exit Sub

    sourceRow = CLng(ListBox1.List(ListBox1.ListIndex, 0))
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
    Set targetSheet = GetDataWorksheet()
    targetRow = GetActiveDataRow(targetSheet)

    With targetSheet
        .Cells(targetRow, 4).Value = CStr(sourceSheet.Cells(sourceRow, 3).Value)
        .Cells(targetRow, 5).Value = CStr(sourceSheet.Cells(sourceRow, 2).Value)
        .Cells(targetRow, 6).Value = CStr(sourceSheet.Cells(sourceRow, 4).Value)
        .Cells(targetRow, 7).Value = NormalizeUnitValue(CStr(sourceSheet.Cells(sourceRow, 21).Value))
    End With
End Sub

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
