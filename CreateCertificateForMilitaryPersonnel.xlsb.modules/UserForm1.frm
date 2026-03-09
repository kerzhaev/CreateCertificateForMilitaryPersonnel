VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Список шаблонов (*.docx)"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







''''''''''''''''''''''''''''''''''''''

'Разработка макроса: excelstore.pro

'E-mail для связи: info@excelstore.pro

''''''''''''''''''''''''''''''''''''''

Option Explicit

' Version: 0.4.1

' Updated: 2026-03-09

Private Sub UserForm_Activate()

    Dim iArray, i As Integer

    iArray = Split(Range("FILE_TEMPLATE").Value, ";", , vbTextCompare)

    For i = 0 To UBound(iArray)

        ListBox1.AddItem (iArray(i))

    Next i

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim i As Integer

    For i = 0 To ListBox1.ListCount - 1

        If ListBox1.Selected(i) = True Then ActiveCell.Value = ListBox1.List(i)

    Next i

    Unload Me

End Sub

Private Sub cbEnter_Click()

    Dim iSTR As String, i As Integer

    For i = 0 To ListBox1.ListCount - 1

        If ListBox1.Selected(i) = True Then iSTR = iSTR & ListBox1.List(i) & ";"

    Next i

    If iSTR = "" Then

        ActiveCell.Value = ""

    Else

        ActiveCell.Value = Left(iSTR, Len(iSTR) - 1)

    End If

    Unload Me

End Sub

Private Sub cbCancel_Click()

    Unload Me

End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyEscape Then Unload Me

End Sub
