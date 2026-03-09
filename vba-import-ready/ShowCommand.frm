VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ShowCommand 
   Caption         =   "Выбор команд"
   ClientHeight    =   1500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3930
   OleObjectBlob   =   "ShowCommand.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ShowCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnImportData_Click()
    CreateAndImportDataSheet
    Unload Me
End Sub

Private Sub btnSaveData_Click()
    CreateDoc
    
    Unload Me
End Sub




Private Sub UserForm_Initialize()

    ' Центрирование формы на экране
    With Me
        .StartUpPosition = 0 ' Используем ручное позиционирование
        .Left = Application.Left + (Application.Width - .Width) / 2
        .Top = Application.Top + (Application.Height - .Height) / 2
    End With


End Sub
