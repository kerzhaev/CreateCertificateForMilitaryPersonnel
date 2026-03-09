VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Select Templates"
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

Option Explicit
' Version: 0.6.0
' Updated: 2026-03-09

Private Const TEMPLATE_MASK As String = "*.docx"

Private Sub UserForm_Initialize()
    Me.Caption = "Select Templates"
    cbEnter.Caption = "Save"
    cbCancel.Caption = "Cancel"
    ListBox1.Clear
    ListBox1.MultiSelect = fmMultiSelectMulti

    LoadTemplateCatalog
End Sub

Private Sub LoadTemplateCatalog()
    Dim templateFolder As String
    Dim storedSelection As String
    Dim templateFileName As String
    Dim listIndex As Long

    On Error GoTo HandleError

    templateFolder = GetTemplateFolderSetting()
    storedSelection = GetTemplateCatalogSetting()
    templateFileName = dir$(templateFolder & TEMPLATE_MASK)

    Do While Len(templateFileName) > 0
        ListBox1.AddItem NormalizeTemplateEntry(templateFileName)
        templateFileName = dir$
    Loop

    If ListBox1.ListCount = 0 Then
        MsgBox "No Word templates (*.docx) were found in the selected template folder.", vbExclamation, "Templates"
        Exit Sub
    End If

    For listIndex = 0 To ListBox1.ListCount - 1
        If IsSelectedTemplate(CStr(ListBox1.List(listIndex)), storedSelection) Then
            ListBox1.Selected(listIndex) = True
        End If
    Next listIndex

    Exit Sub

HandleError:
    MsgBox "Unable to load templates: " & Err.Description, vbExclamation, "Templates"
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SaveSelection
End Sub

Private Sub cbEnter_Click()
    SaveSelection
End Sub

Private Sub SaveSelection()
    Dim selectedTemplates As String
    Dim listIndex As Long

    For listIndex = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(listIndex) Then
            If Len(selectedTemplates) > 0 Then
                selectedTemplates = selectedTemplates & ";"
            End If
            selectedTemplates = selectedTemplates & NormalizeTemplateEntry(CStr(ListBox1.List(listIndex)))
        End If
    Next listIndex

    SaveTemplateCatalogSetting selectedTemplates
    MsgBox "Template list saved.", vbInformation, "Templates"
    Unload Me
End Sub

Private Function NormalizeTemplateEntry(ByVal templateName As String) As String
    templateName = Trim$(templateName)

    If LCase$(Right$(templateName, 5)) = ".docx" Then
        templateName = Left$(templateName, Len(templateName) - 5)
    End If

    NormalizeTemplateEntry = templateName
End Function

Private Function IsSelectedTemplate(ByVal templateName As String, ByVal storedSelection As String) As Boolean
    Dim normalizedSelection As String
    Dim normalizedTemplate As String

    normalizedTemplate = LCase$(NormalizeTemplateEntry(templateName))

    If Len(normalizedTemplate) = 0 Then Exit Function
    If Len(Trim$(storedSelection)) = 0 Then Exit Function

    normalizedSelection = ";" & LCase$(storedSelection) & ";"
    IsSelectedTemplate = InStr(1, normalizedSelection, ";" & normalizedTemplate & ";", vbTextCompare) > 0
End Function

Private Sub cbCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
