Attribute VB_Name = "iMacro"

Option Explicit

' Version: 0.5.2

' Updated: 2026-03-09

Private Const WORD_XML_FORMAT As Long = 12

Private Const WD_FIND_STOP As Long = 0

Private Const WD_COLLAPSE_END As Long = 0

Private Const DATA_REPLACEMENT_START_COLUMN As Long = 4

Private Const UNIT_REPLACEMENT_COLUMN As Long = 7

Private Const MIN_REQUIRED_DATA_COLUMNS As Long = 4

Private Const FILE_WORD_RANGE_NAME As String = "FILE_WORD"

Private Const FILE_TEMPLATE_RANGE_NAME As String = "FILE_TEMPLATE"

Private Const OUTPUT_FOLDER_SETTING_NAME As String = "CERTIFICATE_OUTPUT_FOLDER"

Private Const DATA_SHEET_NAME As String = "data"

Private Const IMPORTED_DATA_SHEET_NAME As String = "ImportedData"

Private Const HISTORY_SHEET_NAME As String = "IssuedDocumentsLog"

Private Const IMPORTED_DBF_SHEET_NAME As String = "ImportedDbfData"

Private Const LEGACY_IMPORTED_DATA_SHEET_NAME As String = "Выгрузка"

Private Const LEGACY_HISTORY_SHEET_NAME As String = "БазаВыдСпр"

Private Const SUMMARY_FILE_PREFIX As String = "Certificates_"

Private Const HISTORY_HEADER_ROW As Long = 1

Public AppWord As Object

Public iWord As Object

Public Sub ShowPopup()
    ShowCommand.Show
End Sub

Public Sub OpenSearchForm()
    Dim dataSheet As Worksheet

    Set dataSheet = GetDataWorksheet()

    If Not Application.ActiveSheet Is dataSheet Then
        dataSheet.Activate
    End If

    If TypeName(Selection) <> "Range" Or Application.ActiveCell.Row < 2 Then
        dataSheet.Cells(2, 4).Select
    End If

    Load frmSearch

    With frmSearch
        .StartUpPosition = 0
        .Top = Application.Top + (Application.Height / 2 - .Height / 2)
        .Left = Application.Left + (Application.Width / 2 - .Width / 2)
        .Show
    End With
End Sub

Public Sub OpenHistorySheet()

    GetHistoryWorksheet.Activate

End Sub
Public Sub SelectTemplateFolder()
    Dim currentFolder As String
    Dim selectedFolder As String

    currentFolder = GetConfiguredTextSetting(FILE_WORD_RANGE_NAME)

    If Len(currentFolder) = 0 Or dir$(currentFolder, vbDirectory) = vbNullString Then
        currentFolder = GetDefaultBaseFolder()
    End If

    selectedFolder = PickFolderPath("Select the template folder", currentFolder)
    If Len(selectedFolder) = 0 Then Exit Sub

    SaveStoredTextSetting FILE_WORD_RANGE_NAME, selectedFolder
End Sub

Public Sub SelectCertificateOutputFolder()
    Dim currentFolder As String
    Dim selectedFolder As String

    currentFolder = GetStoredTextSetting(OUTPUT_FOLDER_SETTING_NAME)

    If Len(currentFolder) = 0 Or dir$(currentFolder, vbDirectory) = vbNullString Then
        currentFolder = GetDefaultBaseFolder()
    End If

    selectedFolder = PickFolderPath("Select a folder for generated certificates", currentFolder)
    If Len(selectedFolder) = 0 Then Exit Sub

    SaveStoredTextSetting OUTPUT_FOLDER_SETTING_NAME, selectedFolder
    MsgBox "Output folder saved:" & vbCrLf & selectedFolder, vbInformation, "Certificates"
End Sub

Public Sub ShowAboutDialog()
    frmAbout.Show
End Sub

Public Sub CreateDoc()

    Dim dataSheet As Worksheet

    Dim dataArray As Variant

    Dim templateFolder As String

    Dim outputFolder As String

    Dim warnings As String

    Dim generatedCount As Long

    Dim loggedCount As Long

    Dim processedRows As Object

    Dim summaryPath As String

    Dim successMessage As String

    Application.ScreenUpdating = False

    On Error GoTo HandleError

    Set dataSheet = GetDataWorksheet()

    ValidateTemplatePickerRange

    templateFolder = GetFolderPath(FILE_WORD_RANGE_NAME)

    outputFolder = GetCertificateOutputFolder()

    dataArray = ReadExcelData(dataSheet)

    Set processedRows = CreateObject("Scripting.Dictionary")

    Set AppWord = CreateObject("Word.Application")

    AppWord.Visible = False

    generatedCount = ProcessArray(dataArray, templateFolder, outputFolder, processedRows, warnings)

    loggedCount = SaveDataToHistorySheet(dataSheet, processedRows)

    summaryPath = ExportSummaryWorkbook(dataSheet)

    successMessage = CStr(generatedCount) & " document(s) created in:" & vbCrLf & outputFolder

    If Len(summaryPath) > 0 Then

        successMessage = successMessage & vbCrLf & "Snapshot workbook saved to:" & vbCrLf & summaryPath

    End If

    If loggedCount > 0 Then

        successMessage = successMessage & vbCrLf & CStr(loggedCount) & " record(s) added to the history sheet."

    End If

    If Len(warnings) > 0 Then

        successMessage = successMessage & vbCrLf & vbCrLf & "Warnings:" & vbCrLf & warnings

    End If

    MsgBox successMessage, vbInformation, "Generation completed"

    GoTo CleanExit

HandleError:

    MsgBox "Document generation failed: " & Err.Description, vbCritical, "Generation error"

CleanExit:

    CloseCurrentDocument

    CleanUpWordApplication

    Application.ScreenUpdating = True

End Sub

Public Sub CreateAndImportDataSheet()

    Dim sourceWorkbook As Workbook

    Dim sourceSheet As Worksheet

    Dim dataSheet As Worksheet

    Dim fileToOpen As Variant

    Application.ScreenUpdating = False

    On Error GoTo HandleError

    fileToOpen = Application.GetOpenFilename( _
        "Excel Files (*.xls;*.xlsx;*.xlsm;*.xlsb), *.xls;*.xlsx;*.xlsm;*.xlsb", _
        , _
        "Select a workbook to import")

    If fileToOpen = False Then GoTo CleanExit

    Set sourceWorkbook = Workbooks.Open(CStr(fileToOpen))

    Set sourceSheet = sourceWorkbook.Worksheets(1)

    Set dataSheet = GetImportedDataWorksheet()

    dataSheet.Cells.Clear

    sourceSheet.UsedRange.Copy Destination:=dataSheet.Range("A1")

    MsgBox "Data imported to the '" & dataSheet.Name & "' worksheet.", vbInformation, "Import completed"

    GoTo CleanExit

HandleError:

    MsgBox "Data import failed: " & Err.Description, vbCritical, "Import error"

CleanExit:

    If Not sourceWorkbook Is Nothing Then

        sourceWorkbook.Close SaveChanges:=False

    End If

    Application.ScreenUpdating = True

End Sub

Public Sub ImportDBF()

    Dim dbfFilePath As Variant

    Dim dbfDirectory As String

    Dim dbfFileName As String

    Dim conn As Object

    Dim rs As Object

    Dim targetSheet As Worksheet

    On Error GoTo HandleError

    dbfFilePath = Application.GetOpenFilename("DBF Files (*.dbf), *.dbf", , "Select a DBF file to import")

    If dbfFilePath = False Then Exit Sub

    dbfDirectory = Left$(CStr(dbfFilePath), InStrRev(CStr(dbfFilePath), "\"))

    dbfFileName = Mid$(CStr(dbfFilePath), InStrRev(CStr(dbfFilePath), "\") + 1)

    Set conn = CreateObject("ADODB.Connection")

    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbfDirectory & ";Extended Properties=dBASE IV;"

    Set rs = CreateObject("ADODB.Recordset")

    rs.Open "SELECT * FROM [" & dbfFileName & "]", conn, 3, 3

    Set targetSheet = ResolveWorksheetByName(IMPORTED_DBF_SHEET_NAME, vbNullString, True)

    targetSheet.Cells.Clear

    targetSheet.Range("A1").CopyFromRecordset rs

    MsgBox "DBF data imported to the '" & targetSheet.Name & "' worksheet.", vbInformation, "Import completed"

    GoTo CleanExit

HandleError:

    MsgBox "DBF import failed: " & Err.Description, vbCritical, "Import error"

CleanExit:

    On Error Resume Next

    If Not rs Is Nothing Then

        If rs.State <> 0 Then rs.Close

    End If

    If Not conn Is Nothing Then

        If conn.State <> 0 Then conn.Close

    End If

    Set rs = Nothing

    Set conn = Nothing

    On Error GoTo 0

End Sub

Public Function GetDataWorksheet() As Worksheet

    Set GetDataWorksheet = ResolveWorksheetByName(DATA_SHEET_NAME, vbNullString, False)

End Function

Public Function GetImportedDataWorksheet() As Worksheet

    Set GetImportedDataWorksheet = ResolveWorksheetByName(IMPORTED_DATA_SHEET_NAME, LEGACY_IMPORTED_DATA_SHEET_NAME, True)

End Function

Public Function GetHistoryWorksheet() As Worksheet

    Set GetHistoryWorksheet = ResolveWorksheetByName(HISTORY_SHEET_NAME, LEGACY_HISTORY_SHEET_NAME, True)

End Function

Public Function NormalizeUnitValue(ByVal sourceValue As String) As String

    Dim normalizedValue As String

    normalizedValue = Replace$(sourceValue, vbTab, " ")

    normalizedValue = Replace$(normalizedValue, Chr$(160), " ")

    normalizedValue = Trim$(Application.Trim(normalizedValue))

    NormalizeUnitValue = normalizedValue

End Function

Private Function ProcessArray(ByVal dataArray As Variant, ByVal templateFolder As String, ByVal outputFolder As String, ByVal processedRows As Object, ByRef warnings As String) As Long

    Dim lastRow As Long

    Dim lastCol As Long

    Dim rowIndex As Long

    Dim columnIndex As Long

    Dim templateItems As Variant

    Dim templateIndex As Long

    Dim templateName As String

    Dim templatePath As String

    Dim outputFileName As String

    Dim placeholderName As String

    Dim replacementValue As String

    Dim recordName As String

    Dim rowGenerated As Boolean

    lastRow = UBound(dataArray, 1)

    lastCol = UBound(dataArray, 2)

    For rowIndex = 2 To lastRow

        If ShouldProcessRow(dataArray(rowIndex, 1)) Then

            recordName = VariantToString(dataArray(rowIndex, 2), True)

            If Len(recordName) = 0 Then

                AddWarning warnings, "Row " & CStr(rowIndex) & " skipped because the output file name is empty."

                GoTo NextRow

            End If

            templateItems = Split(VariantToString(dataArray(rowIndex, 3), True), ";")

            rowGenerated = False

            For templateIndex = LBound(templateItems) To UBound(templateItems)

                templateName = NormalizeTemplateName(VariantToString(templateItems(templateIndex), True))

                If Len(templateName) = 0 Then

                    GoTo NextTemplate

                End If

                templatePath = templateFolder & templateName & ".docx"

                If Not FileExists(templatePath) Then

                    AddWarning warnings, "Template '" & templateName & ".docx' was not found."

                    GoTo NextTemplate

                End If

                Set iWord = AppWord.Documents.Open(templatePath, ReadOnly:=True)

                For columnIndex = DATA_REPLACEMENT_START_COLUMN To lastCol

                    placeholderName = VariantToString(dataArray(1, columnIndex), True)

                    If Len(placeholderName) > 0 Then

                        replacementValue = BuildReplacementValue(dataArray(rowIndex, columnIndex), columnIndex)
                        Call ReplacePlaceholderInDocument(iWord, placeholderName, replacementValue)

                    End If

                Next columnIndex

                outputFileName = BuildUniqueFilePath(outputFolder, recordName & " - " & templateName, ".docx")

                iWord.SaveAs fileName:=outputFileName, FileFormat:=WORD_XML_FORMAT

                iWord.Close False

                Set iWord = Nothing

                ProcessArray = ProcessArray + 1

                rowGenerated = True

NextTemplate:

                CloseCurrentDocument

            Next templateIndex

            If Not rowGenerated Then

                AddWarning warnings, "Row " & CStr(rowIndex) & " did not produce any documents."

            Else

                processedRows(CStr(rowIndex)) = True

            End If

        End If

NextRow:

    Next rowIndex

End Function

Private Function ReplacePlaceholderInDocument(ByVal doc As Object, ByVal placeholderName As String, ByVal replacementValue As String) As Boolean

    Dim searchRange As Object

    Set searchRange = doc.Content

    With searchRange.Find

        .ClearFormatting

        .Replacement.ClearFormatting

        .Text = placeholderName

        .Forward = True

        .Wrap = WD_FIND_STOP

        .Format = False

        .MatchCase = False

        .MatchWholeWord = False

        .MatchAllWordForms = False

        .MatchSoundsLike = False

        .MatchWildcards = False

    End With

    Do While searchRange.Find.Execute

        ReplacePlaceholderInDocument = True

        searchRange.Text = replacementValue

        searchRange.Collapse WD_COLLAPSE_END

    Loop

End Function

Private Function SaveDataToHistorySheet(ByVal wsData As Worksheet, ByVal processedRows As Object) As Long

    Dim wsHistory As Worksheet

    Dim lastDataRow As Long

    Dim lastDataCol As Long

    Dim lastHistoryRow As Long

    Dim dataRow As Long

    Dim recordName As String

    Dim templateList As String

    Dim dataWidth As Long

    If processedRows Is Nothing Then Exit Function

    If processedRows.Count = 0 Then Exit Function

    Set wsHistory = GetHistoryWorksheet()

    lastDataRow = GetLastUsedRow(wsData)

    lastDataCol = GetLastUsedColumn(wsData)

    dataWidth = lastDataCol - 3

    EnsureHistoryHeaders wsHistory, wsData, lastDataCol

    lastHistoryRow = GetLastUsedRow(wsHistory)

    For dataRow = 2 To lastDataRow

        If processedRows.Exists(CStr(dataRow)) Then

            recordName = VariantToString(wsData.Cells(dataRow, 2).Value, True)

            templateList = VariantToString(wsData.Cells(dataRow, 3).Value, True)

            If Not HistoryContainsRecord(wsHistory, recordName, templateList) Then

                lastHistoryRow = lastHistoryRow + 1

                wsHistory.Cells(lastHistoryRow, 1).Value = lastHistoryRow - HISTORY_HEADER_ROW

                wsHistory.Cells(lastHistoryRow, 2).Value = Now

                wsHistory.Cells(lastHistoryRow, 3).Value = recordName

                wsHistory.Cells(lastHistoryRow, 4).Value = templateList

                wsHistory.Cells(lastHistoryRow, 5).Resize(1, dataWidth).Value = wsData.Cells(dataRow, 4).Resize(1, dataWidth).Value

                SaveDataToHistorySheet = SaveDataToHistorySheet + 1

            End If

        End If

    Next dataRow

    wsHistory.Columns.AutoFit

End Function

Private Sub EnsureHistoryHeaders(ByVal wsHistory As Worksheet, ByVal wsData As Worksheet, ByVal lastDataCol As Long)

    Dim headerWidth As Long

    headerWidth = lastDataCol - 3

    wsHistory.Cells(1, 1).Value = "No."

    wsHistory.Cells(1, 2).Value = "Created On"

    wsHistory.Cells(1, 3).Value = "Record Name"

    wsHistory.Cells(1, 4).Value = "Template List"

    wsHistory.Cells(1, 5).Resize(1, headerWidth).Value = wsData.Cells(1, 4).Resize(1, headerWidth).Value

End Sub

Private Function HistoryContainsRecord(ByVal wsHistory As Worksheet, ByVal recordName As String, ByVal templateList As String) As Boolean

    Dim lastRow As Long

    Dim rowIndex As Long

    lastRow = GetLastUsedRow(wsHistory)

    For rowIndex = 2 To lastRow

        If StrComp(VariantToString(wsHistory.Cells(rowIndex, 3).Value, True), recordName, vbTextCompare) = 0 _
           And StrComp(VariantToString(wsHistory.Cells(rowIndex, 4).Value, True), templateList, vbTextCompare) = 0 Then

            HistoryContainsRecord = True

            Exit Function

        End If

    Next rowIndex

End Function

Private Function ExportSummaryWorkbook(ByVal sourceSheet As Worksheet) As String

    Dim newWorkbook As Workbook

    Dim targetSheet As Worksheet

    Dim lastRow As Long

    Dim lastCol As Long

    Dim dataWidth As Long

    Dim rowIndex As Long

    Dim outputPath As String

    Dim hasSortField As Boolean

    lastRow = GetLastUsedRow(sourceSheet)

    lastCol = GetLastUsedColumn(sourceSheet)

    dataWidth = lastCol - 3

    If dataWidth <= 0 Then Exit Function

    Set newWorkbook = Workbooks.Add

    Set targetSheet = newWorkbook.Worksheets(1)

    targetSheet.Cells(1, 1).Value = "No."

    sourceSheet.Range("D1").Resize(lastRow, dataWidth).Copy

    targetSheet.Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    targetSheet.Range("B1").PasteSpecial Paste:=xlPasteFormats

    Application.CutCopyMode = False

    If lastRow > 2 Then

        With targetSheet.Sort

            .SortFields.Clear

            If dataWidth >= 5 Then

                .SortFields.Add Key:=targetSheet.Range("F2:F" & lastRow), Order:=xlAscending

                hasSortField = True

            End If

            If dataWidth >= 2 Then

                .SortFields.Add Key:=targetSheet.Range("C2:C" & lastRow), Order:=xlAscending

                hasSortField = True

            End If

            If hasSortField Then

                .SetRange targetSheet.Range("A1").Resize(lastRow, dataWidth + 1)

                .Header = xlYes

                .Apply

            End If

        End With

    End If

    For rowIndex = 2 To lastRow

        targetSheet.Cells(rowIndex, 1).Value = rowIndex - 1

    Next rowIndex

    targetSheet.Range("A1:A" & lastRow).Borders.LineStyle = xlContinuous

    targetSheet.Cells.EntireColumn.AutoFit

    targetSheet.Cells.EntireRow.AutoFit

    With targetSheet.PageSetup

        .Orientation = xlLandscape

        .Zoom = False

        .FitToPagesWide = 1

        .FitToPagesTall = False

    End With

    outputPath = BuildUniqueFilePath(GetDefaultBaseFolder(), SUMMARY_FILE_PREFIX & Format(Date, "yyyy-mm-dd"), ".xlsx")

    newWorkbook.SaveAs fileName:=outputPath, FileFormat:=xlOpenXMLWorkbook

    newWorkbook.Close SaveChanges:=False

    ExportSummaryWorkbook = outputPath

End Function

Private Function ResolveWorksheetByName(ByVal preferredName As String, ByVal legacyName As String, ByVal createIfMissing As Boolean) As Worksheet

    Dim ws As Worksheet

    Set ws = TryGetWorksheet(preferredName)

    If Not ws Is Nothing Then

        Set ResolveWorksheetByName = ws

        Exit Function

    End If

    If Len(legacyName) > 0 Then

        Set ws = TryGetWorksheet(legacyName)

        If Not ws Is Nothing Then

            ws.Name = preferredName

            Set ResolveWorksheetByName = ws

            Exit Function

        End If

    End If

    If createIfMissing Then

        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

        ws.Name = preferredName

        Set ResolveWorksheetByName = ws

        Exit Function

    End If

    Err.Raise vbObjectError + 1000, "ResolveWorksheetByName", "Worksheet '" & preferredName & "' was not found."

End Function

Private Function TryGetWorksheet(ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets

        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then

            Set TryGetWorksheet = ws

            Exit Function

        End If

    Next ws

End Function

Private Sub ValidateTemplatePickerRange()

    Dim configuredTemplates As String

    configuredTemplates = GetConfiguredTextSetting(FILE_TEMPLATE_RANGE_NAME)

    If Len(configuredTemplates) = 0 Then

        Err.Raise vbObjectError + 1001, "ValidateTemplatePickerRange", "Named range '" & FILE_TEMPLATE_RANGE_NAME & "' is empty."

    End If

End Sub

Private Function GetFolderPath(ByVal rangeName As String) As String

    GetFolderPath = EnsureTrailingSlash(GetConfiguredTextSetting(rangeName))

    If Len(GetFolderPath) = 0 Then

        Err.Raise vbObjectError + 1002, "GetFolderPath", "Named range '" & rangeName & "' is empty."

    End If

    If dir$(GetFolderPath, vbDirectory) = vbNullString Then

        Err.Raise vbObjectError + 1003, "GetFolderPath", "Folder not found: " & GetFolderPath

    End If

End Function

Private Function GetConfiguredTextSetting(ByVal settingName As String) As String

    Dim legacyValue As String

    GetConfiguredTextSetting = GetStoredTextSetting(settingName)

    If Len(GetConfiguredTextSetting) > 0 Then Exit Function

    legacyValue = GetLegacyRangeValue(settingName)

    If Len(legacyValue) > 0 Then

        SaveStoredTextSetting settingName, legacyValue

        GetConfiguredTextSetting = legacyValue

        Exit Function

    End If

    Err.Raise vbObjectError + 1004, "GetConfiguredTextSetting", "Setting '" & settingName & "' was not found."

End Function

Private Function GetLegacyRangeValue(ByVal rangeName As String) As String

    On Error GoTo MissingLegacyRange

    GetLegacyRangeValue = VariantToString(ThisWorkbook.Names(rangeName).RefersToRange.Value, True)

    Exit Function

MissingLegacyRange:

    Err.Clear

End Function

Private Function GetCertificateOutputFolder() As String

    Dim outputFolder As String

    outputFolder = GetStoredTextSetting(OUTPUT_FOLDER_SETTING_NAME)

    If Len(outputFolder) = 0 Or dir$(outputFolder, vbDirectory) = vbNullString Then

        outputFolder = PickFolderPath("Select a folder for generated certificates", GetDefaultBaseFolder())

        If Len(outputFolder) = 0 Then

            Err.Raise vbObjectError + 1005, "GetCertificateOutputFolder", "No output folder was selected."

        End If

        SaveStoredTextSetting OUTPUT_FOLDER_SETTING_NAME, outputFolder

    End If

    GetCertificateOutputFolder = EnsureTrailingSlash(outputFolder)

End Function

Private Function PickFolderPath(ByVal dialogTitle As String, ByVal initialFolder As String) As String

    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = dialogTitle

        If Len(initialFolder) > 0 And dir$(initialFolder, vbDirectory) <> vbNullString Then

            .InitialFileName = EnsureTrailingSlash(initialFolder)

        End If

        If .Show = -1 Then

            PickFolderPath = .SelectedItems(1)

        End If

    End With

End Function

Private Function GetStoredTextSetting(ByVal settingName As String) As String

    On Error GoTo MissingSetting

    GetStoredTextSetting = DecodeStoredNameValue(ThisWorkbook.Names(settingName).refersTo)

    Exit Function

MissingSetting:

    Err.Clear

End Function

Private Sub SaveStoredTextSetting(ByVal settingName As String, ByVal settingValue As String)

    Dim encodedValue As String

    encodedValue = "=" & Chr$(34) & Replace$(settingValue, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)

    On Error Resume Next

    ThisWorkbook.Names(settingName).Delete

    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=settingName, refersTo:=encodedValue

End Sub

Private Function DecodeStoredNameValue(ByVal refersTo As String) As String

    If Left$(refersTo, 2) = "=" & Chr$(34) And Right$(refersTo, 1) = Chr$(34) Then

        DecodeStoredNameValue = Mid$(refersTo, 3, Len(refersTo) - 3)

        DecodeStoredNameValue = Replace$(DecodeStoredNameValue, Chr$(34) & Chr$(34), Chr$(34))

    ElseIf Left$(refersTo, 1) = "=" Then

        DecodeStoredNameValue = Mid$(refersTo, 2)

    End If

End Function

Private Function GetDefaultBaseFolder() As String

    If Len(ThisWorkbook.Path) > 0 Then

        GetDefaultBaseFolder = ThisWorkbook.Path

    Else

        GetDefaultBaseFolder = CreateObject("WScript.Shell").SpecialFolders("Desktop")

    End If

End Function

Private Function BuildUniqueFilePath(ByVal folderPath As String, ByVal baseFileName As String, ByVal extensionWithDot As String) As String

    Dim sanitizedName As String

    Dim candidatePath As String

    Dim counter As Long

    sanitizedName = SanitizeFileName(baseFileName)

    candidatePath = EnsureTrailingSlash(folderPath) & sanitizedName & extensionWithDot

    Do While FileExists(candidatePath)

        counter = counter + 1

        candidatePath = EnsureTrailingSlash(folderPath) & sanitizedName & " (" & CStr(counter) & ")" & extensionWithDot

    Loop

    BuildUniqueFilePath = candidatePath

End Function

Private Function ReadExcelData(ByVal ws As Worksheet) As Variant

    Dim lastRow As Long

    Dim lastCol As Long

    lastRow = GetLastUsedRow(ws)

    lastCol = GetLastUsedColumn(ws)

    If lastCol < MIN_REQUIRED_DATA_COLUMNS Then

        Err.Raise vbObjectError + 1006, "ReadExcelData", "Worksheet '" & ws.Name & "' must contain at least " & CStr(MIN_REQUIRED_DATA_COLUMNS) & " columns."

    End If

    ReadExcelData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value

End Function

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long

    Dim lastCell As Range

    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    If lastCell Is Nothing Then

        GetLastUsedRow = 1

    Else

        GetLastUsedRow = lastCell.Row

    End If

End Function

Private Function GetLastUsedColumn(ByVal ws As Worksheet) As Long

    Dim lastCell As Range

    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)

    If lastCell Is Nothing Then

        GetLastUsedColumn = 1

    Else

        GetLastUsedColumn = lastCell.Column

    End If

End Function

Private Function ShouldProcessRow(ByVal statusValue As Variant) As Boolean

    ShouldProcessRow = StrComp(VariantToString(statusValue, True), "ok", vbTextCompare) = 0

End Function

Private Function BuildReplacementValue(ByVal sourceValue As Variant, ByVal columnIndex As Long) As String

    Dim rawValue As String

    rawValue = VariantToString(sourceValue, False)

    If columnIndex = DATA_REPLACEMENT_START_COLUMN And Len(Trim$(rawValue)) > 0 Then

        BuildReplacementValue = GetRecipientNameValue(Trim$(rawValue))

    ElseIf columnIndex = UNIT_REPLACEMENT_COLUMN And Len(Trim$(rawValue)) > 0 Then

        BuildReplacementValue = GetDeclinedUnitValue(rawValue)

    Else

        BuildReplacementValue = rawValue

    End If

End Function

Private Function GetRecipientNameValue(ByVal fullName As String) As String

    On Error Resume Next

    GetRecipientNameValue = Trim$(FIO(fullName, "D", False))

    On Error GoTo 0

    If Len(GetRecipientNameValue) = 0 Then

        GetRecipientNameValue = DativeCase(fullName)

    End If

End Function

Private Function GetDeclinedUnitValue(ByVal unitValue As String) As String

    Dim normalizedValue As String

    Dim lowerValue As String

    Dim suffixValue As String

    normalizedValue = NormalizeUnitValue(unitValue)

    If Len(normalizedValue) = 0 Then Exit Function

    lowerValue = LCase$(normalizedValue)

    If lowerValue Like "войсковая часть *" Then

        suffixValue = Trim$(Mid$(normalizedValue, Len("Войсковая часть") + 1))

        GetDeclinedUnitValue = "войсковой части " & suffixValue

    ElseIf lowerValue Like "в/ч *" Then

        suffixValue = Trim$(Mid$(normalizedValue, Len("В/Ч") + 1))

        GetDeclinedUnitValue = "войсковой части " & suffixValue

    ElseIf IsNumeric(Replace$(normalizedValue, " ", vbNullString)) Then

        GetDeclinedUnitValue = "войсковой части " & normalizedValue

    Else

        GetDeclinedUnitValue = normalizedValue

    End If

End Function

Private Function VariantToString(ByVal sourceValue As Variant, ByVal trimValue As Boolean) As String

    If IsError(sourceValue) Then Exit Function

    If IsEmpty(sourceValue) Then Exit Function

    If IsNull(sourceValue) Then Exit Function

    VariantToString = CStr(sourceValue)

    If trimValue Then

        VariantToString = Trim$(VariantToString)

    End If

End Function

Private Function NormalizeTemplateName(ByVal templateName As String) As String

    templateName = Trim$(templateName)

    If LCase$(Right$(templateName, 5)) = ".docx" Then

        templateName = Left$(templateName, Len(templateName) - 5)

    End If

    NormalizeTemplateName = templateName

End Function

Private Function SanitizeFileName(ByVal fileName As String) As String

    Dim invalidCharacters As Variant

    Dim item As Variant

    invalidCharacters = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For Each item In invalidCharacters

        fileName = Replace$(fileName, CStr(item), "_")

    Next item

    SanitizeFileName = Trim$(fileName)

End Function

Private Function FileExists(ByVal filePath As String) As Boolean

    FileExists = Len(dir$(filePath, vbNormal Or vbHidden Or vbReadOnly Or vbSystem)) > 0

End Function

Private Function EnsureTrailingSlash(ByVal folderPath As String) As String

    folderPath = Trim$(folderPath)

    If Len(folderPath) = 0 Then Exit Function

    If Right$(folderPath, 1) <> "\" Then

        folderPath = folderPath & "\"

    End If

    EnsureTrailingSlash = folderPath

End Function

Private Sub AddWarning(ByRef warnings As String, ByVal message As String)

    If Len(message) = 0 Then Exit Sub

    If InStr(1, warnings, message, vbTextCompare) > 0 Then Exit Sub

    If Len(warnings) > 0 Then

        warnings = warnings & vbCrLf

    End If

    warnings = warnings & message

End Sub

Private Sub CloseCurrentDocument()

    On Error Resume Next

    If Not iWord Is Nothing Then

        iWord.Close False

        Set iWord = Nothing

    End If

    On Error GoTo 0

End Sub

Private Sub CleanUpWordApplication()

    On Error Resume Next

    If Not AppWord Is Nothing Then

        AppWord.Quit

        Set AppWord = Nothing

    End If

    On Error GoTo 0

End Sub
