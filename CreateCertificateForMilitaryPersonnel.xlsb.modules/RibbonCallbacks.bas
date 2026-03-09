Attribute VB_Name = "RibbonCallbacks"
Option Explicit
' Version: 0.6.0
' Updated: 2026-03-09

Private pRibbon As IRibbonUI

Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    Set pRibbon = ribbon
End Sub

Public Sub RibbonGenerateCertificates(ByVal control As IRibbonControl)
    CreateDoc
End Sub

Public Sub RibbonImportSourceData(ByVal control As IRibbonControl)
    CreateAndImportDataSheet
End Sub

Public Sub RibbonOpenSearch(ByVal control As IRibbonControl)
    OpenSearchForm
End Sub

Public Sub RibbonOpenHistory(ByVal control As IRibbonControl)
    OpenHistorySheet
End Sub

Public Sub RibbonSelectTemplateFolder(ByVal control As IRibbonControl)
    SelectTemplateFolder
End Sub

Public Sub RibbonSelectTemplates(ByVal control As IRibbonControl)
    OpenTemplateManager
End Sub

Public Sub RibbonSelectOutputFolder(ByVal control As IRibbonControl)
    SelectCertificateOutputFolder
End Sub

Public Sub RibbonShowAbout(ByVal control As IRibbonControl)
    ShowAboutDialog
End Sub

Public Sub RibbonInvalidate()
    If Not pRibbon Is Nothing Then
        pRibbon.Invalidate
    End If
End Sub
