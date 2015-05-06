Attribute VB_Name = "mod_exc_SummariseWbkMeta"
' mod_exc_SummariseWbkMeta

' create a summary of metadata for workbooks in a given folder tree

'Walk a tree chosen by user to create an output table in excel
'For each file list:
'   Filename
'   Path
'   Modified Date
'   # rows
'   # cols

' Depends on
' ==========
'
' This module uses the following vba-lib modules
' AND any References specified within them
'
' mod_off_FilesFoldersSitesLinks
' mod_off_ExportListToExcel
'



Sub SummariseWorkbookMetadata()

    Dim strFileNames() As String
    strFileNames() = arrFilteredPathnamesInUserTree(strFilter:=".xlsm", bRecurse:=True)
' func to return the number of elements without error (0 if none)
    If strFileNames(0) <> "" Then
        PrepareListWithHeaders
        Dim ifile As Integer
        For ifile = 0 To UBound(strFileNames)
            AddMetadataToListFor strFileNames(ifile)
        Next
        ExcelOutputShow
    End If
End Sub



Function PrepareListWithHeaders()
    ExcelOutputCreateWorksheet
    ExcelOutputWriteValue "Filename"
    ExcelOutputWriteValue "Path"
    ExcelOutputWriteValue "Modified"
    ExcelOutputWriteValue "Size"
    ExcelOutputWriteValue "Author"
    ExcelOutputWriteValue "NumRows"
    ExcelOutputWriteValue "NumCols"
    ExcelOutputNextRow
End Function


Function AddMetadataToListFor( _
  strWbkName As String _
)
    Dim wbk As Workbook

    ExcelOutputWriteValue JustFileName(strWbkName)
    ExcelOutputWriteValue GetFolderFromFileName(strWbkName)
    ' This would be how to get date if the file was not open
    ' ExcelOutputWriteValue FileDateTime(strWbkName)

    Application.StatusBar = "reading from [" & strWbkName & " ]..."

    ' prevent the "enable macros?" dialog from loading
    ' credit - http://stackoverflow.com/a/16301905
    Application.EnableEvents = False
    Dim iAutoSecSave As Integer
    iAutoSecSave = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

On Error GoTo OpenError:
    Set wbk = Workbooks.Open( _
        FileName:=strWbkName _
        , UpdateLinks:=0 _
        , ReadOnly:=True _
        , IgnoreReadOnlyRecommended:=True _
        )
    Application.EnableEvents = True
    Application.AutomationSecurity = iAutoSecSave

    ExcelOutputWriteValue wbk.BuiltinDocumentProperties("Last Save Time")
    ' help https://msdn.microsoft.com/en-us/library/office/ff197172.aspx
    ' Excel does not maintain BuiltinDocumentProperties("Number of Bytes")
    ExcelOutputWriteValue FileLen(strWbkName)
    ExcelOutputWriteValue wbk.BuiltinDocumentProperties("Last Author")

    ExcelOutputWriteValue wbk.Sheets(1).UsedRange.Rows.Count
    ExcelOutputWriteValue wbk.Sheets(1).UsedRange.Columns.Count

    wbk.Close SaveChanges:=False

OpenError:
    ExcelOutputNextRow
    Application.StatusBar = False
End Function


