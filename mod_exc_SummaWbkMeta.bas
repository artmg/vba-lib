Attribute VB_Name = "mod_exc_SummaWbkMeta"
Option Explicit

' error handling tag             **********************
Const cStrModuleName As String = "mod_exc_SummaWbkMeta"
'                                **********************
'
' (c) Join the Bits ltd
'
' Create a summary of metadata for workbooks in a given folder tree
'
' This module is used to enumerate all XLS files in a folder tree 
' chosen by the user, and create an output table for each file with 
'   Filename
'   Path
'   Modified Date
'   # rows (in first worksheet only)
'   # cols (in first worksheet only)
'
'  160721.AMG  renamed from mod_exc_SummariseWbkMeta
'  150506.AMG  derived from mod_exc_SchemaReader
'

' REFERENCES
' ==========
'
' This module uses the following references (paths and GUIDs may vary)
'   (only those required by it's dependent modules)

' DEPENDENCIES
' ============
'
' This module uses the following vba-lib modules
' AND any References specified within them
'
' vba-lib / mod_off_FilesFoldersSitesLinks
' vba-lib / mod_off_ExportListToExcel
'

' IMPROVEMENTS
' ============
'
' * turn workbook opening code (shared with mod_exc_SummaWkshtSchemas) into generic function in mod_off_FilesFoldersSitesLinks
' * collect # rows and columns from ALL sheets, not just first
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


