Attribute VB_Name = "mod_exc_SummaWkshtSchemas"
Option Explicit

' error handling tag             ***************************
Const cStrModuleName As String = "mod_exc_SummaWkshtSchemas"
'                                ***************************
'
' (c) Join the Bits ltd
'
' This module is used to enumerate all XLS files in a
' folder, chosen by the user, and examine the schemas
' used in every sheet
' Basically it populates the sheet in THIS workbook with
' the spreadsheet name, the worksheet name and all
' column headings from row A
'
'  160721.AMG  renamed from mod_exc_SchemaReader
'  150511.AMG  standardised into vba-lib style and rationatised sub funcs
'  141105.AMG  do xls & xlsx, transpose, rowcount & doublespace
'  071030.AMG  created
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
' * turn workbook opening code (shared with mod_exc_SummaWbkMeta) into generic function in mod_off_FilesFoldersSitesLinks
'


' Const cStrFileFilter As String = "Excel Workbooks, *.xls; *.xlsx"
Const cStrFileFilter As String = ".xls|.xlsx|.xltx|.xlsm"
Const cbDoubleRow As Boolean = True

Sub EnumerateExcelSchemas()
    Dim strFileNames() As String
    strFileNames() = arrFilteredPathnamesInUserTree(strFilter:=cStrFileFilter, bRecurse:=False)
' func to return the number of elements without error (0 if none)
    If strFileNames(0) <> "" Then
        PrepareListWithHeaders
        Dim ifile As Integer
        For ifile = 0 To UBound(strFileNames)
            AddSchemaToListFor strFileNames(ifile)
        Next
        ExcelOutputShow
        MsgBox "Finished reading Excel worksheet schemas from source folder"
    End If
End Sub


Function AddSchemaToListFor( _
  strWbkName As String _
)
    Dim wbk As Workbook

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


    Dim sht As Worksheet
    For Each sht In wbk.Worksheets
        ExcelOutputWriteID
        ExcelOutputWriteValue JustFileName(strWbkName)
        ExcelOutputWriteValue sht.Name
        ExcelOutputWriteValue sht.UsedRange.Rows.Count
        
        Dim col As Range
        For Each col In sht.UsedRange.Columns
            ExcelOutputWriteValue col.Cells(1).Value
        Next col

        ExcelOutputNextRow (cbDoubleRow)
    Next sht

    wbk.Close SaveChanges:=False

OpenError:
    Application.StatusBar = False
End Function


Function PrepareListWithHeaders()
    ExcelOutputCreateWorksheet
    ExcelOutputWriteValue "ID"
    ExcelOutputWriteValue "Workbook"
    ExcelOutputWriteValue "Sheet"
    ExcelOutputWriteValue "RowsUsed"
    ExcelOutputWriteValue "Fields"
    ExcelOutputMakeHeaderRow
    ExcelOutputNextRow
End Function

