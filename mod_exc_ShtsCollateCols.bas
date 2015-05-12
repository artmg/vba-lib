Attribute VB_Name = "mod_exc_ShtsCollateCols"
' mod_exc_ShtsCollateCols

' Open all workbooks in a given folder tree and
' Collate data from columns
' ASSUMPTIONS:
' - ONLY the first worksheet in each workbook contains data
' - Data is organised VERTICALLY
'   - one Record per Worksheet Column
'   - each row represents a data field
' - NO rows have been added or removed in individual workbooks

'  150506.AMG  created

' References
' ==========
'
' This module uses the following references (paths and GUIDs may vary)
'
' (only references from dependencies)


' DEPENDENCIES
' ============
'
' This module uses the following vba-lib modules
' AND any References specified within them
'
' vba-lib / mod_off_FilesFoldersSitesLinks
' vba-lib / mod_off_ExportListToExcel
' vba-lib / mod_exc_WbkShtRngName
'

' IMPROVEMENTS
' ============
'
' * (none identified yet)
'



' set this constant to say how many columns to ignore
' e.g. because they simply contain field names

Const cIntColumnsToIgnore As Integer = 6
Const cIntRowsToIgnore As Integer = 1


Sub CollateColumnDataFromWorkbooks()
    Dim strFileNames() As String
    strFileNames() = arrFilteredPathnamesInUserTree(strFilter:=".xlsm", bRecurse:=True)
' func to return the number of elements without error (0 if none)
    If strFileNames(0) <> "" Then
        PrepareListWithHeaders
        Dim ifile As Integer
        For ifile = 0 To UBound(strFileNames)
            CollateFromBook strFileNames(ifile)
        Next

        ExcelOutputShow
    
    End If
End Sub


Function PrepareListWithHeaders()
    ExcelOutputCreateWorksheet
    ExcelOutputWriteValue "Filename"
    ExcelOutputWriteValue "By"
    ExcelOutputWriteValue "Date"
    ExcelOutputWriteValue "Column"
    ExcelOutputWriteValue "Data Collated"
    ExcelOutputNextRow
End Function


Function CollateFromBook( _
  strWbkName As String _
)
    Dim wbk As Excel.Workbook

'Open Each Workbook
'For columns to used.range
'If (any cell in colun contains data)
'Copy Cells down to UsedRange
'Paste into NewSheet
'Add Filename and Column number into header rows

    Application.StatusBar = "reading from [" & strWbkName & " ]..."
    Application.ScreenUpdating = False

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
On Error GoTo 0
    Dim sht As Excel.Worksheet
    Set sht = wbk.Worksheets(1)

    Dim iCol, iLastRow As Integer
    iLastRow = sht.UsedRange.Rows.Count
    For iCol = (cIntColumnsToIgnore + 1) To sht.UsedRange.Columns.Count

        ' refer to the data in the individual source column
        Dim rngSourceCol As Excel.Range
        Set rngSourceCol = sht.Range(Cells(cIntRowsToIgnore + 1, iCol), Cells(iLastRow, iCol))
        'Ignore empty columns
        If intCountValuesInRange(rngSourceCol) Then
            Application.ScreenUpdating = True

            ExcelOutputWriteValue JustFileName(strWbkName)
            ExcelOutputWriteValue wbk.BuiltinDocumentProperties("Last Author")
            ExcelOutputWriteValue wbk.BuiltinDocumentProperties("Last Save Time")
            ExcelOutputWriteValue iCol

            rngSourceCol.Copy
            ExcelOutputRngCurrentCell.PasteSpecial _
                Paste:=XlPasteType.xlPasteValues _
                , Transpose:=True
'            ' Paste Transpose Values Only from dest col 5
'    ' HOW TO avoid errors on (nulls?)
'            ExcelOutputWriteValue sht.Cells(7, iCol).Value
'            ExcelOutputWriteValue sht.Cells(8, iCol).Value
'            ExcelOutputWriteValue sht.Cells(9, iCol).Value
'            ExcelOutputWriteValue sht.Cells(3, iCol).Value
            ExcelOutputNextRow
            Application.ScreenUpdating = False
        End If
    Next iCol

    ' Cancel Clipboard to avoid messages about large amount of data when closing workbook
    Application.CutCopyMode = False

    wbk.Close SaveChanges:=False
    Application.ScreenUpdating = True

OpenError:
    Application.StatusBar = False
End Function


