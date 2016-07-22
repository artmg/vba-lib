Attribute VB_Name = "mod_exc_SummaWbkCellData"
Option Explicit

' error handling tag             **************************
Const cStrModuleName As String = "mod_exc_SummaWbkCellData"
'                                **************************
'
' (c) Join the Bits ltd
'
' Create a summary table by collating specific cell data from all workbooks in a given folder tree
'
' This module is used to enumerate all XLS files in a folder tree
' chosen by the user, and create an output table for each file.
' The column names, and the cell reference for each data value, are
' taken from a sheet in the current workbook called CellList with columns
'   NewColumnName
'   Worksheet
'   Column
'   Row
' So the output worksheet contains a filename Column and each of
' the columns defined in CellList, along with a row of data for each file
'
'  160721.AMG  use generic workbook opening code in mod_off_FilesFoldersSitesLinks
'  160721.AMG  derived from mod_exc_SummaWbkMeta and mod_exc_SummaWkshtSchemas
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
' vba-lib / mod_exc_DataTables
'

' IMPROVEMENTS
' ============
'
' * none for now
'

Const cStrFileFilter As String = ".xls|.xlsx|.xltx|.xlsm"
Const cStrCellListSheetName = "CellList"
Dim shtCellList As Excel.Worksheet
Dim rngCellList As Excel.Range

Sub SummariseCellDataFromWorkbooksInPath()
    Dim strFileNames() As String
    strFileNames() = arrFilteredPathnamesInUserTree(cStrFileFilter, bRecurse:=True)
' func to return the number of elements without error (0 if none)
    If strFileNames(0) <> "" Then
        
        Set shtCellList = Excel.ActiveWorkbook.Worksheets(cStrCellListSheetName)
        Set rngCellList = rngGetTableDataFromSheet(shtFromWorksheet:=shtCellList, lngNumHeaders:=1)
        PrepareListWithHeaders
        
        Dim ifile As Integer
        For ifile = 0 To UBound(strFileNames)
            AddCellDataToListFor strFileNames(ifile)
        Next
        
        ExcelOutputShow
        MsgBox "Finished summarising Excel workbook cell data from source folder"
    End If
End Sub



Function PrepareListWithHeaders()
    ExcelOutputCreateWorksheet
    ExcelOutputWriteValue "Filename"
    Dim rw As Range
    
    For Each rw In rngCellList.Rows
        ExcelOutputWriteValue CStr(rw.Cells(1, 1).Value)
    Next
    
    ExcelOutputNextRow
End Function

Function AddCellDataToListFor( _
  strWbkName As String _
)
    Application.StatusBar = "reading from [" & strWbkName & " ]..."
    ExcelOutputWriteValue JustFileName(strWbkName)
    Dim wbk As Workbook
    Set wbk = wbkOpenSafelyToRead(strWbkName)
    If Not wbk Is Nothing Then
    
        Dim rw As Excel.Range
        For Each rw In rngCellList.Rows
' debug            ExcelOutputWriteValue CStr(rw.Cells(1, 2).Value) & "|" & CStr(rw.Cells(1, 3).Value) & "|" & CStr(rw.Cells(1, 4).Value)
            ' should this use C<Type> for safety?
            ExcelOutputWriteValue wbk.Worksheets(rw.Cells(1, 2).Value).Cells(rw.Cells(1, 4).Value, rw.Cells(1, 3).Value).Value
        Next rw

    End If
    wbkCloseSafely wbk
    ExcelOutputNextRow
    Application.StatusBar = False
End Function


