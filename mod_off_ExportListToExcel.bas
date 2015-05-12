Attribute VB_Name = "mod_off_ExportListToExcel"
Option Explicit

' error handling tag             ***************************
Const cStrModuleName As String = "mod_off_ExportListToExcel"
'                                ***************************

' May be called by ANY MS Office app to quickly create an Excel table list
'
' 150511.AMG  standardised style and added MakeHeaderRow and ID (and previously added range return frig)
' 150316.AMG  debug pointer issue
' 150303.AMG  created

' References
' ==========
'
' This module uses the following references (paths and GUIDs may vary)
'
' ONLY required if NOT running from EXCEL application
' Microsoft Excel 15.0 Object Library (C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE) {00020813-0000-0000-C000-000000000046}
'

' DEPENDENCIES
' ============
'
' No vba-lib depends yet
'

' IMPROVEMENTS
' ============
'
' * Function to resize columns to fit contents
'


Dim shtOut As Excel.Worksheet
Dim lngCurrentRow As Long
Dim lngCurrentCol As Long
Dim lngCurrentID As Long

Function ExcelOutputCreateWorksheet()
    Dim wbk As Excel.Workbook
    Set wbk = Excel.Application.Workbooks.Add
    Set shtOut = wbk.Worksheets(1)
    
    lngCurrentRow = 1
    lngCurrentCol = 1
    lngCurrentID = 1
End Function

Function ExcelOutputWriteValue(val As Variant)
    shtOut.Cells(lngCurrentRow, lngCurrentCol).Value = val
    lngCurrentCol = lngCurrentCol + 1
End Function

Function ExcelOutputWriteID()
    shtOut.Cells(lngCurrentRow, lngCurrentCol).Value = lngCurrentID
    lngCurrentCol = lngCurrentCol + 1
End Function

Function ExcelOutputNextRow( _
    Optional ByVal bDoubleSpace As Boolean = False _
    )

    lngCurrentRow = lngCurrentRow + IIf(bDoubleSpace, 2, 1)
    lngCurrentID = lngCurrentID + 1
    lngCurrentCol = 1
End Function

Function ExcelOutputMakeHeaderRow()
    shtOut.Rows(lngCurrentRow).Font.Bold = True
    shtOut.Activate
    shtOut.Cells(lngCurrentRow + 1, 1).Select
    ActiveWindow.FreezePanes = True
    lngCurrentID = 0
End Function

Function ExcelOutputShow()
    shtOut.Activate
'    Excel.Application.ActivateMicrosoftApp
End Function

Function ExcelOutputRngCurrentCell() As Excel.Range
' this is a bit of a frig to allow the calling module to do its own thing with the data
' beware if you USE the range to enter data as the Current Row will NOT be changed
' so it should only be used as the last action before ExcelOutputNextRow
    Set ExcelOutputRngCurrentCell = shtOut.Cells(lngCurrentRow, lngCurrentCol)
End Function
