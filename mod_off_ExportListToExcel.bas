Attribute VB_Name = "mod_off_ExportListToExcel"
' mod_off_ExportListToExcel
' 150303.AMG

' References
' ==========
'
' This module uses the following references (paths and GUIDs may vary)
'
' Microsoft Excel 15.0 Object Library (C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE) {00020813-0000-0000-C000-000000000046}

Option Explicit

Dim shtOut As Excel.Worksheet
Dim lngNextRow As Long
Dim lngNextCol As Long

Function ExcelOutputCreateWorksheet()
    Dim wbk As Excel.Workbook
    Set wbk = Excel.Application.Workbooks.Add
    Set shtOut = wbk.Worksheets(1)
    lngNextRow = 1
    lngNextCol = 1
End Function


Function ExcelOutputNextRow()
    lngNextRow = lngNextRow + 1
    lngNextCol = 1
End Function

Function ExcelOutputWriteValue(val As Variant)
    shtOut.Cells(lngNextRow, lngNextCol).Value = val
End Function

