Attribute VB_Name = "mod_exc_ConsolInteligence"

' mod_exc_ConsolInteligence
'
' Consolidate tables intelligently from multiple sources
'
' Use a table defining the source table locations and layouts
' to pull specified columns from various sources
' into one single long table for analysis

' 141113.AMG cribbed from various


' REQUIRES:
'   mod_exc_WorkbooksSheetsNames


' SourceDefinitions sheet contains columns for:
'   SourceID
'   Path
'   File
'   Sheet
'
' then a series of Destination Column Names
'
' and finally a column called
'   Exceptions
' which gives logic stating which rows NOT to import

Dim intFirstDefCol, intLastDefCol As Integer
Dim intNextOutRow As Integer

Sub ConsolidateWithIntelligence()

    Dim shtDefs As Worksheet
    Dim shtOutput As Worksheet

    Set shtDefs = ActiveWorkbook.Worksheets("SourceDefinitions")
    Set shtOutput = getSheetOrCreateIfNotFound(ActiveWorkbook, "ConsolidatedHosts")
    ClearEntireSheet shtOutput
    AddColumnHeadersFrom shtDefs, shtOutput

    Dim intDefRow As Integer
    For intDefRow = 2 To shtDefs.UsedRange.Rows.Count
        CopyDataFromSourceTo shtOutput, shtDefs.Rows(intDefRow)
    Next intDefRow
End Sub


Function AddColumnHeadersFrom(shtDefs As Worksheet, shtOutput As Worksheet)
' also sets first and last column numbers from SourceDefinition
    
    ' copy ID header
    shtOutput.Cells(1, 1).Value = shtDefs.Cells(1, 1).Value
    
    ' copy the rest until "Exceptions" found
    intFirstDefCol = 5
    Dim intCol As Integer
    For intCol = intFirstDefCol To shtDefs.UsedRange.Columns.Count
        If shtDefs.Cells(1, intCol).Value = "Exceptions" Then
            intCol = 9999
        Else
            intLastDefCol = intCol
            shtOutput.Cells(1, intCol - intFirstDefCol + 2).Value = shtDefs.Cells(1, intCol).Value
        End If
    Next intCol
    intNextOutRow = 2
End Function

Function CopyDataFromSourceTo(ByRef shtOutput As Worksheet, ByRef rngDefRow As Range)
    Dim strSourceID, strSourceFile As String
    Dim wbk As Workbook
    Dim shtSource As Worksheet

    strSourceID = rngDefRow.Cells(1, 1).Value
    strSourceFile = rngDefRow.Cells(1, 2).Value & rngDefRow.Cells(1, 3).Value
    On Error GoTo InvalidSource
    Set wbk = Application.Workbooks.Open(strSourceFile, ReadOnly:=True)
    Set shtSource = wbk.Worksheets(rngDefRow.Cells(1, 4).Value)
    On Error GoTo 0

    Dim intSrcRow As Integer
    For intSrcRow = 2 To shtSource.UsedRange.Rows.Count
        CopyRowFromSourceTo shtOutput.Rows(intNextOutRow), rngDefRow, shtSource.Rows(intSrcRow)
        shtOutput.Cells(intNextOutRow, 1).Value = strSourceID
        intNextOutRow = intNextOutRow + 1
    Next intSrcRow

    GoTo Continue:

InvalidSource:
    shtOutput.Cells(intNextOutRow, 1).Value = strSourceID
    shtOutput.Cells(intNextOutRow, 2).Value = "*** INVALID SOURCE! ***"
    intNextOutRow = intNextOutRow + 1

Continue:
    If Not wbk Is Nothing Then wbk.Close
End Function

Function CopyRowFromSourceTo(ByRef rngOutRow As Range, ByRef rngDefRow As Range, ByRef rngSrcRow As Range)
    Dim intCol As Integer
    For intCol = intFirstDefCol To intLastDefCol
        rngOutRow.Cells(1, intCol - intFirstDefCol + 2).Value = rngSrcRow.Cells(1, CInt(rngDefRow.Cells(1, intCol).Value)).Value
    Next intCol

End Function

