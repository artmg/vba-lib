Attribute VB_Name = "mod_exc_DataTables"
Option Explicit

' error handling tag             ********************
Const cStrModuleName As String = "mod_exc_DataTables"
'                                ********************

'
' Practical subfunctions for manipulating data tables easily
'

'  150622.AMG  normalise table with multiple entries in one column
'  150611.AMG  new match options to trim trailing & leading spaces
'  150326.AMG  added Table Match functions previously in Consol module
'  150312.AMG  created with table creation



' REFERENCES
' ==========
'
' This module uses the following references (paths and GUIDs may vary)
'   (only those required by it's dependent modules)

' DEPENDENCIES
' ============
'
' This module requires the following vba-lib dependencies:
'   mod_exc_WbkShtRngName

' IMPROVEMENTS
' ============
'
' * simplify TransTblNormaliseMultiEntries to use mod_off_ExportListToExcel.bas

' PREPARATION
' ===========
'
' no special prep required
'
' Define match types

Public Enum enumDataTableMatchType
    MatchCaseSens
    MatchCaseInsens
    MatchCaseSensTrim
    MatchCaseInsensTrim
End Enum


'
' *** TABLE CREATION *****************************
'

Public Function FillTableDownByCopyingBlanksFromAbove()
    Dim Rng, rRow, rCel As Range
    Set Rng = ActiveSheet.UsedRange
    For Each rRow In Rng.Rows
        For Each rCel In rRow.Cells
            With rCel
                If (.Value = "") And (.row > 1) Then
                    .Value = Rng.Cells(.row - 1, .Column).Value
                End If
            End With
        Next
    Next
End Function

Public Function FillColumnDownByCopyingBlanksFromAbove()
    Dim rCel As Range
    For Each rCel In ActiveSheet.UsedRange.Columns(ActiveCell.Column).Cells
        With rCel
            If (.Value = "") And (.row > 1) Then
                .Value = ActiveSheet.Cells(.row - 1, .Column).Value
            End If
        End With
    Next
End Function



'
' *** TABLE SEARCH / MATCH *****************************
'

' find the first absolute row in a Table (OTPIONALLY between a given range of rows)
' where the value in a certain column matches the VALUE passed)
' or zero if no match is found
Public Function intMatchGetRow _
(ByVal strMatch As String _
, ByVal enumMatchType As enumDataTableMatchType _
, ByRef sht As Worksheet _
, ByVal intCol As Integer _
, ByVal intFirstRow As Integer _
, ByVal intLastRow As Integer _
, Optional ByVal strIgnore As String = "" _
)

'            intMyRow = intMatchGetRow _
'                (strMatch:="" _
'                , enumMatchType:=enumMatchType _
'                , sht:=shtMine _
'                , intCol:=0 _
'                , intFirstRow:=0 _
'                , intLastRow:=0 _
'                , strIgnore:="" _
'                )
        

' IMPROVEMENTS:
' see CAPS in description above
' more value types
' if intLastRow becomes OPTIONAL, then do we just continue until first blank or last row of used range?

    Dim intTryRow As Integer
    Dim strLookFor, strCheckValue As String

    intMatchGetRow = 0
    strLookFor = strMatchPrepareValue(strMatch, enumMatchType, strIgnore)

    If strLookFor <> "" Then
        For intTryRow = intFirstRow To intLastRow
            strCheckValue = strMatchPrepareValue(sht.Cells(intTryRow, intCol), enumMatchType, strIgnore)

            If bMatchCheckValues(strCheckValue, strLookFor, enumMatchType) Then
                ' return the value and break out
                intMatchGetRow = intTryRow
                intTryRow = intLastRow
            End If
        Next
    End If

End Function ' intMatchGetRow
    

Public Function strMatchPrepareValue _
(ByVal strUnprepared As String _
, ByVal enumMatchType As enumDataTableMatchType _
, Optional ByVal strIgnore As String = "" _
) As String

    Dim strKeyToMatch As String
    Dim strToReplace As String
    
    Select Case enumMatchType
        Case enumDataTableMatchType.MatchCaseInsens:
            strKeyToMatch = UCase(strUnprepared)
            strToReplace = UCase(strIgnore)

        Case enumDataTableMatchType.MatchCaseInsensTrim:
            strKeyToMatch = LTrim(RTrim(UCase(strUnprepared)))
            strToReplace = UCase(strIgnore)

        Case enumDataTableMatchType.MatchCaseSens:
            strKeyToMatch = strUnprepared
            strToReplace = strIgnore
        
        Case enumDataTableMatchType.MatchCaseSensTrim:
            strKeyToMatch = LTrim(RTrim(strUnprepared))
            strToReplace = strIgnore
        
    End Select
    
' ONLY USE IGNORE DEPENDING ON MATCH TYPE FLAG ??
    If strToReplace <> "" Then
        strKeyToMatch = Replace(strKeyToMatch, strToReplace, "")
    End If

    strMatchPrepareValue = strKeyToMatch
End Function

Public Function bMatchCheckValues _
(varFirst As Variant _
, varSecond As Variant _
, enumMatchType As enumDataTableMatchType _
) As Boolean
    
    ' default return value
    bMatchCheckValues = False

    ' assuming values already prepared
    If varFirst = varSecond Then
        bMatchCheckValues = True
    End If
End Function


'
' *** TABLE TRANSFORMATION ***
'

Sub SplitCellsWithKey()
Dim WorkRng As Range
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", "Split Cells with Key", WorkRng.Address, Type:=8)
    
' REWORK TO USE THIS FUNCTION
'    Call TransTblNormaliseMultiEntries("NormSvrs", WorkRng.Address)

'based partly on http://www.extendoffice.com/documents/excel/2211-excel-split-cell-by-carriage-return.html
'Update 20141024
Dim Rng As Range
On Error Resume Next
For Each Rng In WorkRng
    Dim lLFs As Long
    lLFs = VBA.Len(Rng) - VBA.Len(VBA.Replace(Rng, vbLf, ""))
    If lLFs > 0 Then
        Rng.Offset(1, 0).Resize(lLFs).Insert shift:=xlShiftDown
        Rng.Resize(lLFs + 1).Value = Application.WorksheetFunction.Transpose(VBA.Split(Rng, vbLf))
        Rng.Offset(1, -1).Resize(lLFs).Insert shift:=xlShiftDown
    End If
Next
End Sub

Function TransTblNormaliseMultiEntries _
    (strRange As String _
    , Optional strNewSheetName As String = "SheetNorm" _
    , Optional strDelim As String = vbLf _
)
' defaults to LineFeed delimeter, used in multi-line cells (ALT-ENTER)

    Dim shtOutput As Excel.Worksheet
    Set shtOutput = getSheetOrCreateIfNotFound(Excel.ActiveWorkbook, strNewSheetName)

    Dim rngRow As Range
    Dim rngSourceTable As Range
    Dim iOutRow As Integer
    iOutRow = 1
    For Each rngRow In rngSourceTable
        Dim iCountDelims As Integer
        'some credit - http://www.extendoffice.com/documents/excel/2211-excel-split-cell-by-carriage-return.html
        On Error Resume Next
        iCountDelims = VBA.Len(rngRow) - VBA.Len(VBA.Replace(rngRow, vbLf, ""))
        If iCountDelims > 0 Then
            rngRow.Offset(1, 0).Resize(iCountDelims).Insert shift:=xlShiftDown
            rngRow.Resize(iCountDelims + 1).Value = Application.WorksheetFunction.Transpose(VBA.Split(rngRow, vbLf))
            rngRow.Offset(1, -1).Resize(iCountDelims).Insert shift:=xlShiftDown
            
            shtOutput.Cells(1, 1).Value = "SourceID"
            shtOutput.Cells(1, 2).Value = "Path"
    
            iOutRow = iOutRow + 1
        End If
    Next
End Function

