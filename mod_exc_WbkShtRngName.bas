Attribute VB_Name = "mod_exc_WbkShtRngName"
Option Explicit

' error handling tag             **********************
Const cStrModuleName As String = "mod_exc_WbkShtRngName"
'                                **********************

'
' generic functions for handling Excel Workbooks, Worksheets, Ranges and Names
'

'  150507.AMG  renamed from mod_exc_WorkbooksSheetsNames
'  141113.AMG  added Clear
'  130821.AMG  created

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
' (none currently)
'

' IMPROVEMENTS
' ============
'
' * Begin declaring Object classes using "Excel." to allow the module to be included in other Office apps
' * Also use "Excel.Application." references
' * move bTestMatch out to mod_off_ConvertLogic
'


''' this does not belong here
' move out to mod_off_ConvertLogic

Function bTestMatch _
(strLookAt As String _
, strLookFor As String _
, bExactMatch As Boolean _
) As Boolean
    If bExactMatch Then
        bTestMatch = UCase(strLookAt) = UCase(strLookFor)
    Else
        bTestMatch = UCase(Left(strLookAt, Len(strLookFor))) = UCase(strLookFor)
    End If
End Function




''' WORKSHEETS ''''''''''''''''''''''''''

Function getSheetOrCreateIfNotFound _
(wbk As Workbook _
, strWorksheetName As String _
) As Worksheet

    Dim sht As Worksheet
    Dim bFound As Boolean
    
    ' first see if we can find the named sheet
    bFound = False
    For Each sht In wbk.Sheets
        If sht.Name = strWorksheetName Then
            bFound = True
            Exit For
        End If
    Next
    
    ' create if not found
    If Not bFound Then
        Set sht = wbk.Worksheets.Add(After:=wbk.Worksheets(wbk.Worksheets.Count))
        sht.Name = strWorksheetName
    End If
    
    Set getSheetOrCreateIfNotFound = sht
End Function

Sub ClearEntireSheet(sht As Worksheet)
    sht.UsedRange.Clear
End Sub




''' RANGES '''''''''''''''''''''''''''''''

' Consumer Help - for notes on how to define ranges in VBA see - https://support.microsoft.com/en-us/kb/291308
' Developer Help - MSDN definition of Range Object Members - https://msdn.microsoft.com/en-us/library/office/ff838238.aspx

Function intCountValuesInRange( _
  Rng As Range _
) As Integer

    intCountValuesInRange = Application.WorksheetFunction.CountA(Rng)

End Function





''' NAMES ''''''''''''''''''''''''''''''''

Function ClearNamedRangeFrom _
(obj As Object _
, strRangeName As String _
, bExactMatch As Boolean _
)

    Dim nm As Name
    Dim bMatch As Boolean
    For Each nm In obj.Names
        bMatch = bTestMatch(nm.Name, strRangeName, bExactMatch:=bExactMatch)
        ' if it's a NamedRange on a worksheet then the value of the 'Name' attribute may include the sheet name
        If (Not bMatch) And (TypeName(obj) = "Worksheet") Then
            bMatch = bTestMatch(nm.Name, obj.Name & "!" & strRangeName, bExactMatch:=bExactMatch)
        End If
        If bMatch Then
            nm.Delete
        End If
    Next
End Function

