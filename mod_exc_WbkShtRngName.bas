Attribute VB_Name = "mod_exc_WbkShtRngName"
Option Explicit

' error handling tag             **********************
Const cStrModuleName As String = "mod_exc_WbkShtRngName"
'                                **********************

'
' generic functions for handling Excel Workbooks, Worksheets, Ranges and Names
'

'  160722.AMG  move bTestMatch out to mod_off_ConvertLogic
'  160722.AMG  altered references to work from other application
'  160721.AMG  workbook open & close functions from mod_exc_Summa*
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
' * mod_off_ConvertLogic
'

' IMPROVEMENTS
' ============
'
' * Test this module from an Office app other than Excel
'


''' WORKBOOKS ''''''''''''''''''''''''''

Public Function wbkOpenSafelyToRead( _
  ByVal strFileName As String _
  , Optional ByRef xlApp As Excel.Application _
) As Excel.Workbook

    If xlApp Is Nothing Then
        Set xlApp = Excel.Application
    End If

    Dim wbk As Excel.Workbook
    Set wbk = Nothing
    
    ' prevent the "enable macros?" dialog from loading
    ' credit - http://stackoverflow.com/a/16301905
    xlApp.EnableEvents = False
    Dim iAutoSecSave As Integer
    iAutoSecSave = xlApp.AutomationSecurity
    xlApp.AutomationSecurity = msoAutomationSecurityForceDisable
'    Application.EnableEvents = False
'    Dim iAutoSecSave As Integer
'    iAutoSecSave = Application.AutomationSecurity
'    Application.AutomationSecurity = msoAutomationSecurityForceDisable

On Error GoTo OpenError:
    Set wbk = xlApp.Workbooks.Open( _
        FileName:=strFileName _
        , UpdateLinks:=0 _
        , ReadOnly:=True _
        , IgnoreReadOnlyRecommended:=True _
        )

OpenError:

    xlApp.EnableEvents = True
    xlApp.AutomationSecurity = iAutoSecSave
'    Application.EnableEvents = True
'    Application.AutomationSecurity = iAutoSecSave

    Set wbkOpenSafelyToRead = wbk
End Function

Public Function wbkCloseSafely( _
  ByRef wbk As Excel.Workbook _
)
    If Not wbk Is Nothing Then
        wbk.Close SaveChanges:=False
    End If
End Function


''' WORKSHEETS ''''''''''''''''''''''''''

Function getSheetOrCreateIfNotFound _
(wbk As Excel.Workbook _
, strWorksheetName As String _
) As Excel.Worksheet

    Dim sht As Excel.Worksheet
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
  rng As Range _
) As Integer

    intCountValuesInRange = Application.WorksheetFunction.CountA(rng)

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

