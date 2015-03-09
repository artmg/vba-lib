Attribute VB_Name = "mod_exc_PivotTables"
' mod_exc_PivotTables

' miscellaneous Pivot Table manipulation routines
'  130910.AMG  added error handling
'  130909.AMG  refactored more functions in for simpler calling code, and added formatting features
'  130906.AMG  created

Option Explicit

' error handling tag
Const cStrModuleName As String = "mod_exc_PivotTables"


Public Enum eNumberFormat
    [_First] = 1
    General = 1
    DecimalNumber
    IntegerNumber
    CurrencyValue
    [_Last]
End Enum

Function PivotTableAddOnNewSheet _
(ByRef wbk As Workbook _
, ByRef cch As PivotCache _
, ByVal strPivotSheetName As String _
, ByVal lIndex As Long _
, Optional lRow As Long = 3 _
) As PivotTable

''' standard procedure error handler begin initialise 130910.AMG '''
Application.DisplayAlerts = False ' DoCmd.SetWarnings False
On Error GoTo ErrorHandler
Dim strErrorObject As String
Const cStrProcedureName As String = "PivotTableAddOnNewSheet"
strErrorObject = strPivotSheetName
''' standard procedure error handler end initialise '''

    Dim sht As Worksheet
    Dim pvt As PivotTable
    
    ' create the sheet and the table
    Set sht = wbk.Sheets.Add
    sht.Name = strPivotSheetName
    
    ' position the sheet
    ' if index parameter is less then 1 then it is relative to LAST existing worksheet...
    Dim lAbsolute As Long
    If lIndex > 0 Then
        lAbsolute = lIndex
    Else
        lAbsolute = wbk.Worksheets.Count + lIndex
    End If
    sht.Move after:=wbk.Worksheets(lAbsolute)
    
    ' create pivot
    Set pvt = cch.CreatePivotTable( _
        TableDestination:=wbk.Worksheets(strPivotSheetName).Cells(lRow, 1), _
        TableName:=strPivotSheetName)

    ' return object
    Set PivotTableAddOnNewSheet = pvt

''' standard procedure error handler begin terminate 130910.AMG '''
Proc_Exit:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
Exit Function
ErrorHandler:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
LogException "ERROR", "Trying to '" & cStrModuleName & "." & cStrProcedureName & "' (" & _
strErrorObject & ") - " & "Error " & Err.Number & vbCrLf & " - """ & Err.Description & """"
Err.Raise (Err.Number)
Resume Proc_Exit
''' standard procedure error handler end terminate '''
End Function


Function PivotHideValue _
(pvt As Object _
, strFieldName As String _
, strCheckValue As String _
) As Boolean
''' standard procedure error handler begin initialise 130910.AMG '''
Application.DisplayAlerts = False ' DoCmd.SetWarnings False
On Error GoTo ErrorHandler
Dim strErrorObject As String
Const cStrProcedureName As String = "PivotHideValue"
strErrorObject = "Pivot: " & pvt.Name & " Field: " & strFieldName & " Value: " & strCheckValue
''' standard procedure error handler end initialise '''

    Dim i As Integer
    For i = 1 To pvt.PivotFields(strFieldName).PivotItems.Count
        If strCheckValue = pvt.PivotFields(strFieldName).PivotItems(i).Name Then
            pvt.PivotFields(strFieldName).PivotItems(strCheckValue).Visible = False
            Exit Function
        End If
    Next i
    ' if no such value is found, simply exit (rather than throw an error)

''' standard procedure error handler begin terminate 130910.AMG '''
Proc_Exit:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
Exit Function
ErrorHandler:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
LogException "ERROR", "Trying to '" & cStrModuleName & "." & cStrProcedureName & "' (" & _
strErrorObject & ") - " & "Error " & Err.Number & vbCrLf & " - """ & Err.Description & """"
Err.Raise (Err.Number)
Resume Proc_Exit
''' standard procedure error handler end terminate '''
End Function


Function PivotHideAllValuesExcept _
(pvt As PivotTable _
, strFieldName As String _
, strCheckValue As String _
)
''' standard procedure error handler begin initialise 130910.AMG '''
Application.DisplayAlerts = False ' DoCmd.SetWarnings False
On Error GoTo ErrorHandler
Dim strErrorObject As String
Const cStrProcedureName As String = "PivotHideAllValuesExcept"
strErrorObject = "Pivot: " & pvt.Name & " Fields: " & strFieldName & " Value: " & strCheckValue
''' standard procedure error handler end initialise '''

'    Dim i As Integer
'    For i = 1 To pvt.PivotFields(strFieldName).PivotItems.Count
'        If strCheckValue = pvt.PivotFields(strFieldName).PivotItems(i).Name Then
'            pvt.PivotFields(strFieldName).PivotItems(strCheckValue).Visible = False
    
    Dim strFullCheck, strRootError As String
    Dim pi As PivotItem
    
    Const cStrDelim As String = ","
    ' strCheckValues is a comma delimited list, so make it easy to search by topping and tailing it
    strFullCheck = cStrDelim & strCheckValue & cStrDelim
    strRootError = strErrorObject
    For Each pi In pvt.PivotFields(strFieldName).PivotItems
        If InStr(strFullCheck, cStrDelim & pi.Name & cStrDelim) > 0 Then
            strErrorObject = strRootError & " setting " & pi.Name & " visible"
            pi.Visible = True
        Else
            strErrorObject = strRootError & " setting " & pi.Name & " invisible"
            pi.Visible = False
        End If
    Next
'    Next i
    ' if no such value is found, simply exit (rather than throw an error)
    
''' standard procedure error handler begin terminate 130910.AMG '''
Proc_Exit:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
Exit Function
ErrorHandler:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
LogException "ERROR", "Trying to '" & cStrModuleName & "." & cStrProcedureName & "' (" & _
strErrorObject & ") - " & "Error " & Err.Number & vbCrLf & " - """ & Err.Description & """"
Err.Raise (Err.Number)
Resume Proc_Exit
''' standard procedure error handler end terminate '''
End Function


Function PivotFieldFilter _
(fld As PivotField _
, Value _
) As Boolean
''' standard procedure error handler begin initialise 130910.AMG '''
Application.DisplayAlerts = False ' DoCmd.SetWarnings False
On Error GoTo ErrorHandler
Dim strErrorObject As String
Const cStrProcedureName As String = "PivotFieldFilter"
strErrorObject = "PivotField: " & fld.Name & " Value: " & CStr(Value)
''' standard procedure error handler end initialise '''

    Application.ScreenUpdating = False
    With fld
        If .Orientation = xlPageField Then
            .CurrentPage = Value
        ElseIf .Orientation = xlRowField Or .Orientation = xlColumnField Then
            Dim i As Long
            On Error Resume Next ' Needed to avoid getting errors when manipulating fields that were deleted from the data source.
            ' Set first item to Visible to avoid getting no visible items while working
            .PivotItems(1).Visible = True
            For i = 2 To fld.PivotItems.Count
                If .PivotItems(i).Name = Value Then _
                    .PivotItems(i).Visible = True Else _
                    .PivotItems(i).Visible = False
            Next i
            If .PivotItems(1).Name = Value Then _
                .PivotItems(1).Visible = True Else _
                .PivotItems(1).Visible = False
        End If
    End With
    Application.ScreenUpdating = True

''' standard procedure error handler begin terminate 130910.AMG '''
Proc_Exit:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
Exit Function
ErrorHandler:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
LogException "ERROR", "Trying to '" & cStrModuleName & "." & cStrProcedureName & "' (" & _
strErrorObject & ") - " & "Error " & Err.Number & vbCrLf & " - """ & Err.Description & """"
Err.Raise (Err.Number)
Resume Proc_Exit
''' standard procedure error handler end terminate '''
End Function

'Function PivotFieldSetAsFirstRow _
'(fld As PivotField _
')
'''' standard procedure error handler begin initialise 130910.AMG '''
'Application.DisplayAlerts = False ' DoCmd.SetWarnings False
'On Error GoTo ErrorHandler
'Dim strErrorObject As String
'Const cStrProcedureName As String = "PivotFieldFilter"
'strErrorObject = "PivotField: " & fld.Name & " Value: " & CStr(Value)
'''' standard procedure error handler end initialise '''
'
'    Application.ScreenUpdating = False
'    With fld
'        .Orientation = xlRowField
'        .Position = 1
'    End With
'    Application.ScreenUpdating = True
'
'''' standard procedure error handler begin terminate 130910.AMG '''
'Proc_Exit:
'Application.DisplayAlerts = True ' DoCmd.SetWarnings True
'Exit Function
'ErrorHandler:
'Application.DisplayAlerts = True ' DoCmd.SetWarnings True
'LogException "ERROR", "Trying to '" & cStrModuleName & "." & cStrProcedureName & "' (" & _
'strErrorObject & ") - " & "Error " & Err.Number & vbCrLf & " - """ & Err.Description & """"
'Err.Raise (Err.Number)
'Resume Proc_Exit
'''' standard procedure error handler end terminate '''
'End Function

Function PivotSetFieldAsFirstRow _
(pvt As PivotTable _
, strFieldName As String _
)
''' standard procedure error handler begin initialise 130910.AMG '''
Application.DisplayAlerts = False ' DoCmd.SetWarnings False
On Error GoTo ErrorHandler
Dim strErrorObject As String
Const cStrProcedureName As String = "PivotSetFieldAsFirstRow"
strErrorObject = "Pivot: " & pvt.Name & " Field: " & strFieldName
''' standard procedure error handler end initialise '''

    Application.ScreenUpdating = False
    With pvt.PivotFields(strFieldName)
        .Orientation = xlRowField
        .Position = 1
    End With
    Application.ScreenUpdating = True

''' standard procedure error handler begin terminate 130910.AMG '''
Proc_Exit:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
Exit Function
ErrorHandler:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
LogException "ERROR", "Trying to '" & cStrModuleName & "." & cStrProcedureName & "' (" & _
strErrorObject & ") - " & "Error " & Err.Number & vbCrLf & " - """ & Err.Description & """"
Err.Raise (Err.Number)
Resume Proc_Exit
''' standard procedure error handler end terminate '''
End Function

Function PivotRowFieldFormat _
(ByRef pvt As PivotTable _
, ByVal strFieldName As String _
, ByVal lFormat As eNumberFormat _
)
''' standard procedure error handler begin initialise 130910.AMG '''
Application.DisplayAlerts = False ' DoCmd.SetWarnings False
On Error GoTo ErrorHandler
Dim strErrorObject As String
Const cStrProcedureName As String = "PivotRowFieldFormat"
strErrorObject = "Pivot: " & pvt.Name & " Field: " & strFieldName
''' standard procedure error handler end initialise '''

    Application.ScreenUpdating = False
    Dim rng As Range
    Set rng = pvt.PivotFields(strFieldName).DataRange
    Select Case lFormat
        Case eNumberFormat.General:
            rng.NumberFormat = "General"
        Case eNumberFormat.DecimalNumber:
            rng.NumberFormat = "0.00"
        Case eNumberFormat.IntegerNumber:
           rng.NumberFormat = "0"
        Case eNumberFormat.CurrencyValue:
           rng.NumberFormat = "£#,##0"
        Case Else:
    End Select
    Application.ScreenUpdating = True

''' standard procedure error handler begin terminate 130910.AMG '''
Proc_Exit:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
Exit Function
ErrorHandler:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
LogException "ERROR", "Trying to '" & cStrModuleName & "." & cStrProcedureName & "' (" & _
strErrorObject & ") - " & "Error " & Err.Number & vbCrLf & " - """ & Err.Description & """"
Err.Raise (Err.Number)
Resume Proc_Exit
''' standard procedure error handler end terminate '''
End Function

' It was NOT valid to set number format for a rowfield :(
' "You can set the NumberFormat property only for a data field"
' quote > http://msdn.microsoft.com/en-us/library/office/bb237338(v=office.12).aspx
' so set it after the pivot is complete
'Function PivotSetFieldAsFirstRow _
'    (pvt As PivotTable _
'    , strFieldName As String _
'    , Optional ByVal lFormat As eNumberFormat _
'    )
'    Application.ScreenUpdating = False
'    With pvt.PivotFields(strFieldName)
'        .Orientation = xlRowField
'        .Position = 1
'        If Not IsMissing(lFormat) Then
'            Select Case lFormat
'                Case eNumberFormat.General:
'                    .NumberFormat = "General"
'                Case eNumberFormat.DecimalNumber:
'                    .NumberFormat = "0.00"
'                Case eNumberFormat.IntegerNumber:
'                   .NumberFormat = "0"
'                Case eNumberFormat.CurrencyValue:
'                   .NumberFormat = "£#,##0"
'                Case Else:
'            End Select
'        End If
'    End With
'    Application.ScreenUpdating = True
'End Function

Function PivotSetFieldAsTotal _
(ByRef pvt As PivotTable _
, ByVal strFieldName As String _
, Optional ByVal strCaption As String _
, Optional lFormat As eNumberFormat _
)
''' standard procedure error handler begin initialise 130910.AMG '''
Application.DisplayAlerts = False ' DoCmd.SetWarnings False
On Error GoTo ErrorHandler
Dim strErrorObject As String
Const cStrProcedureName As String = "PivotSetFieldAsTotal"
strErrorObject = "Pivot: " & pvt.Name & " Field: " & strFieldName
''' standard procedure error handler end initialise '''


    Application.ScreenUpdating = False
        With pvt.PivotFields(strFieldName)
            .Orientation = xlDataField
            .Function = xlSum
            ' Summing the field changes the caption, so either set it as specified,
            ' or simply revert to the original field name
            If IsMissing(strCaption) Then
                .Caption = strFieldName
            Else
                .Caption = strCaption
            End If
            If Not IsMissing(lFormat) Then
                Select Case lFormat
                    Case eNumberFormat.General:
                        .NumberFormat = "General"
                    Case eNumberFormat.DecimalNumber:
                        .NumberFormat = "0.00"
                    Case eNumberFormat.IntegerNumber:
                       .NumberFormat = "0"
                    Case eNumberFormat.CurrencyValue:
                       .NumberFormat = "£#,##0"
                    Case Else:
                End Select
            End If
        End With
    Application.ScreenUpdating = True

''' standard procedure error handler begin terminate 130910.AMG '''
Proc_Exit:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
Exit Function
ErrorHandler:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
LogException "ERROR", "Trying to '" & cStrModuleName & "." & cStrProcedureName & "' (" & _
strErrorObject & ") - " & "Error " & Err.Number & vbCrLf & " - """ & Err.Description & """"
Err.Raise (Err.Number)
Resume Proc_Exit
''' standard procedure error handler end terminate '''
End Function

Function PivotShowSubtotalsOnFields _
(pvt As PivotTable _
, strFieldNames As String _
)
''' standard procedure error handler begin initialise 130910.AMG '''
Application.DisplayAlerts = False ' DoCmd.SetWarnings False
On Error GoTo ErrorHandler
Dim strErrorObject As String
Const cStrProcedureName As String = "PivotShowSubtotalsOnFields"
strErrorObject = "Pivot: " & pvt.Name & " Fields: " & strFieldNames
''' standard procedure error handler end initialise '''

' accepts a comma-delimited list of Field Names
Const cStrDelim As String = ","
' all fields NOT in list will have their subtotals turned off
' e.g.
'     PivotShowSubtotalsOnFields pvt, "Phase,Class"

    Application.ScreenUpdating = False

    Dim fld As PivotField
    ' Make the delimited list easy to search by topping and tailing it
    Dim strDelimitedList, strNameWithDelims As String
    strDelimitedList = cStrDelim & strFieldNames & cStrDelim
    For Each fld In pvt.PivotFields
        strNameWithDelims = cStrDelim & fld.Name & cStrDelim
        If InStr(1, strNameWithDelims, strDelimitedList, vbTextCompare) Then
            fld.Subtotals(1) = True
        Else
            fld.Subtotals(1) = False
        End If
    Next fld
    Application.ScreenUpdating = True

''' standard procedure error handler begin terminate 130910.AMG '''
Proc_Exit:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
Exit Function
ErrorHandler:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
LogException "ERROR", "Trying to '" & cStrModuleName & "." & cStrProcedureName & "' (" & _
strErrorObject & ") - " & "Error " & Err.Number & vbCrLf & " - """ & Err.Description & """"
Err.Raise (Err.Number)
Resume Proc_Exit
''' standard procedure error handler end terminate '''
End Function


Function PivotColumnFormat _
(pvt As PivotTable _
, strFieldName As String _
, Optional dWidth As Double = 0 _
, Optional bWrap As Boolean = False _
)
''' standard procedure error handler begin initialise 130910.AMG '''
Application.DisplayAlerts = False ' DoCmd.SetWarnings False
On Error GoTo ErrorHandler
Dim strErrorObject As String
Const cStrProcedureName As String = "PivotColumnFormat"
strErrorObject = "Pivot: " & pvt.Name & " Field: " & strFieldName
''' standard procedure error handler end initialise '''

    ' credit > useful visual reference for accessing parts of pivot table
    ' and how to use the valuable intersect
    ' http://peltiertech.com/WordPress/referencing-pivot-table-ranges-in-vba/

    Dim rng As Range
    Set rng = pvt.PivotFields(strFieldName).DataRange
    If dWidth <> 0 Then
        rng.ColumnWidth = dWidth
    End If
    If bWrap Then
        rng.WrapText = True
    End If

''' standard procedure error handler begin terminate 130910.AMG '''
Proc_Exit:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
Exit Function
ErrorHandler:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
LogException "ERROR", "Trying to '" & cStrModuleName & "." & cStrProcedureName & "' (" & _
strErrorObject & ") - " & "Error " & Err.Number & vbCrLf & " - """ & Err.Description & """"
Err.Raise (Err.Number)
Resume Proc_Exit
''' standard procedure error handler end terminate '''
End Function

Function PivotColumnsAutoFit _
(pvt As PivotTable _
, Optional strFieldNames As String _
)
''' standard procedure error handler begin initialise 130910.AMG '''
Application.DisplayAlerts = False ' DoCmd.SetWarnings False
On Error GoTo ErrorHandler
Dim strErrorObject As String
Const cStrProcedureName As String = "PivotColumnsAutoFit"
strErrorObject = "Pivot: " & pvt.Name & " Fields: " & CStr(strFieldNames)
''' standard procedure error handler end initialise '''

' Sets Column Widths to AutoFit
' accepts a comma-delimited list of Field Names
Const cStrDelim As String = ","
' e.g.
'     PivotColumnsAutoFit pvt, "Phase,Class"
'
' If no names are specified then ALL columns will be autofitted

    Application.ScreenUpdating = False

    Dim fld As PivotField
    Dim strDelimitedList, strNameWithDelims As String
    ' Make the delimited list easy to search by topping and tailing it
    strDelimitedList = cStrDelim & strFieldNames & cStrDelim
    For Each fld In pvt.PivotFields
        
        ' we autofit the column is EITHER no list was specified...
        Dim bAutoFitThisColumn As Boolean
        bAutoFitThisColumn = IsMissing(strFieldNames)
        ' OR if this field matches in the list
        If Not bAutoFitThisColumn Then
            strNameWithDelims = cStrDelim & fld.Name & cStrDelim
            bAutoFitThisColumn = InStr(1, strNameWithDelims, strDelimitedList, vbTextCompare)
        End If

        If bAutoFitThisColumn Then
            fld.DataRange.AutoFit = True
        End If
    Next fld
    Application.ScreenUpdating = True

''' standard procedure error handler begin terminate 130910.AMG '''
Proc_Exit:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
Exit Function
ErrorHandler:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
LogException "ERROR", "Trying to '" & cStrModuleName & "." & cStrProcedureName & "' (" & _
strErrorObject & ") - " & "Error " & Err.Number & vbCrLf & " - """ & Err.Description & """"
Err.Raise (Err.Number)
Resume Proc_Exit
''' standard procedure error handler end terminate '''
End Function


Function PivotTableFormat( _
  pvt As PivotTable _
, Optional strTableStyle As String _
, Optional lThemeColor As XlThemeColor _
, Optional vTintAndShade As Variant _
)
''' standard procedure error handler begin initialise 130910.AMG '''
Application.DisplayAlerts = False ' DoCmd.SetWarnings False
On Error GoTo ErrorHandler
Dim strErrorObject As String
Const cStrProcedureName As String = "PivotTableFormat"
strErrorObject = "Pivot: " & pvt.Name & " Style: " & strTableStyle
''' standard procedure error handler end initialise '''

    Dim sht As Worksheet
    Set sht = pvt.Parent

    If Not IsMissing(strTableStyle) Then
        pvt.TableStyle2 = strTableStyle
    End If
    If Not IsMissing(lThemeColor) Then
        sht.Tab.ThemeColor = lThemeColor
    End If
    If Not IsMissing(vTintAndShade) Then
        sht.Tab.TintAndShade = vTintAndShade
    End If
    ' move the selection off the pivot to hide the wizard
    ' may need to calc number of page fields to choose correct row
    sht.Cells(3, 1).Select

''' standard procedure error handler begin terminate 130910.AMG '''
Proc_Exit:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
Exit Function
ErrorHandler:
Application.DisplayAlerts = True ' DoCmd.SetWarnings True
LogException "ERROR", "Trying to '" & cStrModuleName & "." & cStrProcedureName & "' (" & _
strErrorObject & ") - " & "Error " & Err.Number & vbCrLf & " - """ & Err.Description & """"
Err.Raise (Err.Number)
Resume Proc_Exit
''' standard procedure error handler end terminate '''
End Function

