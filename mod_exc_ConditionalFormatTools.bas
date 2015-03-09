Attribute VB_Name = "mod_exc_ConditionalFormatTools"
'
' mod_exc_ConditionalFormatTools
'

' 130415.AMG created with DebugPrintConditions

Sub DebugPrintConditions()
    With ActiveCell
        Dim cond As FormatCondition
        For Each cond In .FormatConditions
            Debug.Print "Formula1=[" & cond.Formula1 & "] Type=[" & cond.Type & "] Interior.Color=[" & cond.Interior.Color & "] Font.Color=[" & cond.Font.Color & "]"
        Next
    End With
End Sub


Sub SetConditionsForWorkAndEffort()
' there _might_ have been a way to use xlColorScale, but this probably requires a value in the cell

' Const cstrCondID As String = "WorkAndEffort" ' id to recognise items we have added
Const ciLeaveFirstFormats As Integer = 7

Const cdShadeLimit1 As Double = 0.7
Const cdShadeDepth1 As Double = -0.6
Const cdShadeLimit2 As Double = 0.4
Const cdShadeDepth2 As Double = 0
Const cdShadeLimit3 As Double = 0.2
Const cdShadeDepth3 As Double = 0.4

Const clThemeColor As Long = 8

Dim strShadeEval, strColorEval, strAppliesTo As String

' =IF(AND(M$1>=$H12,M$1<=$I12),($I12-$H12)/$G12,0)
' M = column, 1 = date in heading
' H = start column I = end
' G = duration

'strShadeEval = "$G1/($I1-$H1)"
strShadeEval = "$J1/NETWORKDAYS($K1,$L1)"
strColorEval = "AND(P$1>=$K1,P$1<=$L1)"
strAppliesTo = "=$P:$NP"

    Dim rng As Range
    
    Set rng = ActiveSheet.Range(strAppliesTo)
    ' remove old ones
'    Dim cond As FormatCondition
'    For Each cond In rng.FormatConditions
'        On Error Resume Next
'        If CStr(cond.Text) = cstrCondID Then
'            cond.Delete
'        End If
'        On Error GoTo 0
'    Next

'    Dim cnt As Integer
'    If rng.FormatConditions.Count > ciLeaveFirstFormats Then
'        For cnt = rng.FormatConditions.Count To (ciLeaveFirstFormats + 1) Step -1
'            rng.FormatConditions(cnt).Delete
'        Next
'    End If

' colour
'        With .FormatConditions.Add(xlExpression, , "=IF(" & strColorEval & ",1,0)")
'            .Interior.ThemeColor = 4
'            .StopIfTrue = False
'        End With
    ' select the range to avoid excel doing relative conversion of formulae
    rng.Select

    With rng.FormatConditions.Add(xlExpression, , "=IF(AND(" & strColorEval & "," & strShadeEval & ">" & CStr(cdShadeLimit1) & "),1,0)")
        .Interior.ThemeColor = clThemeColor
        .Interior.TintAndShade = cdShadeDepth1
        .StopIfTrue = True
'        .Text = cstrCondID
    End With
    With rng.FormatConditions.Add(xlExpression, , "=IF(AND(" & strColorEval & "," & strShadeEval & ">" & CStr(cdShadeLimit2) & "),1,0)")
        .Interior.ThemeColor = clThemeColor
        .Interior.TintAndShade = cdShadeDepth2
        .StopIfTrue = True
'        .Text = cstrCondID
    End With
    With rng.FormatConditions.Add(xlExpression, , "=IF(AND(" & strColorEval & "," & strShadeEval & ">" & CStr(cdShadeLimit3) & "),1,0)")
        .Interior.ThemeColor = clThemeColor
        .Interior.TintAndShade = cdShadeDepth3
        .StopIfTrue = True
'        .Text = cstrCondID
    End With
End Sub
