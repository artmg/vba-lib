Attribute VB_Name = "mod_exc_DataTables"

' mod_exc_DataTables
'
' 150326.AMG  added Table Match functions previously in Consol module
' 150312.AMG  created with table creation


Option Explicit

' Define match types
Public Enum enumDataTableMatchType
    MatchCaseSens
    MatchCaseInsens
End Enum


'
' *** TABLE CREATION *****************************
'

Public Function FillTableDownByCopyingBlanksFromAbove()
    Dim rng, rRow, rCel As Range
    Set rng = ActiveSheet.UsedRange
    For Each rRow In rng.Rows
        For Each rCel In rRow.Cells
            With rCel
                If (.Value = "") And (.row > 1) Then
                    .Value = rng.Cells(.row - 1, .Column).Value
                End If
            End With
        Next
    Next
End Function

Public Function FillColumnDownByCopyingBlanksFromAbove()
    Dim rCel As Range
    For Each rCel In ActiveSheet.Columns(ActiveCell.Column).Cells
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
    
    If enumMatchType Then
        strKeyToMatch = UCase(strUnprepared)
        strToReplace = UCase(strIgnore)
    Else
        strToReplace = strIgnore
    End If
    
' ONLY USE IGNORE DEPENDING ON MATCH TYPE FLAG
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


