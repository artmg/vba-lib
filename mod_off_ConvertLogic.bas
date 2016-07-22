Attribute VB_Name = "mod_off_ConvertLogic"
Option Explicit

' error handling tag             **********************
Const cStrModuleName As String = "mod_off_ConvertLogic"
'                                **********************

' General data conversions and simple logical evaluations

'
'  160721.AMG  moved bTestMatch in from mod_exc_WbkShtRngName
'  150519.AMG  created with strBytesReversed
'

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
' vba-lib / none
'

' IMPROVEMENTS
' ============
'
' * none as
'




Function strBytesReversed( _
    strInput As String _
) As String

    ' accept an Input String which is assumed to contain a string representation of hexadecimal bytes
    ' break it into individual bytes (i.e. two characters at a time)
    ' and build them back up in reverse
    '
    ' this is a "two at a time" equivalent to VBA.Strings.StrReverse

    Dim strOriginal, strByte, strResult As String
    Dim iCount As Integer

    ' pad if Input string is not whole number of bytes
    If Len(strInput) Mod 2 = 1 Then
        strOriginal = "0" & strInput
    Else
        strOriginal = strInput
    End If
    strResult = ""

    ' credit http://www.pcreview.co.uk/threads/byte-order-reversal-in-spreadsheet.1006377/
    For iCount = (Len(strOriginal) - 1) To 1 Step -2
        strByte = Mid(strOriginal, iCount, 2)
        strResult = strResult & strByte
    Next iCount

    strBytesReversed = strResult

End Function


' moved in from mod_exc_WbkShtRngName
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






