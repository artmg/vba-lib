Attribute VB_Name = "mod_exc_ErrorHandling"
' ACCESS ONLY ' Option Compare Database
Option Explicit

' error handling tag
Const cStrModuleName As String = "mod_exc_ErrorHandling"

'
' To use the Error Handling...
'
' Un/Comment lines relating to which Office application you run this from
' (e.g. ACCESS or EXCEL)
'
' At the beginning of each MODULE...
' copy the first four lines to each module you want to use this error handler
' and copy the Module Name from the Properties Pane into the string

' At the begining of each PROCEDURE paste in the Initiatise section from below
' paste the Function Name in and choose the most important object passed to paste

' at the end of each PROCEDURE paste in the Terminate section as is

'  130918.AMG  adapted from mod_acc_ErrorHandling
'  130910.AMG  created
'

Private Function StandardErrorCode(varObjectPassed As Variant)

''' standard Access procedure error handler begin initialise 130818.AMG '''
UseStandardErrorHandling False
On Error GoTo ErrorHandler
Dim strErrorObject As String
Const cStrProcedureName As String = "PASTEfunctionNameHERE"
strErrorObject = CStr(varObjectPassed)
''' standard procedure error handler end initialise '''



    ' The main body of your function goes in here
    '
    '


''' standard Access procedure error handler begin terminate 130818.AMG '''
Proc_Exit:
UseStandardErrorHandling True
Exit Function
ErrorHandler:
UseStandardErrorHandling True
HandleErrorWithMessage "Error trying to '" & cStrModuleName & "." & cStrProcedureName & "' (" & _
    strErrorObject & ") " & vbCrLf & vbCrLf & _
    "Error " & Err.Number & vbCrLf & """" & Err.Description & """"
Resume Proc_Exit
''' standard procedure error handler end terminate '''

End Function

Public Function UseStandardErrorHandling _
(bUseItOrNot As Boolean _
)

' ACCESS ONLY ' DoCmd.SetWarnings bUseItOrNot
' EXCEL ONLY ' Application.DisplayAlerts = bUseItOrNot
Application.DisplayAlerts = bUseItOrNot

End Function

Public Function HandleErrorWithMessage(strMessage As String)
    ' develop options for changing delimeters between newlines and pipes
    MsgBox strMessage
    ' can we have a Stop and Debug option?
    Err.Raise (Err.Number)
End Function

