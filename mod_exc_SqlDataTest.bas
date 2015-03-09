Attribute VB_Name = "mod_exc_SqlDataTest"
' References
' ==========
'
' This module uses the following references (paths and GUIDs may vary)
'
' Microsoft ActiveX Data Objects 2.8 Library (C:\Program Files\Common Files\System\ado\msado15.dll) {2A75196C-D9EB-4129-B803-931327F72D5C}
Option Explicit

Function SqlDatabaseTest(strServer As String, strDb As String, strCmd As String) As String
    Application.Volatile (True)
    
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim strResult As String

On Error GoTo ErrorHandler
    Set cnn = New ADODB.Connection
    'Set cmd = New ADODB.Command
    Set rst = New ADODB.Recordset
    

    cnn.Open _
        "Provider=sqloledb;" _
        & "Data Source=" & strServer & ";" _
        & "Initial Catalog=" & strDb & ";" _
        & "Integrated Security=SSPI;"

'    cmd.CommandType = adCmdText
'    cmd.CommandText = StrCmd
'    cmd.ActiveConnection = cnn
'
'    rst = cmd.Execute

    Set rst = cnn.Execute(strCmd)
    strResult = CStr(rst(0))
    
ErrorHandler:
    Select Case Err.Number
        Case 0:
            SqlDatabaseTest = "ALLOWED (" & strResult & ")"
        Case -2147467259: ' 0x80004005
            SqlDatabaseTest = "DENIED! (Database not found, or insufficient permissions)"
        Case -2147217843:
            SqlDatabaseTest = "DENIED! (Database does not recognise user)"
        Case Else:
            SqlDatabaseTest = "DENIED! (" & Err.Number & " - " & Err.Description & ")"
    End Select

    If cnn.State = adStateOpen Then cnn.Close
    Set rst = Nothing
    'Set cmd = Nothing
    Set cnn = Nothing

End Function













