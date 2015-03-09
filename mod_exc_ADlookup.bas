Attribute VB_Name = "mod_exc_ADlookup"
' References
' ==========
'
' This module uses the following references (paths and GUIDs may vary)
'
' Active DS Type Library (C:\WINDOWS\system32\activeds.tlb) {97D25DB0-0363-11CF-ABC4-02608C9E7553}
' Microsoft ActiveX Data Objects 2.8 Library (C:\Program Files\Common Files\System\ADO\msado15.dll) {2A75196C-D9EB-4129-B803-931327F72D5C}

Option Explicit

Public Enum AdObjectType
    AdUser
    AdComputer
End Enum

' credit > http://www.freevbcode.com/ShowCode.Asp?ID=710
Public Function GetAdAttributeFrom(strObjectName As String, strAttributeName As String, Optional objType As AdObjectType = AdObjectType.AdUser) As String

Dim oRoot As ActiveDs.IADs
Dim oDomain As ActiveDs.IADs

Dim strQuery As String
Dim strValue As String

On Error GoTo ErrHandler:

'Get user Using LDAP/ADO.  There is an easier way
'to bind to a user object using the WinNT provider,
'but this way is a better for educational purposes
Set oRoot = GetObject("LDAP://rootDSE")
''''''''work in the default domain
' sDomain = oRoot.Get("defaultNamingContext")
' Set oDomain = GetObject("LDAP://" & sDomain)
Set oDomain = GetObject("LDAP://" & oRoot.Get("defaultNamingContext"))

' add the base to the query
strQuery = "<" & oDomain.ADsPath & ">;"

Select Case objType
    Case AdObjectType.AdUser
        strQuery = strQuery _
            & "(&" _
            & "(objectCategory=person)" _
            & "(objectClass=user)" _
            & "(sAMAccountName=" & strObjectName & ")" _
            & ");"
    Case AdObjectType.AdComputer
        strQuery = strQuery _
            & "(&" _
            & "(objectCategory=computer)" _
            & "(objectClass=computer)" _
            & "(name=" & strObjectName & ")" _
            & ");"
End Select

' Add the attribute name
strQuery = strQuery & strAttributeName & ";"
' this was for the old object-based version
'strQuery = strQuery & "adsPath;"

' set the query depth to check the whole domain tree
strQuery = strQuery & "subTree"



Dim cnn As New ADODB.Connection
Dim rst As ADODB.Recordset

cnn.Open "Data Source=Active Directory Provider;Provider=ADsDSOObject"
  
Set rst = cnn.Execute(strQuery)

If Not rst.EOF Then
On Error Resume Next
    strValue = rst(0)
    If strValue = "" Then strValue = rst(0).Value(0)
    ' this was the old object based version
'    Dim user As ActiveDs.IADsUser
'    Set user = GetObject(rst("adsPath"))
'    strValue = user.ADsPath
End If

GetAdAttributeFrom = strValue


' cleanup
ErrHandler:
On Error Resume Next
If Not rst Is Nothing Then
    If rst.State <> 0 Then rst.Close
    Set rst = Nothing
End If

If Not cnn Is Nothing Then
    If cnn.State <> 0 Then cnn.Close
    Set cnn = Nothing
End If

Set oRoot = Nothing
Set oDomain = Nothing

End Function
