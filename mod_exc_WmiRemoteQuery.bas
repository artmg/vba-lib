Attribute VB_Name = "mod_exc_WmiRemoteQuery"
Option Explicit
 
' this was built to be used by mod_exc_ADGroupsUsers subs and funcs
Function GetWMIRemoteAdminsGroupMembers(ByVal strComputerName As String) As String
    Dim strResultList As String
   
    Application.StatusBar = strComputerName
 
    ' Not sure it uses these LDAP objects here
    ' the other functions use it merely to obtain the Root Domain Name
    ' perhaps that must be instatiated (as ADSI provider???) before WMI query can work?
'    Dim objRootLDAP As DirectoryEntry
'    Dim strRootDomain As String
 '   Set objRootLDAP = GetObject("LDAP://rootDse")
 '   strRootDomain = objDomain.Get("dnsHostName")
    ' or "rootDomainNamingContext" property?
 
 
    ' This uses simple GetObject object management, like VBS style - brute but effective
    ' vba syntax credit > http://www.vbforums.com/showthread.php?p=4001522
    Dim objWmiService, objWmiResults, objWmiResult
'    On Error Resume Next
'    Set objWmiService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2")    ' impersonate?
    Set objWmiService = GetObject("winmgmts:\\" & strComputerName & "\root\cimv2")    ' impersonate?
    ' query credit > http://www.itworld.com/nlswindows071016windows
    Dim strWmiQuery: strWmiQuery = _
          " Select * " _
        & " From Win32_GroupUser " _
        & " Where GroupComponent=""win32_group.name=\""administrators\"",domain=\""" _
        & strComputerName _
        & "\"""""
 
    Set objWmiResults = objWmiService.ExecQuery(strWmiQuery)
   
' without the object existing in the calling function it can't enumerate the instances
    ' so we serialise it as a string, making it possible for the caller to parse
    ' in VBA it's actually far more efficient to extend a string than it is to perpetually redim an array!
    For Each objWmiResult In objWmiResults
        strResultList = strResultList & "," & objWmiResult.PartComponent ' this is the specific property required
    Next
    ' strip off the leading comma as we return the function's result
    GetWMIRemoteAdminsGroupMembers = Mid$(strResultList, 2)
 
    Application.StatusBar = False
 
End Function
