Attribute VB_Name = "mod_exc_ADGroupsUsers"
Option Explicit
Option Base 0

' mod_exc_ADGroupsUsers.bas

' 130620.AMG added Prepare_Workbook function to make it self-evident how to use the module
' 111120.AMG created

Sub Prepare_Workbook_For_AD_Query()
    With Application.Workbooks.Add
        .Worksheets(1).Name = "Query"
        .Worksheets(2).Name = "Results"
        .Worksheets(3).Delete
        .Worksheets(1).Cells(1, 1).Value = "QueryType"
        .Worksheets(1).Cells(2, 1).Value = "Location (e.g. Domain)"
        .Worksheets(1).Cells(3, 1).Value = "Item"
        .Worksheets(1).Cells(4, 1).Value = "More item(s)"
        .Worksheets(1).Cells(5, 1).Value = "..."
        .Worksheets(1).Cells(1, 2).Validation.Add _
            Type:=xlValidateList _
            , Operator:=xlEqual _
            , Formula1:="ListUsers,ListGroups,ComputerGroups"
        .Activate
    End With
End Sub


Sub Execute_AD_Query_And_Overwrite_Results()
'    Dim wksQuery, wksResults As Worksheet
'    Set wksQuery = ActiveWorkbook.Sheets("Query")
'    Set wksResults = ActiveWorkbook.Sheets("Results")
    With ActiveWorkbook.Sheets("Query")
        QueryEachItemAndOutputResults _
              wksResults:=ActiveWorkbook.Sheets("Results") _
            , strQueryType:=.Cells(1, 2) _
            , strLocation:=.Cells(2, 2) _
            , rngItems:=.Range("B3:B1000") _
 
    End With
End Sub
 
Function QueryEachItemAndOutputResults _
    (ByRef wksResults As Worksheet _
    , ByVal strQueryType As String _
    , ByVal strLocation As String _
    , ByVal rngItems As Range _
    )
 
    Dim varItem, varResults
    Dim strResultList As String
    Dim lngOutputRow, lngRow, lngCol As Long
'    Dim lngIndex As Long
    Dim intColItems, intColResults As Integer
 
    Dim strColumns() As String
 
    ' prepare sheet and add titles
    With wksResults
        .Cells.Clear
        .Range("A1").Value = "Item"
        .Range("B1").Value = "Group"
       
        lngOutputRow = 2
   
        For Each varItem In rngItems
            ' get member list as a comma separated list
            If varItem <> "" Then
                ' determine type of query
                ' obtain results and split into an array
                ' then determine which column to output what
                Select Case strQueryType
                    Case "ListUsers":
                        varResults = Split(GetGroupUsers(varItem), ",")
                        intColItems = 2
                        intColResults = 1
                    Case "ListGroups":
                        'strColumns = GetWinntProviderResults(strLocation & "/" & varItem, strQueryType)
                        varResults = Split(GetUserGroups(varItem), ",")
                        intColItems = 1
                        intColResults = 2
                    Case "ComputerGroups":
                        strColumns = GetComputerGroups(varItem)
                        'strColumns = GetWinntProviderResults(varItem, strQueryType)
                        varResults = Array(Split(strColumns(0), ","), Split(strColumns(1), ","))
                        intColItems = 1
                        intColResults = 2
                End Select
               ' loop and output results
                    If strQueryType = "ComputerGroups" Then
'                        Dim lngResultRows As Long
'                        If VarType(varResults(0)) > 8192 Then ' so its a multidimension array
'                            lngResultRows = UBound(varResults(0))
'                        Else
'                            lngResultRows = UBound(varResults)
'                        End If
'                        For lngRow = 0 To lngResultRows
                        For lngRow = LBound(varResults(0)) To UBound(varResults(0))
                            .Cells(lngOutputRow, intColItems).Value = varItem
                            For lngCol = LBound(varResults) To UBound(varResults)
                                .Cells(lngOutputRow, intColResults + lngCol).Value = varResults(lngCol)(lngRow)
                            Next lngCol
        '                    .Cells(lngOutputRow, intColResults + 1).Value = varResults(1)(lngRow)
                            lngOutputRow = lngOutputRow + 1
                        Next lngRow
                    Else
                   
                For lngRow = LBound(varResults) To UBound(varResults)
                    .Cells(lngOutputRow, intColItems).Value = varItem
                    .Cells(lngOutputRow, intColResults).Value = varResults(lngRow)
                    lngOutputRow = lngOutputRow + 1
                Next lngRow
                    End If
            End If
        Next varItem
    End With
End Function
 
Function GetGroupUsers(ByVal strGroupName As String) As String
    ' credit > http://www.excelforum.com/2280511-post12.html
    Application.StatusBar = "Performing " & "GetGroupUsers" & " on item named: " & strGroupName
 
    Dim objGroup, objDomain, objMember
    Dim strMemberList As String, strDomain As String
    On Error Resume Next
    Set objDomain = GetObject("LDAP://rootDse")
    strDomain = objDomain.Get("dnsHostName")
   
    Set objGroup = GetObject("WinNT://" & strDomain & "/" & strGroupName & ",group")
   
    ' without the object existing in the calling function it can't enumerate the instances
    ' so we serialise it as a string, making it possible for the caller to parse
    For Each objMember In objGroup.Members
        strMemberList = strMemberList & "," & objMember.Name
    Next objMember
    ' strip off the leading comma
    GetGroupUsers = Mid$(strMemberList, 2)
    Application.StatusBar = False
End Function
 
 
Function GetUserGroups(ByVal strUserName As String) As String
    Dim objUser, objDomain, objGroup
    Application.StatusBar = "Performing " & "GetUserGroups" & " on item named: " & strUserName
   
    Dim strGroupList As String, strDomain As String
    On Error Resume Next
    Set objDomain = GetObject("LDAP://rootDse")
    strDomain = objDomain.Get("dnsHostName")
   
    Set objUser = GetObject("WinNT://" & strDomain & "/" & strUserName & ",user")
   
    For Each objGroup In objUser.Groups
        strGroupList = strGroupList & "," & objGroup.Name
    Next objGroup
    ' strip off the leading comma
    GetUserGroups = Mid$(strGroupList, 2)
    Application.StatusBar = False
End Function
 
 
Function GetComputerGroups(ByVal strComputerName As String) As Variant
    Dim strResultList(1) As String
 
    Application.StatusBar = "Performing " & "GetComputerGroups" & " on item named: " & strComputerName
 
    ' This uses simple GetObject object management, like VBS style - brute, difficult to debug but effective
    Dim objList, ObjResult, ObjChild
 
'    Set objList = GetObject("WinNT://" & strComputerName & ",group")
    Set objList = GetObject("WinNT://" & strComputerName & "")
    ' credit > http://www.activexperts.com/activmonitor/windowsmanagement/adminscripts/usersgroups/localgroups/
    objList.Filter = Array("group")
    ' help > search WinNT in http://blogs.technet.com/b/heyscriptingguy/archive/2004/12/13/how-can-i-run-a-script-under-alternate-credentials.aspx
 
    ' without the object existing in the calling function it can't enumerate the instances
    ' so we serialise it as a string, making it possible for the caller to parse
    ' in VBA it's actually far more efficient to extend a string than it is to perpetually redim an array!
    For Each ObjResult In objList
        For Each ObjChild In ObjResult.Members
            strResultList(0) = strResultList(0) & "," & ObjResult.Name ' this is the specific property required
            strResultList(1) = strResultList(1) & "," & ObjChild.Name
        Next
    Next
    ' strip off the leading comma as we return the function's result
    strResultList(0) = Mid$(strResultList(0), 2)
    strResultList(1) = Mid$(strResultList(1), 2)
 
    Application.StatusBar = False
 
    GetComputerGroups = strResultList
 
End Function
 
 
 
 
'Function GetWinntProviderResults(ByVal strItemName As String, ByVal strQueryType As String) As Variant
'    Dim strResultList(1) As String
'    Application.StatusBar = "Performing " & strQueryType & " on item named: " & strItemName
'
'    ' This uses simple GetObject object management, like VBS style - brute, difficult to debug but effective
'    Dim objResponse, objEnumerate1, objInstance1, objEnumerate2, objInstance2
'    Dim intEnumerations, intEnum As Integer
'
'    Dim strFilter
'
'    Select Case strQueryType
'        Case "ListUsers":
''            varResults = Split(GetGroupUsers(varItem), ",")
''            intColItems = 2
''            intColResults = 1
'        Case "ListGroups":
'            strFilter = "user"
'            intEnumerations = 1
'        Case "ComputerGroups":
'            strFilter = "group"
'            intEnumerations = 2
''            varResults = Array(Split(strColumns(0), ","), Split(strColumns(1), ","))
''            intColItems = 1
''            intColResults = 2
'    End Select
'
''    'location (e.g. domain) is now passed as a prefix on ItemName
''    Set objDomain = GetObject("LDAP://rootDse")
''    strDomain = objDomain.Get("dnsHostName")
'
'    'Set objResponse = GetObject("WinNT://" & strItemName & "")
'    Set objResponse = GetObject("WinNT://" & strItemName & "," & strFilter)
'    ' credit > http://www.activexperts.com/activmonitor/windowsmanagement/adminscripts/usersgroups/localgroups/
'    'objResponse.Filter = Array(strFilter)
'    ' help > search WinNT in http://blogs.technet.com/b/heyscriptingguy/archive/2004/12/13/how-can-i-run-a-script-under-alternate-credentials.aspx
'
'
'    Select Case strQueryType
'        Case "ListUsers":
'        Case "ListGroups":
'            Set objEnumerate1 = objResponse.Groups
'        Case "ComputerGroups":
'            Set objEnumerate1 = objResponse
'    End Select
'    For Each objInstance1 In objEnumerate1
'
'        Select Case strQueryType
'            Case "ListUsers":
'            Case "ListGroups":
'                Set objEnumerate2 = objInstance1
'            Case "ComputerGroups":
'                Set objEnumerate2 = objInstance1.Members
'        End Select
''        For Each objInstance2 In objEnumerate2
'
'            ' without the object existing in the calling function it can't enumerate the instances
'            ' so we serialise it as a string, making it possible for the caller to parse
'            ' in VBA it's actually far more efficient to extend a string than it is to perpetually redim an array!
'
'            strResultList(0) = strResultList(0) & "," & objInstance1.Name ' this is the specific property required
'            If intEnumerations = 2 Then
'                strResultList(1) = strResultList(1) & "," & objInstance2.Name
'            End If
''        Next
'    Next
'    ' strip off the leading comma as we return the function's result
'    For intEnum = LBound(strResultList) To UBound(strResultList)
'        strResultList(intEnum) = Mid$(strResultList(intEnum), 2)
'    Next intEnum
'
'    Application.StatusBar = False
'
'    GetWinntProviderResults = strResultList
'
'End Function
