Attribute VB_Name = "mod_acc_DataMisc"
Option Compare Database
Option Explicit

'************************************************
'  GENERIC DATA ACCESS CODE
'
'  100414.AMG  added update & source to addDependentRecords and created CSqlFld function
'  100408.AMG  improve file chooser dialog
'  100331.AMG  dialog to choose a single file to open
'  100219.AMG  added sample code for multi-sheet import
'  100219.AMG  added generic XLS import
'  080418.AMG  added InsertIfNotExists
'  080417.AMG  included DropAndRecreate for tables and improved comments and structure
'  080327.AMG  Added ExecuteWithDAO and DBAddDependentRecords
'              also renamed from modDataGeneric.bas and added extras from modData_misc.bas
'  060907.AMG  renamed strSqlVariant to CSql added double quote option and removed strSqlDate
'  060906.AMG  adopted modData_misc from Hart
'
'************************************************

Public Function ExecuteAgainstDB(strSQL As String) As Long
    With Application.CurrentProject.Connection
        .Execute CommandText:=strSQL, _
            RecordsAffected:=ExecuteAgainstDB
    End With
End Function

Public Function ExecuteWithDAO(strSQL As String) As Long
    CurrentDb.Execute strSQL
End Function


Public Function RecordsetFromDB(strSQL As String) As ADODB.Recordset
    Set RecordsetFromDB = New ADODB.Recordset

    With RecordsetFromDB

        .Open Source:=strSQL, _
            ActiveConnection:=CurrentProject.Connection, _
            CursorType:=adOpenForwardOnly, _
            LockType:=adLockReadOnly
    End With
End Function

Public Function ExecuteAgainstDBReturnID(strSQL As String) As Long
    ExecuteAgainstDBReturnID = -1

    ExecuteAgainstDB (strSQL)
    With RecordsetFromDB("SELECT @@IDENTITY;")
        If Not .EOF Then
            ExecuteAgainstDBReturnID = .Fields(0).Value
        End If
    End With
End Function


Public Function DBreturnLong(strSQL As String) As Long
    DBreturnLong = 0
    On Error Resume Next
    
    With RecordsetFromDB(strSQL)
        If Not .EOF Then
            DBreturnLong = CLng(.Fields(0).Value)
        End If
    End With
End Function


Public Function DBreturnString(strSQL As String) As String
    DBreturnString = ""
    On Error Resume Next

    With RecordsetFromDB(strSQL)
        If Not .EOF Then
            DBreturnString = CStr(.Fields(0).Value)
        End If
    End With
End Function

' ----------------------------------------
' Preparing values to pass via SQL strings
' ----------------------------------------
'
' These functions return strings containing
' valid SQL expression fragments
' where VB variables have been correctly converted
' ready for interpretation by a SQL database engine

Public Function CSql(var As Variant) As String
' format any type as a string in the format Jet SQL expects
' numerics - in US format (not localised)
' dates - in US date format enclosed by hashes
' strings containing single quotes or apostrophes - enclosed in double quotes
' all other strings - simply enclosed in single quotes
    If IsNumeric(var) Then
        CSql = Str(var)
    ElseIf IsDate(var) Then
        CSql = Format(var, "\#MM/DD/YYYY\#")
    ElseIf InStr(var, "'") > 0 Then
        CSql = """" & var & """"
    Else
        CSql = "'" & var & "'"
    End If
End Function
Public Function CSqlFld(strFieldName As String) As String
    CSqlFld = "[" & strFieldName & "]"
End Function

Public Function strSqlPartialMatch(strFieldName As String, varValue As Variant) As String
    strSqlPartialMatch = CSqlFld(strFieldName) & " LIKE '*" & varValue & "*'"
End Function
Public Function strSqlExactMatch(strFieldName As String, varValue As Variant) As String
    strSqlExactMatch = CSqlFld(strFieldName) & " = " & CSql(varValue)
End Function

' -------------------------
' Preparing returned values
' -------------------------
'
' These functions get around "invalid use of Null" errors
Public Function strIgnoreNulls(varString As Variant) As String
    If IsNull(varString) Then
        strIgnoreNulls = ""
    Else
        strIgnoreNulls = CStr(varString)
    End If
End Function
Public Function lngIgnoreNulls(varString As Variant) As Long
    If IsNull(varString) Then
        lngIgnoreNulls = 0
    Else
        lngIgnoreNulls = CLng(varString)
    End If
End Function


Function strChooseFileToOpen(Optional strTitle As String) As String
' The Excel version would use...
'    strFileName = Application.GetOpenFilename("Excel Worksbooks (*.xls), *.xls", , "Please select the GDC Move workbook")
' There is a long winded code to do somehting similar at
'   sample > http://www.mvps.org/access/api/api0001.htm
' but the simple way is...
' Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
' making sure References includes Microsoft Office xx.0 Object Library
' credit > http://www.ozgrid.com/forum/showthread.php?t=28754
' credit > http://support.microsoft.com/kb/288543

    Dim dlgOpen As FileDialog
    Dim vrtSelectedItem As Variant ' need variant to extract choices from list

    Set dlgOpen = Application.FileDialog(msoFileDialogFilePicker)
    With dlgOpen
        If Not IsMissing(strTitle) Then
            .Title = strTitle
        End If
        .AllowMultiSelect = False
        If .Show = -1 Then ' if the user DIDN'T cancel
            For Each vrtSelectedItem In .SelectedItems
                strChooseFileToOpen = vrtSelectedItem
            Next vrtSelectedItem
        Else
            strChooseFileToOpen = ""
        End If
    End With
    Set dlgOpen = Nothing
End Function

Sub DbImportXls(strTableName As String, strExcelFilename As String, Optional strTableDef As String, Optional strRange As String)

    On Error Resume Next
    ExecuteAgainstDB "DROP TABLE " & strTableName
    On Error GoTo 0
    If Not IsMissing(strTableDef) Then
        If strTableDef <> "" Then
            ExecuteAgainstDB "CREATE TABLE " & strTableName & " ( " & strTableDef & " );"
        End If
    End If

    If IsMissing(strRange) Then
        DoCmd.TransferSpreadsheet _
            TransferType:=acImport, _
            SpreadsheetType:=acSpreadsheetTypeExcel9, _
            TableName:=strTableName, _
            FileName:=strExcelFilename, _
            HasFieldNames:=True
    '        Range:="", _
    '        UseOA:=False
    Else
        DoCmd.TransferSpreadsheet _
            TransferType:=acImport, _
            SpreadsheetType:=acSpreadsheetTypeExcel9, _
            TableName:=strTableName, _
            FileName:=strExcelFilename, _
            HasFieldNames:=True, _
            Range:=strRange
    '        UseOA:=False
    End If

End Sub

Sub sample_multi_sheet_import()
' credit > http://blogs.technet.com/heyscriptingguy/archive/2008/01/21/how-can-i-import-multiple-worksheets-into-an-access-database.aspx
' NB: this is VB script
'    Const acImport = 0
'    Const acSpreadsheetTypeExcel9 = 8
'
'    Set objAccess = CreateObject("Access.Application")
'    objAccess.OpenCurrentDatabase "C:\Scripts\Personnel.mdb"
'
'    Set objExcel = CreateObject("Excel.Application")
'    objExcel.Visible = True
'
'    strFileName = "C:\Scripts\ImportData.xls"
'
'    Set objWorkbook = objExcel.Workbooks.Open(strFileName)
'    Set colWorksheets = objWorkbook.Worksheets
'
'    For Each objWorksheet In colWorksheets
'        Set objRange = objWorksheet.UsedRange
'        strWorksheetName = objWorksheet.Name & "!" & objRange.Address(False, False)
'        objAccess.DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, _
'            "Employees", strFileName, True, strWorksheetName
'    Next
End Sub

' ####### Relationship Management and Database Definition ########
' ################################################################
'
' This section tends to use DAO as I have found it simpler so far
' to use Database Definition Language in Access through DAO than ADO


' When you want to Import Records into a relational database
' call this function before doing your INSERT INTO x SELECT y FROM z;
' So that the relevant Lookup values are there and do not
' cause integrity issues or dropped inserts

Public Function DBAddDependentRecords _
    (strImportTable As String _
    , strImportField As String _
    , strLookupTable As String _
    , strLookupField As String _
    , strLookupId As String _
    , Optional strUpDateField As String _
    , Optional strSourceField As String _
    , Optional strSourceString As String _
)

    Dim strSQL As String
    strSQL = _
    "INSERT INTO " & strLookupTable _
    & " SELECT " & strLookupField

    If Not IsMissing(strUpDateField) Then
        strSQL = strSQL _
        & " , Now() AS " & CSqlFld(strUpDateField) & " "
    End If

    If Not IsMissing(strSourceString) Then
        strSQL = strSQL _
        & " , " & CSql(strSourceString) & " AS " & CSqlFld(strSourceField) & " "
    End If

    strSQL = strSQL _
    & " FROM ( " _
        & " SELECT DISTINCT " & strImportTable & "." & strImportField & " AS " & strLookupField _
        & " FROM " & strImportTable & " LEFT JOIN " & strLookupTable _
        & " ON " & strImportTable & "." & strImportField _
        & " = " & strLookupTable & "." & strLookupField _
        & " GROUP BY " & strImportTable & "." & strImportField _
        & " HAVING (Count(" & strLookupTable & "." & strLookupId & ")=0) " _
        & " AND NOT (" & strImportTable & "." & strImportField & " Is Null) " _
    & " ) ;"

    ' haven't worked out yet why this query fails with ADO - seems ok with DAO
    ExecuteWithDAO strSQL
End Function


' Old column manipulation stuff - not fully tested
'        ' tidy from import XLS - rename erroneous column name and remove cols after 30
'        ' credit > http://forums.devx.com/showthread.php?t=50878
'        With CurrentDb.TableDefs("GDC_Not_moving_List")
'            .Fields("Numbr of CPU").Name = "Number of CPU"
'            'While .Fields.Count > 30
'            '    .Fields.Delete (.Fields(30).Name)
'            'Wend
'        End With



' Use this to insert a record if it doesn't already exist
' Very useful with lookup tables
'
' NB: FAILS IF THE TABLE IS EMPTY - THERE MUST BE AT LEAST ONE ROW
'
' Thanks to Marco De Luca (delucam@xebec.ca)
' for saving me from having to work the logic out from scratch
'
' The plain SQL code is...
'
' INSERT INTO LookupTable
'     (LookupField, DetailField)
' SELECT DISTINCT
'     'Lookup Value' as LookupField,
'     'Detail Value' as DetailField
' FROM LookupTable
' WHERE 'Lookup Value' NOT In
'     (SELECT LookupField from LookupTable);

Public Function InsertIfNotExists _
    (strLookupTable As String _
    , strLookupField As String _
    , strLookupValue As String _
    , Optional strDetailField As String _
    , Optional strDetailValue As String _
) As Long
        
    Dim strSQL As String
    strSQL = _
       "INSERT INTO " & strLookupTable _
    & " ( " & strLookupField
    
    If Not IsMissing(strDetailValue) Then ' only add the field if the value is there too
        strSQL = strSQL _
    & " , " & strDetailField
    End If
    
    strSQL = strSQL _
    & " ) " _
    & " SELECT DISTINCT " _
    & CSql(strLookupValue) & " AS " & strLookupField

    If Not IsMissing(strDetailValue) Then
        strSQL = strSQL _
    & " , " & CSql(strDetailValue) & " AS " & strDetailField
    End If
    
    strSQL = strSQL _
    & " FROM " & strLookupTable _
    & " WHERE " & CSql(strLookupValue) & " NOT IN " _
        & " (SELECT " & strLookupField _
        & "  FROM " & strLookupTable _
    & " ) ;"

    InsertIfNotExists = ExecuteAgainstDBReturnID(strSQL)
    
End Function


Public Function DropAndRecreateTable(strTableName As String, strSQLTableDef As String)
    On Error GoTo JustCreate
    If CurrentDb.TableDefs(strTableName).Fields.Count <> 0 Then
        ExecuteWithDAO "DROP TABLE " & strTableName & ";"
    End If

'    Dim td As TableDef
'    For Each td In CurrentDb.TableDefs
'        If td.Name = strTableName Then
'            ExecuteWithDAO "DROP TABLE " & strTableName & ";"
'        End If
'    Next

JustCreate:
    On Error GoTo 0

    ExecuteWithDAO "CREATE TABLE " & strTableName & " ( " & strSQLTableDef & ");"
End Function



Public Sub CreateQueryFromString(strQryName As String, strSQL As String)
On Error Resume Next
    If CurrentDb.QueryDefs(strQryName).SQL <> strSQL Then
        CurrentDb.QueryDefs(strQryName).SQL = strSQL
    End If

    If Err.Number = 3265 Then ' Error: Object not found in this collection
        Err.Clear
        CurrentDb.CreateQueryDef strQryName, strSQL
        If Err.Number <> 0 Then
            MsgBox "Could not create query" & vbCrLf & vbCrLf _
                & strQryName & vbCrLf & vbCrLf _
                & "Error " & Err.Number & " - " & Err.Description, _
                vbCritical, _
                "Error creating Query!"
        End If
    ElseIf Err.Number = 3012 Then ' Error: Object <name> already exists
        On Error GoTo 0
        CurrentDb.QueryDefs(strQryName).SQL = strSQL
    ElseIf Err.Number <> 0 Then
        MsgBox "Could not recreate query" & vbCrLf & vbCrLf _
            & strQryName & vbCrLf & vbCrLf _
            & "Error " & Err.Number & " - " & Err.Description, _
            vbCritical, _
            "Error recreating Query!"
    End If
On Error GoTo 0
End Sub



'
' ####### Deprecated ##########################
' #############################################
'
' The following code may not be very generic, so may be of little value...
'
'

' This is used to modify the tables we link to from the interface
' It uses DAO to find the linked location of the named table
' It then accesses the linked database directly via ADO to
' make the modification
'
Public Sub UpgradeDB(strSQL As String)
On Error GoTo ErrorHandler
    Dim strDAOConnect As String
    Dim strADODBConnectionString As String
    Dim cnn As ADODB.Connection

'   For now we can use any table in the database,
'   as they are all in the smae location, but
'   if ever the back end was split, the calling
'   function would have to pass the table name
    Const strTableName As String = "Audits"
    strDAOConnect = CurrentDb.TableDefs(strTableName).Connect

    If Left(strDAOConnect, 10) <> ";DATABASE=" Then
        MsgBox "Cannot correctly identify data source location" & vbCrLf & vbCrLf _
                & "DAO.TableDef.Connect = """ & strDAOConnect & """", _
                vbCritical + vbOKOnly, _
                "Database upgrade Error"
        Exit Sub
    End If
    
    strADODBConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" _
        & "Data Source=" & Mid(strDAOConnect, 11) & ";" _

    Set cnn = New ADODB.Connection
    cnn.Open strADODBConnectionString

    cnn.Execute strSQL

    cnn.Close
    Set cnn = Nothing

ErrorHandler:
    Select Case Err.Number
        Case 0: ' no action required
        Case Else
            MsgBox "We had not made contingencies for this error..." & vbCrLf & vbCrLf _
                    & "Number: " & Err.Number & vbCrLf _
                    & "Descxription: " & Err.Description & vbCrLf _
                    & "Source: " & Err.Source & vbCrLf & vbCrLf _
                    & "Procedure: ""UpgradeDB""", _
                    vbCritical + vbOKOnly, _
                    "Unanticipated Error"
    End Select
End Sub

