Attribute VB_Name = "mod_acc_Documenter"
Option Compare Database

'
' mod_acc_Documenter
'
'
' 100505.AMG  added TableInfo, renamed from mod_acc_References
' 0603xx.AMG  created as mod_vba_references within HART
' 080229.AMG  renamed from mod_vba_references
' 080417.AMG  added quick reference documentor
'

' References
' ==========
'
' This module may require the following references (paths and GUIDs might vary)
'
' DAO (C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll) 5.0 - {00025E01-0000-0000-C000-000000000046}


Sub DocumentReferences()
    Dim ref As Reference
    
    Debug.Print "' References"
    Debug.Print "' =========="
    Debug.Print "'"
    Debug.Print "' This module may require the following references (paths and GUIDs might vary)"
    Debug.Print "'"
    For Each ref In Application.References
'        If Not ref.BuiltIn Then
            Debug.Print "' " & ref.Name & " (" & ref.FullPath & ") " & ref.Major & "." & ref.Minor & " - " & ref.Guid
'        End If
    Next
End Sub


''' ### Copied Code starts ###
' source > http://www.everythingaccess.com/tutorials.asp?ID=Dump-table-details-in-VBA-(DAO)
' credit > Allen Browne, allen@allenbrowne.com. Updated June 2006
'TableInfo() function
'This function displays in the Immediate Window (Ctrl+G) the structure of any table in the current database.
'For Access 2000 or 2002, make sure you have a DAO reference.
'The Description property does not exist for fields that have no description, so a separate function handles that error.

Function TableInfo(strTableName As String)
On Error GoTo TableInfoErr
   ' Purpose:   Display the field names, types, sizes and descriptions for a table.
   ' Argument:  Name of a table in the current database.
   Dim db As DAO.Database
   Dim tdf As DAO.TableDef
   Dim fld As DAO.Field
   
   Set db = CurrentDb()
   Set tdf = db.TableDefs(strTableName)
   Debug.Print "FIELD NAME", "FIELD TYPE", "SIZE", "DESCRIPTION"
   Debug.Print "==========", "==========", "====", "==========="

   For Each fld In tdf.Fields
      Debug.Print fld.Name,
      Debug.Print FieldTypeName(fld),
      Debug.Print fld.Size,
      Debug.Print GetDescrip(fld)
   Next
   Debug.Print "==========", "==========", "====", "==========="

TableInfoExit:
   Set db = Nothing
   Exit Function

TableInfoErr:
   Select Case Err
   Case 3265&  'Table name invalid
      MsgBox strTableName & " table doesn't exist"
   Case Else
      Debug.Print "TableInfo() Error " & Err & ": " & Error
   End Select
   Resume TableInfoExit
End Function


Function GetDescrip(obj As Object) As String
    On Error Resume Next
    GetDescrip = obj.Properties("Description")
End Function


Function FieldTypeName(fld As DAO.Field) As String
    'Purpose: Converts the numeric results of DAO Field.Type to text.
    Dim strReturn As String    'Name to return

    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn
End Function
''' ### Copied Code ends ###



' The rest of this code is depreciated #####################################################
'###########################################################################################

Private Sub SetReferences()
'
' You can find the GUIDs & Versions in the registry
' of destination clients under HKCR/TypeLib
'
    If Not RefreshReference("ADODB", "{2A75196C-D9EB-4129-B803-931327F72D5C}", "2.8") Then
        RefreshReference "ADODB", "{00000201-0000-0010-8000-00AA006D2EA4}", "2.1"
    End If
    RefreshReference "DAO", "{00025E01-0000-0000-C000-000000000046}", "5.0"
    RefreshReference "Office", "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", "2.4, 2.3, 2.2, 2.1"
End Sub

Private Function RefreshReference(strName As String, strGUID As String, strVersions As String) As Boolean
'
' Samples for strVersions include "1.0" and "1.1, 1.2"
' Versions separated by commas, major and minor split by a period
' If one succedes then we stop trying, so list them in order of preference
'
    Const ERR_OBJECTLIB_NOT_REG As Long = -2147319779
    Const ERR_OBJECTLIB_CONFLICT As Long = 32813

    Dim ref As Reference
    Dim i As Integer
    Dim iMaxUBound As Integer
    Dim strEachVer() As String
    Dim strMajMin() As String

    ' default return value
    RefreshReference = False

    For Each ref In Application.References
        If ref.Name = strName Then
            References.Remove ref
        End If
    Next
    ' an alternative might have been to test the property
    '   ref.IsBroken
    ' and only procede if it were

    strEachVer = Split(strVersions, ",")
    For i = 0 To UBound(strEachVer)
        strMajMin = Split(strEachVer(i), ".")
        If UBound(strMajMin) = 1 Then
            On Error Resume Next
            Set ref = Application.References.AddFromGuid(strGUID, CLng(strMajMin(0)), CLng(strMajMin(1)))
            Debug.Print strName, CLng(strMajMin(0)), CLng(strMajMin(1)), Err.Number, Err.Description
            Select Case Err.Number
                Case ERR_OBJECTLIB_NOT_REG:       ' if badly registered try to remove
                        References.Remove ref
                Case ERR_OBJECTLIB_CONFLICT:      ' this should only happen if we did't exit on success
                                                    ' no action for now
                Case 0:
                        RefreshReference = True
                        Exit For
            End Select
        End If
    Next
End Function



