Attribute VB_Name = "mod_acc_SchemaTableIntoExcel"
Option Compare Database
Option Explicit

' mod_acc_SchemaTableIntoExcel

' 130927.AMG  cribbed from forum and allen browne site


' credit > http://www.access-programmers.co.uk/forums/showpost.php?p=1036536&postcount=10

Sub ListTablesAndFields()
     'Macro Purpose:  Write all table and field names to and Excel file
     'Source:  vbaexpress.com/kb/getarticle.php?kb_id=707
     'Updates by Derek - Added column headers, modified base setting for loop to include all fields,
     '                   added type, size, and description properties to export
     
    Dim lTbl As Long
    Dim lFld As Long
    Dim dBase As Database
    Dim xlApp As Object
    Dim wbExcel As Object
    Dim lRow As Long

     'Set current database to a variable adn create a new Excel instance
    Set dBase = CurrentDb
    Set xlApp = CreateObject("Excel.Application")
    Set wbExcel = xlApp.workbooks.Add
     
     'Set on error in case there are no tables
    On Error Resume Next
     
    'DJK 2011/01/27 - Added in column headers below
    lRow = 1
    With wbExcel.sheets(1)
        .Range("A" & lRow) = "Table Name"
        .Range("B" & lRow) = "Field Name"
        .Range("C" & lRow) = "Type"
        .Range("D" & lRow) = "Size"
        .Range("E" & lRow) = "Description"
    End With
    
     'Loop through all tables
    For lTbl = 0 To dBase.TableDefs.Count
         'If the table name is a temporary or system table then ignore it
        If Left(dBase.TableDefs(lTbl).Name, 1) = "~" Or _
        Left(dBase.TableDefs(lTbl).Name, 4) = "MSYS" Then
             '~ indicates a temporary table
             'MSYS indicates a system level table
        Else
             'Otherwise, loop through each table, writing the table and field names
             'to the Excel file
            For lFld = 0 To dBase.TableDefs(lTbl).Fields.Count - 1  'DJK 2011/01/27 - Changed initial base from 1 to 0, and added type, size, and description
                lRow = lRow + 1
                With wbExcel.sheets(1)
                    .Range("A" & lRow) = dBase.TableDefs(lTbl).Name
                    .Range("B" & lRow) = dBase.TableDefs(lTbl).Fields(lFld).Name
                    .Range("C" & lRow) = FieldTypeName(dBase.TableDefs(lTbl).Fields(lFld))
                    .Range("D" & lRow) = dBase.TableDefs(lTbl).Fields(lFld).Size
                    .Range("E" & lRow) = GetDescrip(dBase.TableDefs(lTbl).Fields(lFld))
                End With
            Next lFld
        End If
    Next lTbl
     'Resume error breaks
    On Error GoTo 0
     
     'Set Excel to visible and release it from memory
    xlApp.Visible = True
    Set xlApp = Nothing
    Set wbExcel = Nothing
     
     'Release database object from memory
    Set dBase = Nothing
     
End Sub

Function FieldTypeName(fld As DAO.Field) As String
    'Purpose: Converts the numeric results of DAO Field.Type to text.
' credit > http://allenbrowne.com/func-06.html

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

        'Constants for complex types don't work prior to Access 2007 and later.
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

Function GetDescrip(obj As Object) As String
    On Error Resume Next
    GetDescrip = obj.Properties("Description")
End Function



