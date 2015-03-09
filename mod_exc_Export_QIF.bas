Attribute VB_Name = "mod_exc_Export_QIF"
Option Explicit

' 121230.AMG

' some credit > http://gnucash.1415818.n4.nabble.com/Importing-xls-or-similar-files-tp1425825.html
' however using very simple QIF format as per FDxx downloads

' column numbers for BPEC output
Enum ColNo
    Date = 1
    Debit = 3
    Credit = 4
    Desc = 7
End Enum


Sub ExportEntriesAsQIF()
    
    ' create QIF file based on XLS name
    Dim strPath As String
    strPath = Application.ActiveWorkbook.FullName & ".qif"
    Dim fso, qif As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set qif = fso.CreateTextFile(strPath, True)

    ' very simple header
    qif.writeline ("!Type:Bank")

    ' for each row, except the header, and any rows with a blank date ...
    Dim bHeader As Boolean
    bHeader = True
    Dim rRow As Range
    For Each rRow In ActiveSheet.UsedRange.Rows
        If bHeader Then
            bHeader = False
        Else
            If rRow.Cells(1, ColNo.Date).Value <> "" Then

                ' write the record as a simple QIF entry
                qif.writeline ("D") & rRow.Cells(1, ColNo.Date).Value
                qif.writeline ("P") & rRow.Cells(1, ColNo.Desc).Value
                qif.writeline ("T") & rRow.Cells(1, ColNo.Debit).Value & rRow.Cells(1, ColNo.Credit).Value
                qif.writeline ("^")
            End If
        End If
    Next

    MsgBox ("QIF exported to " & strPath)
End Sub


