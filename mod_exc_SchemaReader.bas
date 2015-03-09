Attribute VB_Name = "modSchemaReader"
Option Explicit

' YOU MUST ADD A REFERENCE TO *** SCRIPTING *** OBJECT TO USE THIS
' Tools / References / Microsoft Scripting Runtime
' e.g. C:\Windows\SysWOW64\scrrun.dll

Const cStrFileFilter As String = "Excel Workbooks, *.xls; *.xlsx"
Const cbDoubleRow As Boolean = True

Sub EnumerateExcelSchemas()

' This module is used to enumerate all XLS files in a
' folder, chosen by the user, and examine the schemas
' used in every sheet
' Basically it populates the sheet in THIS workbook with
' the spreadsheet name, the worksheet name and all
' column headings from row A
'
'  141105.AMG  do xls & xlsx, transpose, rowcount & doublespace
'  071030.AMG  created
'

Dim OutSheet As Worksheet
Dim OutCount, OutRow As Integer

Dim SourceFilename As String
Dim SourceFolderName As String
Dim SourceBook As Workbook
Dim sht As Worksheet
Dim col As Range

Dim FieldCount As Integer


' Ask user to identify a file in source folder
'

SourceFilename = CStr(Application.GetOpenFilename( _
                    FileFilter:=cStrFileFilter, _
                    Title:="Choose one of the files in the source folder", _
                    ButtonText:="Select"))

SourceFolderName = GetFolderFromFileName(SourceFilename)

' Prepare output sheet
'

Set OutSheet = ActiveWorkbook.Worksheets("Schemas")
OutSheet.UsedRange.Clear
OutCount = 0
OutRow = OutCount + 1
OutSheet.Cells(OutRow, 1).Value = "ID"
OutSheet.Cells(OutRow, 2).Value = "Workbook"
OutSheet.Cells(OutRow, 3).Value = "Sheet"
OutSheet.Cells(OutRow, 4).Value = "Rows"
OutSheet.Cells(OutRow, 5).Value = "Fields"
OutSheet.Rows(OutRow).Font.Bold = True
OutSheet.Activate
OutSheet.Cells(2, 1).Select
ActiveWindow.FreezePanes = True

' enumerate all XLS files in the folder

' I wanted to use FileSearch VBA object,
' but it looks like SearchScopes were getting in the way
'
'With Application.FileSearch
'    .NewSearch
'    .LookIn = SourceFolder
'    .SearchSubFolders = False
'    '.FileType = msoFileTypeExcelWorkbooks
'    .FileName = "*.xls"
'    If .Execute > 0 Then
'        MsgBox "There were " & .FoundFiles.Count & _
'            " file(s) found."
'        Dim FileCount As Integer
'        For FileCount = 1 To .FoundFiles.Count
'            MsgBox .FoundFiles(FileCount)
'
'
'        Next FileCount
'    End If
'End With

' so I went back to good ol' FileSystemObject from Shell Scripting
'
Dim FSO As Scripting.FileSystemObject
Dim SourceFolder As Scripting.Folder
Dim SourceFile As Scripting.File
Set FSO = New Scripting.FileSystemObject
Set SourceFolder = FSO.GetFolder(SourceFolderName)
'Application.ScreenUpdating = False
For Each SourceFile In SourceFolder.Files
    If LCase(Right(SourceFile.Name, 5)) = ".xlsx" _
      Or LCase(Right(SourceFile.Name, 4)) = ".xls" Then

' Open each workbook and put the name in the status bar
        Application.StatusBar = "reading from [" & SourceFile.Name & " ]..."
        Workbooks.Open _
            FileName:=SourceFolderName & Application.PathSeparator & SourceFile.Name, _
            UpdateLinks:=0, _
            ReadOnly:=True, _
            IgnoreReadOnlyRecommended:=True

        Set SourceBook = ActiveWorkbook
        OutSheet.Activate


' get the details from each source worksheet

        For Each sht In SourceBook.Worksheets
            OutCount = OutCount + 1
            OutRow = IIf(cbDoubleRow, OutCount * 2, OutCount + 1)
            OutSheet.Cells(OutRow, 1).Value = OutCount
            OutSheet.Cells(OutRow, 2).Value = SourceFile.Name
            OutSheet.Cells(OutRow, 3).Value = sht.Name
            OutSheet.Cells(OutRow, 4).Value = sht.UsedRange.Rows.Count
            FieldCount = 0

            For Each col In sht.UsedRange.Columns
                FieldCount = FieldCount + 1
                OutSheet.Cells(OutRow, FieldCount + 4).Value = col.Cells(1).Value
            Next col
'            OutSheet.Columns(OutRow).EntireColumn.AutoFit
        Next sht

' close and loop
'
        SourceBook.Close SaveChanges:=False
    End If
Next SourceFile

'Application.ScreenUpdating = True
Application.StatusBar = False
MsgBox "Finished reading Excel worksheet schemas from source folder"

End Sub




Function GetFolderFromFileName(FileName As String) As String
' Folder Name extraction routine loosely based on code from ExcelTip.com (Function FileOrFolderName)
    Dim Position As Integer
    Position = 0
    While InStr(Position + 1, FileName, Application.PathSeparator) > 0
        Position = InStr(Position + 1, FileName, Application.PathSeparator)
    Wend
    If Position = 0 Then
        GetFolderFromFileName = CurDir
    Else
        GetFolderFromFileName = Left(FileName, Position - 1)
    End If
End Function


