Attribute VB_Name = "mod_exc_LinksFiles"
Option Explicit

' error handling tag
Const cStrModuleName As String = "mod_exc_LinksFiles"

' generic functions for manipulating filesystem objects
' and web and sharepoint sites and URLs
'
'  150309.AMG  also get URL sub-address
'  150219.AMG  added GetURL for hyperlinks
'

' This module may require the following references (paths and GUIDs might vary)
'      Tools / References / Microsoft Scripting Runtime
' Scripting (C:\WINDOWS\system32\scrrun.dll) 1.0 - {420B2830-E718-11CF-893D-00A0C9054228}
'    (or C:\Windows\SysWOW64\scrrun.dll)
' MSXML2 (C:\WINDOWS\system32\msxml6.dll) 6.0 - {F5078F18-C551-11D3-89B9-0000F81FE221}


Const cStrExcFileFilter As String = "Excel Workbooks, *.xls; *.xlsx"
'        Case "xls": strFilter = "Excel Workbooks (*.xls), *.xls"
'        Case "txt": strFilter = "Text Files (*.txt), *.txt"
'        Case Else: strFilter = "All Files (*.*), *.*"

' *********** HYPERLINKS *********************************************

Function GetURL(rngCell As Range) As String
    If rngCell.Hyperlinks.Count > 0 Then
        GetURL = Replace _
            (rngCell.Hyperlinks(1).Address, "mailto:", "")

        If rngCell.Hyperlinks(1).SubAddress <> "" Then
            ' credit http://excel.tips.net/T003281_Extracting_URLs_from_Hyperlinks.html
            GetURL = GetURL & "#" & rngCell.Hyperlinks(1).SubAddress
        End If
    End If

End Function




' *********** FILE AND PATH NAMES *********************************************

' Make this use generic arrFilteredPathnamesInUserTree in mod_off_FilesFoldersSitesLinks

' re cast as array of full file paths
Function EnumerateExcelFiles()
' This module is used to enumerate all XLS files in a
' folder, chosen by the user
'  071030.AMG  created

Dim SourceFilename As String
Dim SourceFolderName As String
Dim wbk As Workbook
' Ask user to identify a file in source folder
'

SourceFilename = CStr(Application.GetOpenFilename( _
                    FileFilter:=cStrExcFileFilter, _
                    Title:="Choose one of the files in the source folder", _
                    ButtonText:="Select"))

SourceFolderName = GetFolderFromFileName(SourceFilename)

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
Dim fso As Scripting.FileSystemObject
Dim SourceFolder As Scripting.folder
Dim SourceFile As Scripting.file
Set fso = New Scripting.FileSystemObject
Set SourceFolder = fso.GetFolder(SourceFolderName)
'Application.ScreenUpdating = False
For Each SourceFile In SourceFolder.Files
    If LCase(Right(SourceFile.Name, 5)) = ".xlsx" _
      Or LCase(Right(SourceFile.Name, 4)) = ".xls" Then

' Open each workbook and put the name in the status bar
        Application.StatusBar = "reading from [" & SourceFile.Name & " ]..."
        Set wbk = Workbooks.Open( _
            FileName:=SourceFolderName & Application.PathSeparator & SourceFile.Name, _
            UpdateLinks:=0, _
            ReadOnly:=True, _
            IgnoreReadOnlyRecommended:=True _
            )

' ******************
' Do your stuff here
' ******************

        wbk.Close SaveChanges:=False
    End If
Next SourceFile


End Function



