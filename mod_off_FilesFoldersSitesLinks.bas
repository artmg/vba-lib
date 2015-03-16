Attribute VB_Name = "mod_off_FilesFoldersSitesLinks"
Option Explicit

' error handling tag
Const cStrModuleName As String = "mod_off_FilesFoldersSitesLinks"

' generic functions for manipulating filesystem objects
' and web and sharepoint sites and URLs
'
'  150316.AMG  added recursion into subfolders
'  150304.AMG  renamed from mod_exc_FilesFoldersSitesLinks as actually generic
'  150219.AMG  added GetURL for hyperlinks
'  150219.AMG  cribbed from other VBA modules - NB: not ALL functions have been tested since cribbing!
'

' References
' ==========
'
' This module uses the following references (paths and GUIDs may vary)
'
' Scripting = Microsoft Scripting Runtime (C:\Windows\SysWOW64\scrrun.dll) {420B2830-E718-11CF-893D-00A0C9054228}
' MSXML2 = Microsoft XML, v6.0 (C:\WINDOWS\System32\msxml6.dll) {F5078F18-C551-11D3-89B9-0000F81FE221}

' kludge for apps without Application.PathSeparator
Const cStrPathSeparator = "\"

Const cStrExcFileFilter As String = "Excel Workbooks, *.xls; *.xlsx"
'        Case "xls": strFilter = "Excel Workbooks (*.xls), *.xls"
'        Case "txt": strFilter = "Text Files (*.txt), *.txt"
'        Case Else: strFilter = "All Files (*.*), *.*"

' was mod_Acc_ExportToSharePoint
' credit > http://www.mrexcel.com/forum/excel-questions/332415-visual-basic-applications-code-excel-copy-source-file-sharepoint-another-destination.html
Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
                                           "URLDownloadToFileA" ( _
                                           ByVal pCaller As Long, ByVal szURL As String, _
                                           ByVal szFileName As String, _
                                           ByVal dwReserved As Long, _
                                           ByVal lpfnCB As Long) As Long


' *********** HYPERLINKS *********************************************





' *********** FILE AND PATH NAMES *********************************************


' was mod_acc_ImportBank.bas!JustFileName 060131.AMG
' Look for the last backslash and return just the characters following it
Public Function JustFileName(FullPath As Variant)
    Dim LastBackslash As Long
    LastBackslash = InStrRev(FullPath, cStrPathSeparator)
    If LastBackslash > 1 Then
        JustFileName = Mid(FullPath, LastBackslash + 1)
    Else
        JustFileName = FullPath
    End If
End Function


' was mod_exc_SchemaReader.bas GetFolderFromFileName 071030.AMG
Public Function GetFolderFromFileName(FileName As String) As String
' Folder Name extraction routine loosely based on code from ExcelTip.com (Function FileOrFolderName)
    Dim Position As Integer
    Position = 0
    While InStr(Position + 1, FileName, cStrPathSeparator) > 0
        Position = InStr(Position + 1, FileName, cStrPathSeparator)
    Wend
    If Position = 0 Then
        GetFolderFromFileName = CurDir
    Else
        GetFolderFromFileName = Left(FileName, Position - 1)
    End If
End Function


' was mod_exc_ParseAuditFiles 080326.AMG
' now uses InStrRev to get the last "." 150219.AMG
Function FileNameWithoutExtension(strFileName As String) As String
    Dim str As String
    Dim iPosn As Integer
    
    iPosn = InStrRev(strFileName, ".")
    If iPosn > 1 Then
        str = Left(strFileName, iPosn - 1)
    Else
        str = strFileName
    End If
    
    FileNameWithoutExtension = str
End Function



' *********** FOLDER FUNCTIONS *********************************************

' was mod_exc_ParseAuditFiles GetUserToPickFolder 080326.AMG
Function strFolderChosenByUser(strTitle As String) As String
' Ask user to identify a file to choose that folder


' credit - http://www.office-forums.com/threads/filedialog-visio-2003-dilemma.1621339/#post-5072038
    Dim xlApp As New Excel.Application

    ' help - https://msdn.microsoft.com/en-us/library/office/ff863983.aspx
    Dim dlg As FileDialog
'    Dim dlg As Office.FileDialog
    Set dlg = xlApp.FileDialog(msoFileDialogFolderPicker)

    dlg.Title = strTitle
'    dlg.ButtonName = "Select"
' dlg.Filters = strFilter

' we just need the folder name really
' SourceFilename = dlg.Execute

    ' value if none chosen is empty string
    strFolderChosenByUser = ""
    If dlg.Show Then
        ' credit http://www.mrexcel.com/forum/excel-questions/737619-visual-basic-applications-get-folder-path-using-msofiledialogfolderpicker.html
        strFolderChosenByUser = dlg.SelectedItems(1)

' Set folder = fso.GetFolder(FldPath)
    Else
        ' or would we use = CurDir
    End If
    xlApp.Quit
    Set xlApp = Nothing
End Function



Function arrFilteredPathnamesInUserTree( _
  strFilter As String _
, Optional bRecurse As Boolean = True _
) As String()
' this will return an array of full file and path names to files meeting a filter criteria
' using FileSystemObject from Shell Scripting
' in a folder chosen by the user
' or if none found then returnString(0)=""

    Dim strArrReturn() As String
    Dim intElement As Integer
    Dim SourceFilename As String
    Dim strFolderName As String
    
    intElement = 0
    ReDim strArrReturn(0)
    ' default value if none found
    strArrReturn(0) = ""
    
    strFolderName = strFolderChosenByUser("Please choose a folder")
    
    If strFolderName <> "" Then
    
        ' assuuming strFilter is single element but delimited (e.g. ; or | ), break it into array for easier match looping
        
        ' first add the current
        AddMatchingNamesFromFolderToArray strArrReturn, strFolderName, strFilter, intElement
        
        If bRecurse Then ' do tree not just folder
            Dim fso As Scripting.FileSystemObject
            Dim fsoFolder As Scripting.folder
            Dim fsoSubFolder As Scripting.folder
            
            Set fso = New Scripting.FileSystemObject
            Set fsoFolder = fso.GetFolder(strFolderName)
            'Application.ScreenUpdating = False

            For Each fsoSubFolder In fsoFolder.SubFolders
                AddMatchingNamesFromFolderToArray strArrReturn, fsoSubFolder.Path, strFilter, intElement
            Next fsoSubFolder
        End If
    End If
    
    arrFilteredPathnamesInUserTree = strArrReturn
End Function


Function AddMatchingNamesFromFolderToArray(strArray() As String, strFolderName As String, strFilter As String, intElement As Integer)
    Dim fso As Scripting.FileSystemObject
    Dim fsoFolder As Scripting.folder
    Dim fsoFile As Scripting.file
    Set fso = New Scripting.FileSystemObject
    
    Set fsoFolder = fso.GetFolder(strFolderName)
    
    For Each fsoFile In fsoFolder.files
        
        ' check against each of the filters in the array
        ' ONLY DOES ONE for the moment
        If LCase(Right(fsoFile.Name, Len(strFilter))) = LCase(strFilter) Then
        
            ' as redimming each item affects performance,
            ' consider doing it say 10 or 100 at a time then shrinking at the end
            ReDim Preserve strArray(intElement)

            strArray(intElement) = strFolderName & cStrPathSeparator & fsoFile.Name
            intElement = intElement + 1
        End If
    Next fsoFile

End Function


' from mod_exc_FileLocations FindParentFolderFromPath 130828.AMG
Function FindParentFolderFromPath(strFullPath As String, Optional theSlash As String = "\") As String
    FindParentFolderFromPath = Left(strFullPath, InStrRev(strFullPath, theSlash) - 1)
End Function




' *********** SHAREPOINT FUNCTIONS *********************************************


' was mod_Acc_ExportToSharePoint
Public Function DownloadFromSharePoint(strSharePointURL As String, strLocalPathFileName As String) As Long
' simple wrapper function
    Dim lngReturn  As Long
    lngReturn = URLDownloadToFile(0, strSharePointURL, strLocalPathFileName, 0, 0)
    DownloadFromSharePoint = lngReturn
End Function


' was mod_Acc_ExportToSharePoint
Function SharePointCheckIfFileExists(URLStr As String) As Boolean
    
    ' credit > http://stackoverflow.com/questions/13493756/is-it-possible-to-check-via-vba-if-file-exist-on-a-sharepoint-site
    Dim oHttpRequest As Object
    Set oHttpRequest = New MSXML2.ServerXMLHTTP60
    With oHttpRequest
        .Open "GET", URLStr, False
        .setRequestHeader "Cache-Control", "no-cache"
        .setRequestHeader "Pragma", "no-cache"
        .setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
        .send
    End With
    If oHttpRequest.Status = 200 Then
        SharePointCheckIfFileExists = True
    Else
        SharePointCheckIfFileExists = False
    End If

End Function





''''''''''''''
' Credit > http://allenbrowne.com/ser-59.html
' Alternatives > http://my.advisor.com/doc/16279
'
''' START OF COPIED CODE ''''''''''''''''''''''''''''
'
'
Public Function ListFiles(strPath As String, Optional strFileSpec As String, _
    Optional bIncludeSubfolders As Boolean, Optional lst As ListBox)
On Error GoTo Err_Handler
    'Purpose:   List the files in the path.
    'Arguments: strPath = the path to search.
    '           strFileSpec = "*.*" unless you specify differently.
    '           bIncludeSubfolders: If True, returns results from subdirectories of strPath as well.
    '           lst: if you pass in a list box, items are added to it. If not, files are listed to immediate window.
    '               The list box must have its Row Source Type property set to Value List.
    'Method:    FilDir() adds items to a collection, calling itself recursively for subfolders.
    Dim colDirList As New Collection
    Dim varItem As Variant
    
    Call FillDir(colDirList, strPath, strFileSpec, bIncludeSubfolders)
    
    'Add the files to a list box if one was passed in. Otherwise list to the Immediate Window.
    If lst Is Nothing Then
        For Each varItem In colDirList
            Debug.Print varItem
        Next
    Else
        For Each varItem In colDirList
        lst.AddItem varItem
        Next
    End If

Exit_Handler:
    Exit Function

Err_Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume Exit_Handler
End Function

Private Function FillDir(colDirList As Collection, ByVal strFolder As String, strFileSpec As String, _
    bIncludeSubfolders As Boolean)
    'Build up a list of files, and then add add to this list, any additional folders
    Dim strTemp As String
    Dim colFolders As New Collection
    Dim vFolderName As Variant

    'Add the files to the folder.
    strFolder = TrailingSlash(strFolder)
    strTemp = Dir(strFolder & strFileSpec)
    Do While strTemp <> vbNullString
        colDirList.Add strFolder & strTemp
        strTemp = Dir
    Loop

    If bIncludeSubfolders Then
        'Build collection of additional subfolders.
        strTemp = Dir(strFolder, vbDirectory)
        Do While strTemp <> vbNullString
            If (strTemp <> ".") And (strTemp <> "..") Then
                If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0& Then
                    colFolders.Add strTemp
                End If
            End If
            strTemp = Dir
        Loop
        'Call function recursively for each subfolder.
        For Each vFolderName In colFolders
            Call FillDir(colDirList, strFolder & TrailingSlash(vFolderName), strFileSpec, True)
        Next vFolderName
    End If
End Function

Public Function TrailingSlash(varIn As Variant) As String
    If Len(varIn) > 0& Then
        If Right(varIn, 1&) = "\" Then
            TrailingSlash = varIn
        Else
            TrailingSlash = varIn & "\"
        End If
    End If
End Function
'
'
''' END OF COPIED CODE '''''''''''''''''''''''''''''''''''''






