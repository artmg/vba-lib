Attribute VB_Name = "mod_off_VBProject"
'Option Explicit

' error handling tag             *******************
Const cStrModuleName As String = "mod_off_VBProject"
'                                *******************
'
' (c) Join the Bits ltd
'

' Utility subs from working with VBA code and objects to do with programming
'
' you must "Trust access to the VBA project object model"
' File tab, click Options, click Trust Center, and then click Trust Center Settings
' in Trust Center dialog box / Macro Settings page / Developer Macro Settings / check the box
'
' help - Document.VBProject Property (Visio) https://msdn.microsoft.com/en-us/library/office/ff765161.aspx
'

'
' 160722.AMG  reformatted header only
' 150324.AMG  enabled for (and test with) Excel VBProj objects
' 150309.AMG  renamed from mod_off_References and added Export
' 150303.AMG  made it generic for any office app
' from mod_acc_References
' 080417.AMG  added quick reference documentor
' 080229.AMG  renamed from mod_vba_references
' 0603xx.AMG  created as mod_vba_references within HART
'

' REFERENCES
' ==========
'
' This module uses the following references (paths and GUIDs may vary)
'   (only those required by it's dependent modules)

' This should work as is, but if you want to extend this code (e.g. under Visio)
' you can access the object model from the IDE if you add the following reference (paths and GUIDs may vary)
'
' VBIDE = Microsoft Visual Basic for Applications Extensibility 5.3 (C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB) {0002E157-0000-0000-C000-000000000046}


' DEPENDENCIES
' ============
'
' This module uses the following vba-lib modules
' AND any References specified within them
'
' vba-lib / mod_off_FilesFoldersSitesLinks
' vba-lib / mod_off_ExportListToExcel
' vba-lib / mod_exc_DataTables
'

' IMPROVEMENTS
' ============
'
' * none for now
'


Sub ListReferences()
    Debug.Print "' References"
    Debug.Print "' =========="
    Debug.Print "'"
    Debug.Print "' This module uses the following references (paths and GUIDs may vary)"
    Debug.Print "'"
    Debug.Print "' Name = Description (FullPath) GUID"
    Debug.Print "' ----   -----------  --------  ----"

    ' Dim ref As Reference
    Dim ref As Variant
    With objCurrentVBProject
        For Each ref In .References
            Debug.Print "' " & ref.Name & " = " & ref.Description & " (" & ref.FullPath & ") " & ref.GUID
    '        If Not ref.BuiltIn Then
    '            Debug.Print "' " & ref.Name & " (" & ref.FullPath & ") " & ref.Major & "." & ref.Minor & " - " & ref.Guid
    '        End If

        Next
    End With
End Sub

Function objCurrentVBProject() As Object
    Select Case Application.Name
        Case "Microsoft Visio":
            Set objCurrentVBProject = Visio.Application.ActiveDocument.VBProject
'        Case "Microsoft Access":
'            Set objCurrentVBProject = Access.Application
        Case "Microsoft Excel":
            Set objCurrentVBProject = Excel.ThisWorkbook.VBProject
'        Case Else
'            Set objCurrentVBProject = VBIDE.VDE.ActiveVBProject
'            Set objCurrentVBProject = Application.Vbe.ActiveVBProject

    End Select
End Function



Public Sub ExportModules()
    Dim bExport As Boolean
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent
    Dim objProj As Object
'    Dim cmpComponent As Object

'    credit http://www.rondebruin.nl/win/s9/win002.htm

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If

    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

    Set objProj = objCurrentVBProject
    With objProj
    
        If .Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
            "not possible to export the code"
        Exit Sub
        End If

        szExportPath = FolderWithVBAProjectFiles & "\"
        
        For Each cmpComponent In .VBComponents
            
            bExport = True
            szFileName = cmpComponent.Name
    
            ''' Concatenate the correct filename for export.
            Select Case cmpComponent.Type
                Case vbext_ct_ClassModule
                    szFileName = szFileName & ".cls"
                Case vbext_ct_MSForm
                    szFileName = szFileName & ".frm"
                Case vbext_ct_StdModule
                    szFileName = szFileName & ".bas"
                Case vbext_ct_Document
                    ''' This is a worksheet or workbook object.
                    ''' Don't try to export.
                    bExport = False
            End Select
            
            If bExport Then
                ''' Export the component to a text file.
                cmpComponent.Export szExportPath & szFileName
                
            ''' remove it from the project if you want
            '''wkbSource.VBProject.VBComponents.Remove cmpComponent
            
            End If
       
        Next cmpComponent
    End With
    
    MsgBox "Export is ready"
End Sub



' required for ExportModules() - credit http://www.rondebruin.nl/win/s9/win002.htm
Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim fso As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("scripting.filesystemobject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    If fso.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VBAProjectFiles"
        On Error GoTo 0
    End If
    
    If fso.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function


' this was from Access - not sure it will compile in all cases
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
