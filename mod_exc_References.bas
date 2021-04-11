Attribute VB_Name = "mod_exc_References"
Option Explicit

' 130821.AMG

Sub ListReferences()
    Debug.Print "' References"
    Debug.Print "' =========="
    Debug.Print "'"
    Debug.Print "' This module uses the following references (paths and GUIDs may vary)"
    Debug.Print "'"
    Dim ref As Variant
    For Each ref In ThisWorkbook.VBProject.References
        Debug.Print "' " & ref.Description & " (" & ref.fullpath & ") " & ref.GUID
    Next
End Sub

Public Sub RemoveProject14Reference()
    Dim ref, refs As Variant
    Set refs = ActiveWorkbook.VBProject.References
    For Each ref In refs
        If ref.GUID = "{A7107640-94DF-1068-855E-00DD01075445}" Then
            Debug.Print "Found " & ref.GUID
            ' source > http://support.microsoft.com/kb/308340
            refs.Remove ref
            'refs.AddFromFile "C:\Program Files\Microsoft Office\Office12\MSPRJ.OLB"
            '   MSProject   Microsoft Project 12.0 Object Library   (C:\Program Files\Microsoft Office\Office12\MSPRJ.OLB)
        End If
    Next
End Sub

Public Sub AddProject12Reference()
    Dim ref, refs As Variant
    Set refs = ActiveWorkbook.VBProject.References
    refs.AddFromFile "C:\Program Files\Microsoft Office\Office12\MSPRJ.OLB"
        '   MSProject   Microsoft Project 12.0 Object Library   (C:\Program Files\Microsoft Office\Office12\MSPRJ.OLB)
End Sub


