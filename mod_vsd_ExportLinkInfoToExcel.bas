Attribute VB_Name = "mod_vsd_ExportLinkInfoToExcel"
' mod_vsd_ExportLinkInfoToExcel
' 150316.AMG added headers and better descriptive columns
' 150303.AMG created

' depends on:
'   mod_vsd_ShapesLinks

Option Explicit

Dim strCurrentFileFolder As String
Dim strCurrentFileNameOnly As String

Public Sub OutputLinkDetailsToWorksheet()
    Dim strFileNames() As String
    strFileNames() = arrFilteredPathnamesInUserTree(strFilter:=".vsd", bRecurse:=True)
' func to return the number of elements without error (0 if none)
    If strFileNames(0) <> "" Then
        PrepareListWithHeaders
        Dim ifile As Integer
        For ifile = 0 To UBound(strFileNames)
            strCurrentFileFolder = GetFolderFromFileName(strFileNames(ifile))
            strCurrentFileNameOnly = JustFileName(strFileNames(ifile))
            VisioOpenAndRecurseAllShapesInDoc strFileNames(ifile)
        Next
        ExcelOutputShow
    End If
End Sub

Function PrepareListWithHeaders()
    ExcelOutputCreateWorksheet
    ExcelOutputWriteValue "DiagramFolder"
    ExcelOutputWriteValue "DiagramFilename"
    ExcelOutputWriteValue "ShapeName"
    ExcelOutputWriteValue "ShapeText"
    ExcelOutputWriteValue "HyperlinkText"
    ExcelOutputWriteValue "CurrentURL"
    ExcelOutputNextRow
End Function

Function AddHyperlinkDetailToList(hlk As Hyperlink)
    ExcelOutputWriteValue strCurrentFileFolder
    ExcelOutputWriteValue strCurrentFileNameOnly
    ExcelOutputWriteValue hlk.Shape.Name
    ExcelOutputWriteValue hlk.Shape.Text
    ExcelOutputWriteValue hlk.Description
    ExcelOutputWriteValue hlk.Address
    ExcelOutputNextRow
End Function

