Attribute VB_Name = "mod_vsd_ExportLinkInfoToExcel"
' mod_vsd_ExportLinkInfoToExcel

' 150414.AMG allow trial run for hyperlink update tests
' 150413.AMG ignore empty hyperlinks (descr and url blank)
' 150316.AMG added headers and better descriptive columns
' 150303.AMG created

' depends on:
'   mod_vsd_ShapesLinks

Option Explicit

Public strCurrentFileFolder As String
Public strCurrentFileNameOnly As String
Public bTrialRun As Boolean

Public Sub OutputLinkDetailsToWorksheet()
    bTrialRun = True
    Dim strFileNames() As String
    strFileNames() = arrFilteredPathnamesInUserTree(strFilter:=".vsd", bRecurse:=True)
' func to return the number of elements without error (0 if none)
    If strFileNames(0) <> "" Then
        PrepareListWithHeaders
        Dim ifile As Integer
        For ifile = 0 To UBound(strFileNames)
            strCurrentFileFolder = GetFolderFromFileName(strFileNames(ifile))
            strCurrentFileNameOnly = JustFileName(strFileNames(ifile))
            AddDiagramToList
            VisioOpenAndRecurseAllShapesInDoc strFileNames(ifile), Not (bTrialRun)
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
    ExcelOutputWriteValue "NewURL"
    ExcelOutputNextRow
End Function

Function AddDiagramToList()
    ExcelOutputWriteValue strCurrentFileFolder
    ExcelOutputWriteValue strCurrentFileNameOnly
    ExcelOutputWriteValue ""
    ExcelOutputNextRow
End Function

Function AddHyperlinkDetailToList( _
  hlk As Hyperlink _
, Optional strNewAdress As String = "" _
)
    ' ignore empty hyperlinks
    If hlk.Description & hlk.Address <> "" Then
        ExcelOutputWriteValue strCurrentFileFolder
        ExcelOutputWriteValue strCurrentFileNameOnly
        ExcelOutputWriteValue hlk.Shape.Name
        ExcelOutputWriteValue hlk.Shape.Text
        ExcelOutputWriteValue hlk.Description
        ExcelOutputWriteValue hlk.Address
        ExcelOutputWriteValue strNewAdress
        ExcelOutputNextRow
    End If
End Function


