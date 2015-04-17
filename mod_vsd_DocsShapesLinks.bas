Attribute VB_Name = "mod_vsd_DocsShapesLinks"
' mod_vsd_DocsShapesLinks

' 150413.AMG begin doc and hyperlink update code
' 150313.AMG added Doc stuff renamed from mod_vsd_ShapesLinks
' 150303.AMG created


'
' Visio Object Model Overview https://msdn.microsoft.com/en-us/library/cc160740.aspx
' Visio Object Model Reference https://msdn.microsoft.com/en-us/library/office/ff765377(v=office.15).aspx
'
'
' Visio Shapes ***********************
'
' Shapes collections are sub-objects of Page, Master or Shape
' Shapes contained by other Shapes are caused by Grouping (a common occurance) and are known as sub-shapes
'
' Shapes Object https://msdn.microsoft.com/en-us/library/office/ff767583.aspx
'
'
' Visio Hyperlinks ***********************
'
' Hyperlinks collections are sub-objects of Shape
'
' Hyperlinks object https://msdn.microsoft.com/en-us/library/office/ff766930.aspx
' Hyperlink object https://msdn.microsoft.com/en-us/library/office/ff767835.aspx
'


Option Explicit


Function EnumHyperlinks(shp As Shape)
    Dim hlk As Hyperlink
    If shp.Hyperlinks.Count > 0 Then
        For Each hlk In shp.Hyperlinks
            ' DoSomethingWith hlk
'            AddHyperlinkDetailToList hlk
            UpdateHyperlinkDetail hlk
        Next
    End If
    
End Function

Function UpdateHyperlinkDetail(hlk As Hyperlink)
    Dim strNewAddress As String
    ' ignore empty hyperlinks
    If hlk.Description & hlk.Address <> "" Then
        strNewAddress = ""

        AddHyperlinkDetailToList hlk, strNewAddress

'        ExcelOutputWriteValue strCurrentFileFolder
'        ExcelOutputWriteValue strCurrentFileNameOnly
'        ExcelOutputWriteValue hlk.Shape.Name
'        ExcelOutputWriteValue hlk.Shape.Text
'        ExcelOutputWriteValue hlk.Description
'        ExcelOutputWriteValue hlk.Address
'        ExcelOutputWriteValue strNewAddress
'        ExcelOutputNextRow
    
        If Not bTrialRun Then
            hlk.Address = strNewAddress
        End If
    End If
End Function





' Docs



' Docs and Shapes

Function VisioOpenAndRecurseAllShapesInDoc( _
  strFileName As String _
, Optional bSave As Boolean = False _
)
    Dim doc As Document
    Set doc = Application.Documents.Open(strFileName)
    RecurseAllShapesInDoc doc
    If bSave Then
        doc.Save
    End If
    doc.Close
    Set doc = Nothing
End Function



Function RecurseAllShapesInDoc(doc As Document)
    Dim pg As Page
    Dim shp As Shape
    
    For Each pg In ActiveDocument.Pages
        For Each shp In pg.Shapes
            DoEachShapeAndSubShape shp
        Next
    Next

End Function


Function DoEachShapeAndSubShape(shp As Shape)
    Dim subshp As Shape
    
' do the main shape
    EnumHyperlinks shp

' if there are subshapes then recurse into them
    If shp.Shapes.Count() <> 0 Then
        For Each subshp In shp.Shapes
            DoEachShapeAndSubShape subshp
        Next
    End If
End Function


