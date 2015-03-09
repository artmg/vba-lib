Attribute VB_Name = "mod_vsd_ShapesLinks"
' mod_vsd_ShapesLinks
' 150303.AMG


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
            AddHyperlinkDetailToList hlk
        Next
    End If
    
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


