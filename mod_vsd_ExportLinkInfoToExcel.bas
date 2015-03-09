Attribute VB_Name = "mod_vsd_ExportLinkInfoToExcel"
' mod_vsd_ExportLinkInfoToExcel
' 150303.AMG

' depends on:
'   mod_vsd_ShapesLinks

Option Explicit

Public Sub OutputLinkDetailsToWorksheet()
        
    ExcelOutputCreateWorksheet
    RecurseAllShapesInDoc ActiveDocument
End Sub


Function AddHyperlinkDetailToList(hlk As Hyperlink)
    ExcelOutputWriteValue hlk.Shape.Name
    ExcelOutputWriteValue hlk.Address
    ExcelOutputNextRow

End Function

