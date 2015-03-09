Attribute VB_Name = "mod_vsd_FromExcel"
Option Explicit

' References
' ==========
'
' This module may require the following references (paths and GUIDs might vary)
'
' Excel (C:\Program Files\Microsoft Office\OFFICE11\EXCEL.EXE) 1.5 - {00020813-0000-0000-C000-000000000046}

Enum entryType
    Environments = 1
    Hosts = 2
End Enum


Sub Import_From_EnvTables()
    Dim xlwbk As Excel.Workbook
    Set xlwbk = xlwbkFromBrowse(bDebug:=False)
    If xlwbk Is Nothing Then
        MsgBox "No Excel Workbook selected"
    Else
'        Application.StatusBar = "checking environments..."
        CheckEntries _
            entryType.Environments, _
            xlwbk.Sheets("Environments"), _
            Application.ActiveDocument.Pages("Environments")
        
        CheckEntries _
            entryType.Hosts, _
            xlwbk.Sheets("Hosts"), _
            Application.ActiveDocument.Pages("Environments")
        
        AddHeadFoot _
            Application.ActiveDocument.Pages("Environments")

'        Application.StatusBar = False
        xlwbk.Close SaveChanges:=False
        Set xlwbk = Nothing
    End If
End Sub

Function CheckEntries(typ As entryType, sht As Excel.Worksheet, pg As Page)
    Dim rw As Excel.Range
    Dim boolFirst As Boolean

    boolFirst = True
    For Each rw In sht.UsedRange.Rows
        'skip the first (header) row
        If boolFirst Then
            boolFirst = False
        Else
            Dim strEntry, strParent, strOrder, strInfo1, strInfo2 As String
            strEntry = rw.Cells(1, 1).Text
            If shpFromNameIfExists(pg, strEntry) Is Nothing Then
                If typ = Environments Then
                    strOrder = rw.Cells(1, 5).Text
                    strInfo1 = rw.Cells(1, 4).Text
                    AddEnvironment pg, strEntry, strOrder, strInfo1
                Else
                    If rw.Cells(1, 3).Text = "I" Then ' only add `Internal` hosts
                        strOrder = rw.Cells(1, 8).Text
                        strParent = rw.Cells(1, 2).Text
                        strInfo1 = rw.Cells(1, 7).Text
                        AddHost pg, strEntry, strParent, strOrder, strInfo1
                    End If
                End If
 '               cnt = cnt + 1
            End If
        End If
    Next
'    Debug.Print cnt & " " & CStr(typ) & " entries added"
End Function

Function AddHeadFoot(pg As Page)
    Dim shp As Shape
    Dim thePage As Shape
    Set thePage = pg.Shapes("thePage")
    
    Set shp = pg.DrawRectangle(0, thePage.Cells("PageHeight"), thePage.Cells("PageWidth"), thePage.Cells("PageHeight") - 1)
    shp.Name = "PageHeader"
    shp.Text = "Environments and Hosts"
    shp.Cells("Char.Size").Formula = "=24pt"
    shp.Cells("Char.Style").Formula = "=17"
    shp.Cells("FillPattern").Formula = "=0"
    shp.Cells("LinePattern").Formula = "=0"

    Set shp = pg.DrawRectangle(1, 0, thePage.Cells("PageWidth") - 1, 0.5)
    shp.Name = "PageFooter"
    shp.Text = Format(Now(), "dd mmm yyyy")
    shp.Cells("Char.Size").Formula = "=12pt"
    shp.Cells("Char.Style").Formula = "=21"
    shp.Cells("FillPattern").Formula = "=0"
    shp.Cells("LinePattern").Formula = "=0"
    shp.Cells("Para.HorzAlign").Formula = "=2"

End Function

Function AddEnvironment(ByRef pg As Page, ByVal strEnvCode As String, ByVal strOrder As String, ByVal strInfo1 As String)
' Size, Distance, Starting Offset and Direction of shapes
Const cShapeHeight As Double = 1.5
Const cShapeWidth As Double = 2
Const cShapeVDist As Double = 0.1
Const cShapeHDist As Double = 0.1
Const cShapeVOff As Double = -1
Const cShapeHOff As Double = -1
Const cShapeVDir As Double = 1
Const cShapeHDir As Double = 1

    ' The new column and row are in the strOrder as a numberic C.R
    Dim dblInfo As Double
    If IsNumeric(strOrder) Then
        dblInfo = CDbl(strOrder)
    Else
        dblInfo = 0
    End If
    Dim intShapeCol, intShapeRow As Integer
    intShapeCol = CInt(Int(dblInfo))
    intShapeRow = CInt((dblInfo - Int(dblInfo)) * 10)
    
    ' calculate the position of the new shape
    Dim dblShapeVPos, dblShapeHPos As Double
    dblShapeVPos = cShapeVOff + (cShapeVDir * intShapeRow * (cShapeVDist + cShapeHeight))
    dblShapeHPos = cShapeHOff + (cShapeHDir * intShapeCol * (cShapeHDist + cShapeWidth))
    
    Dim shp As Shape
'    If offset = 0 Then offset = EnvNewRow
'    EnvNewPos = EnvNewRow - offset * (EnvHeight + EnvNewVDist)
    Set shp = pg.DrawRectangle(dblShapeHPos, dblShapeVPos, dblShapeHPos + cShapeWidth, dblShapeVPos + cShapeHeight)
    shp.Name = strEnvCode
    shp.Text = strEnvCode
    shp.Cells("Char.Size").Formula = "=14pt"
    shp.Cells("Char.Style").Formula = "=17"
'    shp.CellsSRC(visSectionCharacter, visRowCharacter + 0, visCharacterSize) = "14pt"
    shp.Cells("VerticalAlign").Formula = "=0"
'    shp.Cells("Geometry1.NoFill").Formula = "=True"

    Dim strFill As String
    Select Case strInfo1
        Case "In Progress": strFill = "=RGB(255,255,0)"
        Case "Not Started": strFill = "=RGB(204,204,204)"
        Case "Complete": strFill = "=RGB(0,255,0)"
        Case Else: strFill = "=RGB(255,255,255)"
    End Select
    shp.Cells("FillForegnd").Formula = strFill
End Function

Function AddHost(ByRef pg As Page, ByVal strHostname As String, ByVal strEnvCode As String, ByVal strOrder As String, ByVal strInfo As String)
' Size, Distance, Starting Offset and Direction of shapes
Const cShapeHeight As Double = 0.15
Const cShapeWidth As Double = 1.8
Const cShapeVDist As Double = 0.05
Const cShapeHDist As Double = 0.05
Const cShapeVOff As Double = -0.13
Const cShapeHOff As Double = 0.1
Const cShapeVDir As Double = 1
Const cShapeHDir As Double = 0

    Dim intShapeCol, intShapeRow As Integer
    If IsNumeric(strOrder) Then
        intShapeRow = CInt(strOrder)
    Else
        intShapeRow = 0
    End If
    intShapeCol = 0

'    Dim RectNewPos As Double
    Dim shpParent, shp As Shape

'    Dim ObjNewPos As Double
    
    ' calculate the position of the new shape
    Dim dblShapeVPos, dblShapeHPos As Double
    dblShapeVPos = cShapeVOff + (cShapeVDir * intShapeRow * (cShapeVDist + cShapeHeight))
    dblShapeHPos = cShapeHOff + (cShapeHDir * intShapeCol * (cShapeHDist + cShapeWidth))
    
    Set shpParent = shpFromNameIfExists(pg, strEnvCode)
    If Not shpParent Is Nothing Then
'        RectNewPos = ObjNewRow - (cnt - 1) * (RectHeight + ObjNewVDist)
'        Set shp = shpParent.DrawRectangle(ObjNewCol, RectNewPos, ObjNewCol + RectWidth, ObjNewPos + RectHeight)
'        Set shp = shpParent.DrawRectangle(dblShapeHPos, dblShapeVPos, dblShapeHPos + cShapeWidth, dblShapeVPos + cShapeHeight)
        dblShapeHPos = dblShapeHPos + shpParent.Cells("PinX") - (shpParent.Cells("Width") / 2)
        dblShapeVPos = dblShapeVPos + shpParent.Cells("PinY") - (shpParent.Cells("Height") / 2)
        Set shp = pg.DrawRectangle(dblShapeHPos, dblShapeVPos, dblShapeHPos + cShapeWidth, dblShapeVPos + cShapeHeight)
        shp.Name = strHostname
        shp.Text = strHostname & " " & strInfo
        
' if I can't AddRow to Character Section, how can I format text to different sizes?
'        dim srw as
'        srw = shp.AddRow(visSectionCharacter, visRowLast, visRowCharacter) ' Len(strHostname),
'        shp.Cells("Char.Size" & Len(strHostname)).Formula = "=10pt"
        shp.Cells("Geometry1.NoFill").Formula = "=True"
        shp.Cells("VerticalAlign").Formula = "=1"
'    shp.Cells("FillForegnd").Formula = "=RGB(51,102,255)"
    End If
End Function

Function shpFromNameIfExists(ByRef pg As Page, ByVal strShapeName As String) As Shape
    Dim shp As Shape
    For Each shp In pg.Shapes()
        If shp.Name = strShapeName Then
            Set shpFromNameIfExists = shp
            Exit For
        End If
    Next
End Function

Function xlwbkFromBrowse(Optional bDebug As Boolean = False) As Excel.Workbook
    Dim strWbkFilename As String
    Dim xlwbk As New Excel.Workbook

    If bDebug Then
       strWbkFilename = "C:\Path\tables.XLS"
    Else
        strWbkFilename = CStr(Excel.Application.GetOpenFilename( _
            FileFilter:="Excel Workbooks (*.xls), *.xls", _
            Title:="Please choose the Excel Workbook containing the Environment Management Tables", _
            ButtonText:="Read"))
    End If

    If strWbkFilename <> "False" Then
        ' GetObject is probably not the most efficient way to open files so ...
'        set xlwbk = GetObject(strWbkFilename)
        
        Set xlwbk = Excel.Workbooks.Open( _
            FileName:=strWbkFilename, _
            UpdateLinks:=0, _
            ReadOnly:=True, _
            IgnoreReadOnlyRecommended:=True _
            )
        If Not xlwbk Is Nothing Then
            Set xlwbkFromBrowse = xlwbk
        End If
    End If
End Function

'Sub ScanObjects()
'    Dim shp As Shape
'
'    For Each shp In Application.ActiveDocument.Pages(1).Shapes
'        Debug.Print shp.Name
'
'        shp.Text = shp.Name
'    Next
'
'End Sub
