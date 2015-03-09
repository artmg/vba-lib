Attribute VB_Name = "mod_exc_MediaWikiTableExport"


' credit > http://de.wikipedia.org/wiki/Wikipedia:Textverarbeitung/EXCEL-2003_Tabellenumwandlung_VBA


'
'
' <MS-EXCEL VBA code: format_as_wikitable generates a wiki-Table from a EXCEL-cellrange>
'
' (c) Othmar Lippuner>, 2006, 2007
'     Version V18; last changed 07.4.2011
'licenced under      GNU GENERAL PUBLIC LICENSE at  10 April 2006 by author <Othmar Lippuner>
'                    GNU-License Version from 2, June 1991
'
' Everyone is permitted to copy and distribute verbatim copies
' of this license document, but changing it is not allowed.
'
'Installation:
'            1. Copy the Makrocode into a textfile FORMAT_AS_WIKITABLE.BAS
'            2. Import the macrofile FORMAT_AS_WIKITABLE_V17.BAS into a VBA-project of your EXCEL-File
'
'Usage:
 
'            1. Select the range you wan't to publish in EXCEL
'            2. Execute the macro FORMAT_AS_WIKITABLE
'            3. copy the complete wiki-text in outputtable WIKIOUTPUT into clipboard
'            4. paste the clipboardtext into your wikieditor
'
'    The main formatting attributes of excel are converted into wiki-parameters
'    Some strategies are applied to minimize the wiki-textcode generated, e.g. if possible
'    attributes are written als lineparameter instead of cellparameters thus reducing
'    textvolume and DB-load to the wikiservers, an increasing the readability of the tablecode
'    while editing.
'
' Attributes converted
'              bold
'              italic
'              textsize
'              underline
'              backgroundcolor
'              textcolor
'              horizontalalignment
'              verticalaligment
'              numberformats
'
'
' Attributes not converted
'              character font just uses the standard font settings of your favortie wiki-skin
'              styles
'              borders  just uses the standard border settings of class="wikitable"
'
' not supported features
'              nested table (excel can not do that)
'              connected cells in EXCEL, please dont use connected cells
'              charts or any other graphical gagets
'
'
'Software Requirements
'    Software is tested under EXCEL 2003, should be fine also with EXCEL-2000, its up to you to check it out
'
'    Caution: Any worksheet named "wikioutput" will be deleted, recreated and then overwritten
'             when executing the macro. In other words: By executing the macro 'format_as_wikitablle'
'             you accept that the name and content of this worksheet is reserved to the macro
'            'format_as_wikitablle'.
'
'   Version history
'
'           V10     10.4.2006, released
'           V11     17.4.2006, ernonous formatting corrected
'           V12     26.5.2006, verify that selection is a cellrange
'           V13     28.9.2006, V13: replace linebreaks in cellcontent with a Wiki-<BR>
'           V14     15.2.2006, V14: empty cells get &nbsp for correct rendering of cellheight
'           V15     21.4.2007  V15: class="prettytable" instead of [[Prettytable]]
'           V17     30.7.2007  V17: width and height rounded to integer px
'           V18     07.4.2011  V18: Force numeric content of table to be aligned to the right
'
'    Copyright (C) <2006>  <Othmar Lippuner ,Switzerland>
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin Street, Fifth Floor,
'    Boston, MA  02110-1301, USA
'
'
'
'    format_as_wikitablle.bas version 13, Copyright (C) Othmar Lippuner
'    format_as_wikitablle.bas comes with ABSOLUTELY NO WARRANTY;
'    This is free software, and you are welcome to redistribute it
'    under certain conditions; consult the GNU-Public license for these
'    conditions.
'
'
'
'  <Othmar Lippuner>, 10 April 2006  meet me at [[:de:Benutzer Diskussion:Ollio]]
'
'
Option Explicit
Const co = 1 ' all output is written in column 1
Const VersionID = "V1.8"
Const prettytable = True
Const emptyCell_nbsp = True '<< 5.3.2007
Dim iline As Long
Dim icolumn As Long
Dim os As String
Dim oline As Long 'lineindex in outputtable
Dim iLineMax As Long
Dim iColumnMax As Long
Dim selrange As Range  'inputrange
Dim orange As Range 'outputrange
Dim outtabName As String
Dim tableformatting As String
Dim sh As Worksheet
Dim wasUnderlined As Boolean  ' remember Textdecoration:underline state
 
' document the setting of lookahead attributation in line parameter
' if lineparameter is set then skip over cell-attributation
Dim lineattribut_borders_set                As Boolean
Dim lineattribut_fontsize_set               As Boolean
Dim lineattribut_bold_set                   As Boolean
Dim lineattribut_italic_set                 As Boolean
Dim lineattribut_backgroundcolor_set        As Boolean
Dim lineattribut_fondcolor_set              As Boolean
Dim lineattribut_Halignment_set             As Boolean
Dim lineattribut_Valignment_set             As Boolean
 
Dim lineattribut_borders    As Long
Dim lineattribut_fontsize   As Long
Dim lineattribut_backgroundcolor    As Long
Dim lineattribut_fondcolor  As Long
Dim lineattribut_Halignment As Long
Dim lineattribut_Valignment As Long
 
 
 
 
 
Function hexdigit(wrk As Long) As String
If wrk > 15 Then
  MsgBox "illegal hexdigit value : " & wrk
Else
  Select Case wrk
        Case 0:      hexdigit = "0"
        Case 1:      hexdigit = "1"
        Case 2:      hexdigit = "2"
        Case 3:      hexdigit = "3"
        Case 4:      hexdigit = "4"
        Case 5:      hexdigit = "5"
        Case 6:      hexdigit = "6"
        Case 7:      hexdigit = "7"
        Case 8:      hexdigit = "8"
        Case 9:      hexdigit = "9"
        Case 10:     hexdigit = "A"
        Case 11:     hexdigit = "B"
        Case 12:     hexdigit = "C"
        Case 13:     hexdigit = "D"
        Case 14:     hexdigit = "E"
        Case 15:     hexdigit = "F"
  End Select
  End If
End Function 'hexdigit
 
Function myhex(num As Long) As String
'konvert a 16-Bit long to HEX-String inkl fixecd leading zeros
Dim lastdivisor As Long
Dim divisor As Long
Dim wrk As Long
Dim k As Long
Dim result As String
If wrk > 16 ^ 6 Then
      MsgBox "illegal hexdigit value : " & wrk
    Else
    lastdivisor = 1
    result = ""
    divisor = 16
    For k = 1 To 6
        wrk = (num Mod divisor) \ lastdivisor
        result = hexdigit(wrk) & result
        lastdivisor = divisor
        If k < 7 Then ' avoid overflow
            divisor = divisor * 16
        End If
    Next k
    myhex = result
End If
End Function 'myhex
 
 
Private Sub write_tablehead()
tableformatting = " <hiddentext>generated with [[:de:Wikipedia:Helferlein/VBA-Macro for EXCEL tableconversion]] " & VersionID & "</hiddentext>"
If prettytable Then
   tableformatting = " class=" & """wikitable""" & tableformatting
End If
oline = oline + 1: orange.Cells(oline, 1) = "{|" & tableformatting
End Sub 'write_tablehead
 
Private Sub write_lineheader()
Dim col_lookahead As Long
Dim lineheader As String
lineattribut_borders_set = True
lineattribut_fontsize_set = True
lineattribut_bold_set = True
lineattribut_italic_set = True
lineattribut_backgroundcolor_set = True
lineattribut_fondcolor_set = True
lineattribut_Halignment_set = True
lineattribut_Valignment_set = True
 
' init variables for delta-detection
' xxxx lineattribut_borders = selrange.Cells(iline, 1).Borders
If Not IsNull(selrange.Cells(iline, 1).Font.Size) Then
     lineattribut_fontsize = selrange.Cells(iline, 1).Font.Size
Else
     lineattribut_fontsize = 10 'take default
End If
If Not IsNull(selrange.Cells(iline, 1).Font.Bold) Then
    lineattribut_bold_set = selrange.Cells(iline, 1).Font.Bold
Else
    lineattribut_bold_set = False
End If
If Not IsNull(selrange.Cells(iline, 1).Font.Italic) Then
    lineattribut_italic_set = selrange.Cells(iline, 1).Font.Italic
Else
    lineattribut_italic_set = False
End If
lineattribut_backgroundcolor = selrange.Cells(iline, 1).Interior.Color
lineattribut_fondcolor = selrange.Cells(iline, 1).Font.Color
lineattribut_Halignment = selrange.Cells(iline, 1).HorizontalAlignment
lineattribut_Valignment = selrange.Cells(iline, 1).VerticalAlignment
' loop on line for deltadectection
For col_lookahead = 2 To iColumnMax
' xxxx   If lineattribut_borders <> selrange.Cells(iline, 1).Borders Then
' xxxx      lineattribut_borders_set = False: End If
 
    If Not IsNull(selrange.Cells(iline, col_lookahead).Font.Size) Then
        If lineattribut_fontsize <> selrange.Cells(iline, col_lookahead).Font.Size Then
            lineattribut_fontsize_set = False: End If
    End If
    If Not selrange.Cells(iline, col_lookahead).Font.Bold Then
        lineattribut_bold_set = False: End If
    If Not selrange.Cells(iline, col_lookahead).Font.Italic Then
        lineattribut_italic_set = False: End If
    If lineattribut_backgroundcolor <> selrange.Cells(iline, col_lookahead).Interior.Color Then
        lineattribut_backgroundcolor_set = False:
        End If
    If lineattribut_fondcolor <> selrange.Cells(iline, col_lookahead).Font.Color Then
        lineattribut_fondcolor_set = False: End If
    If lineattribut_Halignment <> selrange.Cells(iline, col_lookahead).HorizontalAlignment Then
        lineattribut_Halignment_set = False: End If
    If lineattribut_Valignment <> selrange.Cells(iline, col_lookahead).VerticalAlignment Then
        lineattribut_Valignment_set = False: End If
Next col_lookahead
lineheader = formatstring_for_a_linecontent
' write linetrailer
oline = oline + 1: orange.Cells(oline, 1) = "|- " & lineheader
End Sub 'write_lineheader
 
Private Sub write_linetrailer()
' write linebuffer to output  ==== anyway sofare it is empty
oline = oline + 1: orange.Cells(oline, 1) = os
' flush the linebuffer
os = ""
End Sub 'write_linetrailer
 
 
 
Function excelHexStr2HTML(str As String) As String
Dim a_str As String
Dim b_str As String
Dim c_str As String
a_str = Left(str, 2)
c_str = Right(str, 2)
b_str = Left(Right(str, 4), 2)
excelHexStr2HTML = c_str & b_str & a_str
End Function
 
Private Function skip_underline(str As String) As String
Dim k As Long
Dim so As String
so = ""
' skip unwanted underscores in EXCEL-transforms
For k = 1 To Len(str)
   If Mid$(str, k, 1) <> "_" Then
        so = so & Mid$(str, k, 1)
   End If
Next k
skip_underline = so
End Function
 
 
Private Function process_cellcontent(cellcontent As String) As String
Const verbose = False
Dim hyperlink As String
'dont use .NumberFormatlocal because it
' returns wrong Dateformatstrings "[$-807]TTTT, T. MMMM JJJJ"; instead of "TTTT, T. MMMM JJJJ;" that won't work with format
With selrange.Cells(iline, icolumn)
If verbose Then
    Debug.Print iline; "/"; icolumn, .NumberFormat, .Value
End If
If .NumberFormat <> "General" And .NumberFormat <> "Standard" Then
     cellcontent = skip_underline(Format(.Value, .NumberFormat))
Else
    If cellcontent = "" Then       '<< 15.2.2007
        If Not emptyCell_nbsp Then '<< 05.3.2007
            cellcontent = " " '<< 05.3.2007
        Else                       '<< 05.3.2007
            cellcontent = "&nbsp;" '<< 15.2.2007
        End If                     '<< 05.3.2007
    Else
        cellcontent = cellcontent
    End If                     '<< 15.2.2007
End If
 
' Process hyperlinks
'----------------------------------------
If .Hyperlinks.Count > 0 Then
    hyperlink = .Hyperlinks(1).Address
    If Len(WorksheetFunction.Substitute(hyperlink, "http://", "")) <> Len(hyperlink) Then 'There may be a neater way to do this
        cellcontent = " [" & hyperlink & " " & cellcontent & "]" 'http link
    Else
        cellcontent = " [[" & hyperlink & "|" & cellcontent & "]]" 'assume that anything without http is a local wiki link
    End If
End If
 
End With
' V13: replace linebreaks in cellcentent with a Wiku-<BR> to avoid havoc in wiki-rendering
'      thanks feedback of ManWing2, 26. Sep 2006
process_cellcontent = Replace(cellcontent, vbLf, "<br />")
End Function
 
Private Sub writefirstlinecell(colnr As Long)
With selrange.Cells(iline, icolumn)
    If .MergeArea.Column = .Column And .MergeArea.Row = .Row Then
        oline = oline + 1: orange.Cells(oline, 1) = formatstring_for_a_cellcontent(True, colnr = 1) & " | " & _
                                                    process_cellcontent(selrange.Cells(iline, icolumn))
    End If
End With
End Sub
 
Private Sub writecell(colnr As Long)
With selrange.Cells(iline, icolumn)
    If .MergeArea.Column = .Column And .MergeArea.Row = .Row Then
        oline = oline + 1: orange.Cells(oline, 1) = formatstring_for_a_cellcontent(False, colnr = 1) & " | " & _
                                                    process_cellcontent(selrange.Cells(iline, icolumn))
    End If
End With
End Sub
 
Private Sub write_tabletail()
oline = oline + 1: orange.Cells(oline, 1) = "|}"
End Sub
 
 
Function doublequotestring(str As String, Placeholderchar As String) As String
Dim k As Long
Dim so As String
so = ""
For k = 1 To Len(str)
   If Mid$(str, k, 1) = Left(Placeholderchar, 1) Then
        so = so & """"
   Else
        so = so & Mid$(str, k, 1)
   End If
Next k
doublequotestring = so
End Function
 
 
Function WorksheetExits(tabname As String) As Boolean
Dim found As Boolean
found = False
On Error GoTo err_exit
Worksheets(tabname).Select
found = True
err_exit:
WorksheetExits = found
End Function 'WorksheetExits
 
Public Sub Format_as_wikitable()
' implicit parameter: selected range
' writes the output into table: wikioutput
' caution if this table exists it is deleted !!!
 
 
If Not TypeOf Selection Is Range Then
    MsgBox "Error: You must select a cellrange, to convert to a wiki-table, but you " _
    & vbCrLf & " have selected a " & TypeName(Selection)
Else
    Set selrange = Selection
    wasUnderlined = False
    iLineMax = selrange.Rows.Count
    iColumnMax = selrange.Columns.Count
    outtabName = "wikioutput"
    If WorksheetExits(outtabName) Then
       Worksheets(outtabName).Delete
    End If
    oline = 0
    ' create output worksheet
    Set sh = Worksheets.Add(ActiveWorkbook.Sheets(1), , , xlWorksheet) 'always add  Worksheets(outtabName) at first place
    sh.Name = outtabName 'was Worksheets(1).name = outtabName
    sh.Select
    Set orange = sh.Range(Cells(1, 1), Cells(65353, 1))
    orange.Select
    '( Rows(65534), Columns(1))
    write_tablehead
    For iline = 1 To iLineMax
       write_lineheader
       For icolumn = 1 To iColumnMax
          If iline = 1 Then
           writefirstlinecell (icolumn)
          Else
           writecell (icolumn)
          End If
       Next icolumn
       write_linetrailer
    Next iline
    write_tabletail
End If 'Not TypeOf selrange Is Range Then
End Sub
 
 
Function formatstring_for_a_cellcontent(firstline As Boolean, firstrow As Boolean) As String
Dim str As String
Dim stylestring As String
Dim attribute_String As String
Dim colhexval As String
Dim prop As String
stylestring = ""
attribute_String = ""
With selrange.Cells(iline, icolumn)
   ' Determine backgroundcolor_prop
   '----------------------------------------
   If Not lineattribut_backgroundcolor_set Then
        colhexval = excelHexStr2HTML(myhex(.Interior.Color))
        prop = "@background-color:#" & colhexval
        ' Apply backgroundcolor_prop to Stylestring
        If colhexval <> "FFFFFF" Then 'don't write defaultvalue for white, to help to save wikidb-tablespace
             If stylestring = "" Then
                   stylestring = prop
                Else
                  stylestring = stylestring & ";" & prop
              End If
        End If
   End If
 
    ' Added by Thomas Tausend 4.7.2011
    ' If cell contains a numeric value align to the right!
    If IsNumeric(.Value) Then
        prop = "align=@right@"
        attribute_String = attribute_String & " " & prop
    End If
    ' / Added by Thomas Tausend 4.7.2011

   ' Determine Borders_prop
   '----------------------------------------
   '.Borders
   ' do something
 
   ' Determine Width_prop
   '----------------------------------------
   If firstline Then
      prop = "width=@" & Round(.Width, 0) & "@" '<V17
   ' Apply Width_prop to Stylestring
      attribute_String = attribute_String & " " & prop
    End If
 
   ' Determine Colspan_prop
   '----------------------------------------
   If .MergeArea.Columns.Count > 1 Then
      prop = "colspan=@" & .MergeArea.Columns.Count & "@"
      attribute_String = attribute_String & " " & prop
    End If
 
   ' Determine Rowspan_prop
   '----------------------------------------
   If .MergeArea.Rows.Count > 1 Then
      prop = "rowspan=@" & .MergeArea.Rows.Count & "@"
      attribute_String = attribute_String & " " & prop
    End If
 
      ' Determine Font_prop
   '========================================
   '.Font
   ' Determine Font prop font.size
   '----------------------------------------
    With .Font
       If Not IsNull(.Size) And .Size <> 10 And Not lineattribut_fontsize_set Then  ' trapped ISnull-Condition and ignore standard fontsize
            prop = "font-size:" & .Size
            If stylestring = "" Then
                   stylestring = "@" & prop & "pt"
                Else
                  stylestring = stylestring & ";" & prop & "pt"
             End If
       End If
   ' Determine Font prop font.bold
   '----------------------------------------
       If .Bold And Not lineattribut_bold_set Then
            prop = "font-weight:bold"
            If stylestring = "" Then
                   stylestring = "@" & prop
                Else
                  stylestring = stylestring & ";" & prop
             End If
       End If
      ' Determine Font prop underline
   '----------------------------------------
       If .Italic Then
            prop = "font-style:Italic"
            If stylestring = "" Then
                   stylestring = "@" & prop
                Else
                  stylestring = stylestring & ";" & prop
             End If
       End If
 
 
      ' Determine Font prop font.italic
   '----------------------------------------
       If .Underline = xlUnderlineStyleNone And Not lineattribut_italic_set Then ' toggle switch off
            If wasUnderlined Then  ' toggle switch off
                 prop = "text-decoration:none"
                 wasUnderlined = False ' toggle switch on
                 If stylestring = "" Then
                        stylestring = "@" & prop
                     Else
                       stylestring = stylestring & ";" & prop
                  End If
            End If
       Else '.Underline <> xlUnderlineStyleNone
            If Not wasUnderlined Then
                      prop = "text-decoration:underline"
                      wasUnderlined = True ' toggle switch on
                      If stylestring = "" Then
                             stylestring = "@" & prop
                          Else
                            stylestring = stylestring & ";" & prop
                      End If
             End If
       End If
 
   ' Determine Color prop font.color
   '----------------------------------------
       If Not IsNull(.Color) And .Color <> 0 And Not lineattribut_fondcolor_set Then  ' trapped ISnull-Condition and ignore standard color
            prop = "color:#" & excelHexStr2HTML(myhex(.Color))
            If stylestring = "" Then
                   stylestring = "@" & prop
                Else
                  stylestring = stylestring & ";" & prop
             End If
       End If
    End With
   ' Determine Height_prop
   '----------------------------------------
'   .Height
   If firstrow Then
      prop = "height=@" & Round(.Height, 0) & "@" '<V17
   ' Apply Height_prop to Stylestring
      attribute_String = attribute_String & " " & prop  '<V17
    End If
   ' Determine HorizontalAlignment_prop
   '----------------------------------------
   '.HorizontalAlignment
    If .HorizontalAlignment <> xlHAlignLeft And Not lineattribut_Halignment_set Then ' dont write the default
      prop = ""
      Select Case .HorizontalAlignment
        Case xlHAlignRight:     prop = "align=@right@"
        Case xlHAlignCenter:  prop = "align=@center@"
      End Select
      ' Apply HorizontalAlignment_prop to Stylestring
      attribute_String = attribute_String & " " & prop
      End If
 
   ' Determine VerticalAlignment_prop
   '----------------------------------------
    If .VerticalAlignment <> xlVAlignCenter And Not lineattribut_Halignment_set Then  ' dont write the default
    prop = ""
      Select Case .VerticalAlignment
        Case xlVAlignTop:     prop = "valign=@top@"
        Case xlVAlignBottom:  prop = "valign=@bottom@"
      End Select
      ' Apply VerticalAlignment_prop to Stylestring
      attribute_String = attribute_String & " " & prop
      End If
   ' Determine IndentLevel_prop
   '----------------------------------------
   '.IndentLevel >> maybe later to come
   ' Determine Style_prop
   '----------------------------------------
   '.Style  >> maybe later to come
   '----------------------------------------
   '.WrapText << Attribut is wiki not relevant, while unconditional default
   '----------------------------------------
   '
   If stylestring <> "" Then
       str = doublequotestring("style=" & stylestring & "@", "@")
   End If
   str = str & doublequotestring(attribute_String, "@")
End With
If str <> "" Then
    str = "|" & str
End If
formatstring_for_a_cellcontent = str
End Function 'formatstring_for_a_cellcontent
 
 
 
Function formatstring_for_a_linecontent() As String
Dim prop As String
Dim stylestring As String
Dim colhexval As String
 
Dim attribute_String As String
Dim ostr As String
With selrange.Cells(iline, 1)  'take first column as reference
   ' Determine backgroundcolor_prop
   '----------------------------------------
   If lineattribut_backgroundcolor_set Then
        colhexval = excelHexStr2HTML(myhex(.Interior.Color))
        prop = "@background-color:#" & colhexval
        ' Apply backgroundcolor_prop to Stylestring
        If colhexval <> "FFFFFF" Then 'don't write defaultvalue for white, to help to save wikidb-tablespace
             If stylestring = "" Then
                   stylestring = prop
                Else
                  stylestring = stylestring & ";" & prop
              End If
        End If
   End If
   ' Determine Borders_prop
   '----------------------------------------
   '.Borders
   ' do something
 
      ' Determine Font_prop
   '========================================
   '.Font
   ' Determine Font prop font.size
   '----------------------------------------
    With .Font
       If Not IsNull(.Size) And .Size <> 10 And lineattribut_fontsize_set Then   ' trapped ISnull-Condition and ignore standard fontsize
            prop = "font-size:" & .Size
            If stylestring = "" Then
                   stylestring = "@" & prop & "pt"
                Else
                  stylestring = stylestring & ";" & prop & "pt"
             End If
       End If
   ' Determine Font prop font.bold
   '----------------------------------------
       If lineattribut_bold_set Then
            prop = "font-weight:bold"
            If stylestring = "" Then
                   stylestring = "@" & prop
                Else
                  stylestring = stylestring & ";" & prop
             End If
       End If
      ' Determine Font prop underline
   '----------------------------------------
       If lineattribut_italic_set Then
            prop = "font-style:Italic"
            If stylestring = "" Then
                   stylestring = "@" & prop
                Else
                  stylestring = stylestring & ";" & prop
             End If
       End If
 
 
      ' Determine Font prop font.italic
   '----------------------------------------
       If lineattribut_italic_set Then  ' toggle switch off
                      prop = "text-decoration:underline"
                      wasUnderlined = True ' toggle switch on
                      If stylestring = "" Then
                             stylestring = "@" & prop
                          Else
                            stylestring = stylestring & ";" & prop
                      End If
          End If
 
   ' Determine Color prop font.color
   '----------------------------------------
       If Not IsNull(.Color) And .Color <> 0 And lineattribut_fondcolor_set Then   ' trapped ISnull-Condition and ignore standard color
            prop = "color:#" & excelHexStr2HTML(myhex(.Color))
            If stylestring = "" Then
                   stylestring = "@" & prop
                Else
                  stylestring = stylestring & ";" & prop
             End If
       End If
    End With
   ' Determine Height_prop
   '----------------------------------------
   ' Determine HorizontalAlignment_prop
   '----------------------------------------
   '.HorizontalAlignment
    If .HorizontalAlignment <> xlHAlignLeft And lineattribut_Halignment_set Then  ' dont write the default
      prop = ""
      Select Case .HorizontalAlignment
        Case xlHAlignRight:     prop = "align=@right@"
        Case xlHAlignCenter:  prop = "align=@center@"
      End Select
      ' Apply HorizontalAlignment to Stylestring
      attribute_String = attribute_String & " " & prop
      End If
 
   ' Determine VerticalAlignment_prop
   '----------------------------------------
    If .VerticalAlignment <> xlVAlignCenter And lineattribut_Halignment_set Then   ' dont write the default
    prop = ""
      Select Case .VerticalAlignment
        Case xlVAlignTop:     prop = "valign=@top@"
        Case xlVAlignBottom:  prop = "valign=@bottom@"
      End Select
      ' Apply VerticalAlignment_prop to Stylestring
      attribute_String = attribute_String & " " & prop
      End If
   ' Determine IndentLevel_prop
   '----------------------------------------
   '.IndentLevel >> maybe later to come
   ' Determine Style_prop
   '----------------------------------------
   '.Style  >> maybe later to come
   '----------------------------------------
   '.WrapText << Attribut is wiki not relevant, while unconditional default
   '----------------------------------------
   '
   If stylestring <> "" Then
       ostr = doublequotestring("style=" & stylestring & "@", "@")
   End If
   ostr = ostr & doublequotestring(attribute_String, "@")
End With
'If ostr <> "" Then
'    ostr = "|" & ostr
'End If
formatstring_for_a_linecontent = ostr
 
End Function 'formatstring_for_a_linecontent


