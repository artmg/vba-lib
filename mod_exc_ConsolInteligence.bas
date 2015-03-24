Attribute VB_Name = "mod_exc_ConsolInteligence"

' mod_exc_ConsolInteligence
'
' Consolidate tables intelligently from multiple sources
'
' Use a table defining the source table locations and layouts
' to pull specified columns from various sources
' into one single long table for analysis
'
' Now with Wide option, with a single line per unique index value
' and multiple columns for various source values
'
' (c) Join the Bits ltd

'  150324.AMG  minor tweaks and potential improvements
'  150115.AMG  extended with wide option
'  141113.AMG  cribbed from various mod_exc's


' REQUIRES:
'   mod_exc_WorkbooksSheetsNames


' IMPROVE:
' mention trailing slash
' X make output name generic
'
'  move Option Variables out of module into Variables tab
'  routine to create SourceDefinitions AND Variables Tab if do not exist
'  populate it with Path = current path (to show adding trailing slash

' PREPARE:
' SourceDefinitions sheet contains columns for:
'   SourceID
'   Path
'   File
'   Sheet
'
' then a series of Destination Column Names
' first destination column is unique key for Wide option
'
' and finally a column called
'   Exceptions
' which gives logic stating which rows NOT to import

Option Explicit

Const cStrMultiValDelim As String = "; "

Dim intFirstDefCol, intLastDefCol As Integer
Dim intNextOutRow As Integer
    
Dim intSrcs As Integer ' number of sources and values per source when in Wide mode
Dim intVals As Integer

' Option Variables to pass around
Dim bWide As Boolean
Dim bMatchCase As Boolean
Dim strKeyIgnore As String

Sub ConsolidateWithIntelligence()

    Dim shtDefs As Worksheet
    Dim shtOutput As Worksheet
    
    ' set Option Variables
    ' manually for now - IMPROVE move onto worksheet
    bWide = True ' Output will be wide
    bMatchCase = False ' Key match will be case insensitive
    strKeyIgnore = "" ' text to strip before matching
    ' IMPROVE - MAKE these variable too
    Set shtDefs = ActiveWorkbook.Worksheets("SourceDefinitions")
    Set shtOutput = getSheetOrCreateIfNotFound(ActiveWorkbook, "Consolidated")

    ClearEntireSheet shtOutput
    
    intSrcs = shtDefs.UsedRange.Rows.Count - 1 ' does not yet allow for blank rows!
    AddColumnHeadersFrom shtDefs, shtOutput

    Dim intDefRow As Integer
    For intDefRow = 2 To shtDefs.UsedRange.Rows.Count
'        CopyDataFromSourceTo shtOutput, shtDefs.Rows(intDefRow)
        CopyDataFromDefSrcTo shtOutput, shtDefs, intDefRow
    Next intDefRow
End Sub


Function AddColumnHeadersFrom( _
    shtDefs As Worksheet _
    , shtOutput As Worksheet _
)
' also sets first and last column numbers from SourceDefinition
' ***** is this now redundant ?? ******

    intVals = 0
    intFirstDefCol = 5
    
    ' count the number of Value items until "Exceptions" found
    ' this will be returned ByRef to the calling function
    Dim intCol As Integer
    For intCol = intFirstDefCol To shtDefs.UsedRange.Columns.Count
        If shtDefs.Cells(1, intCol).Value <> "Exceptions" Then
            intVals = intCol - intFirstDefCol + 1
' make this redundant
intLastDefCol = intCol
        Else
            intCol = 9999
        End If
    Next

    ' make a single copy of each header, unless in Wide mode
    Dim intDefs, intDef As Integer
    If bWide Then
        intDefs = intSrcs ' one copy of headers per source
        ' First header is Index Field Header
        shtOutput.Cells(1, 1).Value = shtDefs.Cells(1, intFirstDefCol).Value
    Else
        intDefs = 1 ' single copy of headers
        ' first header is Source ID header
        shtOutput.Cells(1, 1).Value = shtDefs.Cells(1, 1).Value
    End If

    For intDef = 1 To intDefs
        For intCol = 1 To intVals
            Dim strHeadText As String
            strHeadText = shtDefs.Cells(1, intCol + intFirstDefCol - 1).Value
            If bWide Then ' prepend with SourceID
                If intCol = 1 Then ' first set of cols is "is present" not copied values
                    strHeadText = shtDefs.Cells(1 + intDef, 1).Value
                Else
                    strHeadText = shtDefs.Cells(1 + intDef, 1).Value + "_" + strHeadText
                End If
            End If
            shtOutput.Cells(1, intDestCol(intDef, intCol)).Value = strHeadText
        Next intCol
    Next intDef

    intNextOutRow = 2
End Function

Function intDestCol(intDef As Integer, intCol As Integer) As Integer
    If bWide Then
'        intDestCol = (intCol - 1) * intSrcs + intCopy + 1
        intDestCol = (intCol - 1) * intSrcs + intDef + 1
    Else
        intDestCol = intCol + 1
    End If
End Function


Function CopyDataFromDefSrcTo( _
    ByRef shtOutput As Worksheet _
    , ByRef shtDefs As Worksheet _
    , ByVal intDefRow As Integer _
)

'    Dim rngDefRow As Range
    Dim strSourceFile As String
    Dim wbk As Workbook
    Dim shtSource As Worksheet

    strSourceFile = shtDefs.Cells(intDefRow, 2).Value & shtDefs.Cells(intDefRow, 3).Value
    On Error GoTo InvalidSource
    Set wbk = Application.Workbooks.Open(strSourceFile, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False, CorruptLoad:=False)
    Set shtSource = wbk.Worksheets(shtDefs.Cells(intDefRow, 4).Value)
    On Error GoTo 0

    Dim intSrcRow As Integer
    For intSrcRow = 2 To shtSource.UsedRange.Rows.Count
        CopyRowFromSourceTo shtOutput, shtDefs, intDefRow, shtSource, intSrcRow
'        CopyRowFromSourceTo shtOutput.Rows(intNextOutRow), rngDefRow, shtSource.Rows(intSrcRow)
    Next intSrcRow

    GoTo Continue:

InvalidSource:
    shtOutput.Cells(intNextOutRow, 1).Value = strSourceFile
    shtOutput.Cells(intNextOutRow, 2).Value = "*** INVALID SOURCE! ***"
    intNextOutRow = intNextOutRow + 1

Continue:
    If Not wbk Is Nothing Then wbk.Close
End Function

Function CopyRowFromSourceTo( _
    ByRef shtOutput As Worksheet _
    , ByRef shtDef As Worksheet _
    , ByVal intDefRow As Integer _
    , ByRef shtSource As Worksheet _
    , ByVal intSourceRow As Integer _
)
    ' unless wide mode finds an alternative we will output to next available line
    Dim rngOutRow As Range
    Set rngOutRow = shtOutput.Rows(intNextOutRow)
    intNextOutRow = intNextOutRow + 1

    If bWide Then ' search for existing line to consolidate onto
        Dim intMatchOutRow As Integer
        Dim strNewKey, strCheckKey As String
        ' lookup new key value
        strNewKey = strPrepareKeyForMatch(shtSource.Cells(intSourceRow, CInt(shtDef.Cells(intDefRow, intFirstDefCol).Value)).Value)
        If strNewKey = "" Then
            intNextOutRow = intNextOutRow - 1   ' don't need new row any more
            GoTo keyEmpty:
        End If
        For intMatchOutRow = 2 To (intNextOutRow - 1)
            strCheckKey = strPrepareKeyForMatch(shtOutput.Cells(intMatchOutRow, 1))
            If strNewKey = strCheckKey Then
                Set rngOutRow = shtOutput.Rows(intMatchOutRow)
                intNextOutRow = intNextOutRow - 1   ' don't need new row any more
                intMatchOutRow = intNextOutRow ' break out
            End If
        Next
        rngOutRow.Cells(1, 1).Value = strNewKey
    Else ' if not wide, copy SourceID onto destination
        rngOutRow.Cells(1, 1).Value = shtDef.Cells(intDefRow, 1).Value
    End If

    Dim intCol As Integer
    For intCol = intFirstDefCol To intLastDefCol
        If CInt(shtDef.Cells(intDefRow, intCol).Value) > 0 Then
            Dim strNewValue As String
            Dim celExisting As Range
'shtSource.Rows (intSrcRow)
'            rngOutRow.Cells(1, intCol - intFirstDefCol + 2).Value = rngSrcRow.Cells(1, CInt(rngDefRow.Cells(1, intCol).Value)).Value
            strNewValue = shtSource.Cells(intSourceRow, CInt(shtDef.Cells(intDefRow, intCol).Value)).Value
            Set celExisting = rngOutRow.Cells(1, intDestCol(intDefRow - 1, intCol - intFirstDefCol + 1))
            If (CStr(celExisting.Value) <> "") And (strNewValue <> "") Then
'                If bMatchCase Then
'                    strToReplace = strKeyIgnore
'                Else
                    If UCase(strNewValue) <> UCase(celExisting.Value) Then
                        strNewValue = CStr(celExisting.Value) + cStrMultiValDelim + strNewValue
                    End If
'                End If
                ' if not empty
'            rngOutRow.Cells(1, intCheckCol).Value = strNewValue
            End If
            celExisting.Value = strNewValue
        End If
    Next intCol

keyEmpty:
End Function

Function strPrepareKeyForMatch(strCurrentKey) As String
    Dim strKeyToMatch As String
    Dim strToReplace As String
    
    If bMatchCase Then
        strToReplace = strKeyIgnore
    Else
        strKeyToMatch = UCase(strCurrentKey)
        strToReplace = UCase(strKeyIgnore)
    End If
    
    If strToReplace <> "" Then
        strKeyToMatch = Replace(strKeyToMatch, strToReplace, "")
    End If
    strPrepareKeyForMatch = strKeyToMatch
End Function



