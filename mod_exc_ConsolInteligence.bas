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

'  150326.AMG  use Equivalents table to combine similar Key values and move Match into DataTables module
'  150324.AMG  minor tweaks and potential improvements
'  150115.AMG  extended with wide option
'  141113.AMG  cribbed from various mod_exc's


' REQUIRES:
'   mod_exc_DataTables
'   mod_exc_WorkbooksSheetsNames


' IMPROVE:
'  move Option Variables out of module into Variables tab
'  routine to create SourceDefinitions AND Variables Tab if do not exist
'  populate it with Path = current path (to show adding trailing slash
'  allow Numerical Add when additional rows match, if column specified as '+N

' PREPARE:
' SourceDefinitions sheet contains columns for:
'   SourceID
'   Path
'   File    (including trailing folder delimeter)
'   Sheet
'
' then a series of Destination Column Names
' first destination column is unique key for Wide option
'
' and finally a column (NOT YET IMPLEMENTED) called
'   Exceptions
' which gives logic stating which rows NOT to import
'
' IF UseEquivalents option is true (ASSUME Wide is also true)
' then also PREPARE:
' KeyEquivalents sheet contains columns for:
'    EquivalentIncorrectKey
'    RefersToCorrectKey
'
' If the 'incorrect' key is NOT found in the consolidated table,
' the corresponding 'Correct' value is sought instead,
' before creating a separate line
'


Option Explicit

Const cStrMultiValDelim As String = "; "

Dim intFirstDefCol, intLastDefCol As Integer
Dim intNextOutRow As Integer
    
Dim intSrcs As Integer ' number of sources and values per source when in Wide mode
Dim intVals As Integer

' Option Variables to pass around
Dim bWide As Boolean
Dim bMatchCase As Boolean
Dim bUseEquivalents As Boolean
Dim strKeyIgnore As String

Sub ConsolidateWithIntelligence()

    Dim shtDefs As Worksheet
    Dim shtEquivs As Worksheet
    Dim shtOutput As Worksheet
    Dim strSourceSheetName As String
    Dim strEquivsSheetName As String
    
    ' set Option Variables
    ' manually for now - IMPROVE move onto worksheet
    bWide = True ' Output will be wide
    bUseEquivalents = True ' use KeyEquivalents sheet
    bMatchCase = False ' Key match will be case insensitive
    strKeyIgnore = "" ' text to strip before matching

    ' IMPROVE - MAKE these variable too
    strSourceSheetName = "SourceDefinitions"
    strEquivsSheetName = "KeyEquivalents"

    Set shtDefs = ActiveWorkbook.Worksheets(strSourceSheetName)
    If bUseEquivalents Then
        Set shtEquivs = ActiveWorkbook.Worksheets(strEquivsSheetName)
    End If
    Set shtOutput = getSheetOrCreateIfNotFound(ActiveWorkbook, "Consolidated")

    ClearEntireSheet shtOutput
    
    intSrcs = shtDefs.UsedRange.Rows.Count - 1 ' does not yet allow for blank rows!
    AddColumnHeadersFrom shtDefs, shtOutput

    Dim intDefRow As Integer
    For intDefRow = 2 To shtDefs.UsedRange.Rows.Count
'        CopyDataFromSourceTo shtOutput, shtDefs.Rows(intDefRow)
        CopyDataFromDefSrcTo shtOutput, shtDefs, shtEquivs, intDefRow
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
    , ByRef shtEquivs As Worksheet _
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
        CopyRowFromSourceTo shtOutput, shtDefs, shtEquivs, intDefRow, shtSource, intSrcRow
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
    , ByRef shtEquivs As Worksheet _
    , ByVal intDefRow As Integer _
    , ByRef shtSource As Worksheet _
    , ByVal intSourceRow As Integer _
)

    Dim enumMatchType As enumDataTableMatchType
    If bMatchCase Then
        enumMatchType = MatchCaseSens
    Else
        enumMatchType = MatchCaseInsens
    End If

    Dim intOutRow, intMatchOutRow, intMatchEquivRow  As Integer
    Dim intSourceCol As Integer
    intSourceCol = CInt(shtDef.Cells(intDefRow, intFirstDefCol).Value)

    ' unless wide mode finds an alternative we will output to next available line
    Dim rngOutRow As Range

    Dim strNewKey, strMatchKey, strEquivKey As String
    intMatchOutRow = 0

    ' when consolidating wide,
    ' search for existing line to consolidate onto
    ' first do match on output
    ' if not found then do match on equivalent
    ' then do match of equivalent on output

    If Not bWide Then
        strNewKey = shtDef.Cells(intDefRow, 1).Value
    Else
        ' first look for strNewKey in shtOutput
        strNewKey = shtSource.Cells(intSourceRow, intSourceCol).Value

        strMatchKey = strMatchPrepareValue _
            (strUnprepared:=strNewKey _
            , enumMatchType:=enumMatchType _
            , strIgnore:=strKeyIgnore _
            )

        intMatchOutRow = intMatchGetRow _
            (strMatch:=strMatchKey _
            , enumMatchType:=enumMatchType _
            , sht:=shtOutput _
            , intCol:=1 _
            , intFirstRow:=2 _
            , intLastRow:=intNextOutRow - 1 _
            , strIgnore:=strKeyIgnore _
            )

        ' If not found then look for strNewKey in shtEquiv
        If (intMatchOutRow = 0) And bUseEquivalents Then
            intMatchEquivRow = intMatchGetRow _
                (strMatch:=strNewKey _
                , enumMatchType:=enumMatchType _
                , sht:=shtEquivs _
                , intCol:=1 _
                , intFirstRow:=2 _
                , intLastRow:=300 _
                , strIgnore:=strKeyIgnore _
                )
            ' WHAT IS intLastRow for END OF EQUIVS - currently fixed at 300!

            ' If Equiv found then look for set strEquivKey
            If intMatchEquivRow <> 0 Then
                ' This assumes New Key WILL be exact value from Equivs
                strNewKey = shtEquivs.Cells(intMatchEquivRow, 2).Value
                ' and leaves strEquivKey variable UNUSED

                strMatchKey = strMatchPrepareValue _
                    (strUnprepared:=strNewKey _
                    , enumMatchType:=enumMatchType _
                    , strIgnore:=strKeyIgnore _
                    )
        
                ' and look for THAT in shtOutput
                intMatchOutRow = intMatchGetRow _
                    (strMatch:=strMatchKey _
                    , enumMatchType:=enumMatchType _
                    , sht:=shtOutput _
                    , intCol:=1 _
                    , intFirstRow:=2 _
                    , intLastRow:=intNextOutRow - 1 _
                    , strIgnore:=strKeyIgnore _
                    )
            End If
        End If
    End If

    ' if we have a match use that row
    If intMatchOutRow <> 0 Then
        intOutRow = intMatchOutRow
        ' AND ASSUMES that NEW KEY will be 'untreated',
        ' NOT the 'prepared' value
'        ' IS KEY ALWAYS FIRST ?
'        strNewKey = shtSource.Cells(intSourceRow, intSourceCol).Value
    Else
        ' else add a new one on the end
        intOutRow = intNextOutRow
        intNextOutRow = intNextOutRow + 1
'        ' CAN WE PULL THIS OUT A LEVEL TO DEDUPE ABOVE?
'        If bWide Then
'            strNewKey = shtSource.Cells(intSourceRow, intSourceCol).Value
'        Else
'            strNewKey = shtDef.Cells(intDefRow, 1).Value
'        End If
    End If


    Set rngOutRow = shtOutput.Rows(intOutRow)
    rngOutRow.Cells(1, 1).Value = strNewKey
'    Else ' if not wide, copy SourceID onto destination
'        Set rngOutRow = shtOutput.Rows(intNextOutRow)
'        strNewKey = shtDef.Cells(intDefRow, 1).Value
'        rngOutRow.Cells(1, 1).Value = shtDef.Cells(intDefRow, 1).Value
'    End If


' SINCE ADDING WIDE option, has this lost the original non-wide functionality????
    Dim intCol As Integer
    Dim bTreatAsNum As Boolean
    For intCol = intFirstDefCol To intLastDefCol
         intSourceCol = CInt(shtDef.Cells(intDefRow, intCol).Value)
        ' if there is a + in the column number, treat the value as a number
         bTreatAsNum = (InStr(CStr(shtDef.Cells(intDefRow, intCol).Value), "+") > 0)
         If intSourceCol > 0 Then
            Dim strNewValue As String
            Dim celExisting, celNew As Range
'shtSource.Rows (intSrcRow)
'            rngOutRow.Cells(1, intCol - intFirstDefCol + 2).Value = rngSrcRow.Cells(1, CInt(rngDefRow.Cells(1, intCol).Value)).Value
            Set celNew = shtSource.Cells(intSourceRow, intSourceCol)
            Set celExisting = rngOutRow.Cells(1, intDestCol(intDefRow - 1, intCol - intFirstDefCol + 1))
            If bTreatAsNum Then
                celExisting.Value = celExisting.Value + celNew.Value
            Else
                strNewValue = celNew.Value
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
        End If
    Next intCol
End Function

