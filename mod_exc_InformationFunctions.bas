Attribute VB_Name = "mod_exc_InformationFunctions"
' mod_exc_InformationFunctions
'
' User Defined Functions (UDF) that add additional valuable functionality to the regular
' Excel Information worksheet functions
'
' 080526.AMG  created
'

' NB: To make the Functions accessible as User Defined Functions in worksheets
'     ensure that this code is inserted into a regular Module (not into Microsoft Excel Objects)


Option Explicit

Public Function IsFormula(r As Range)
    IsFormula = r(1).HasFormula

' This is the VBA alternative to the "Name Formula" technique
' Define a Name called
'     CellHasFormula
' set Refers To as
'     =GET.CELL(48,INDIRECT("rc",FALSE))
' Use Conditional Formatting with Formula Is set to
'     =CellHasFormula
'
'http://www.j-walk.com/ss/excel/usertips/tip045.htm
'http://www.pcmag.com/article2/0,1759,1573749,00.asp

End Function


