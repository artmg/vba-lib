
This file does NOT contain VBA code, but instead useful tips and knowledge about Worksheet Functions.
It gives guidance about using the Built-In functions in Excel to build handy features using plain old worksheets 
without needing any macro code or other 'executables'

## SMALL to dynamically filter excel table

This is an alternative to the CopyTo filtering mechanism, 
to dynamically create a secondary table (and multiple instances) 
based on filtered contents from a main 'source' table, 
as if Querying the data from another place in the current workbook. 

* Use SMALL to filter table using formulas
* SMALL orders an array and picks the nth element
* credit - http://trumpexcel.com/2015/01/dynamic-excel-filter/
    * see also compacted (single formula) equivalent in comments


## COUNT DISTINCT

There is no Out Of the Box (OOB) worksheet function in Excel to do this, 
but there are a number of compound formulas you can useful

If you have numerical values and NO blanks...
=SUM(1/COUNTIF(B2:B111,B2:B111)) 

This one should always work providing you enter it as an array formula (CTRL-SHIFT-ENTER)
=SUM(IF(FREQUENCY(IF(LEN(B2:B111)>0,MATCH(B2:B111,B2:B111,0),""), IF(LEN(B2:B111)>0,MATCH(B2:B111,B2:B111,0),""))>0,1))

See other articles for alternatives

credit - http://www.cpearson.com/excel/Duplicates.aspx
help - http://www.get-digital-help.com/2009/06/09/count-unique-values-in-a-column-in-excel/

## Statistics

For a useful introduction to Statistical Worksheet Functions in Excel, search for 

UCL Excel Statistics Manual workbook on Advanced Excel Statistical Functions and Formulae
