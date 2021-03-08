
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

## Truncate at last delimiter

Suppose you have a file path, and want to convert it to a path name only, you would want to 
remove the last delimiter and all that follows.

If VBA you can find from the right, but not in worksheet formulas. Therefore you must use the logic:
* Count the number of delimiters (N)
	* to do this Substitute the delimiters with nul and compare the length before and after
* Substitute the Nth delimeter with a string you won't find elsewhere
* Find this string's position and truncate just before there

So the last delimiter's position is as follows,
unless you have a multicharacter delimiter where you would need to divide the difference

=FIND("^^",SUBSTITUTE(A2,"/","^^",LEN(A2)-LEN(SUBSTITUTE(A2,"/",""))))

To return the first portion

=MID(A2,1,FIND("^^",SUBSTITUTE(A2,"/","^^",LEN(A2)-LEN(SUBSTITUTE(A2,"/",""))))-1)

and to return the last portion

=RIGHT(A2,LEN(A2)-FIND("^^",SUBSTITUTE(A2,"/","^^",LEN(A2)-LEN(SUBSTITUTE(A2,"/","")))))

Replace A2 with the cell you want and "/" with your delimiter
NB: There is no error checking in these formulas (if no delimiter found)
If A2 contains 'escaped' delimiters you will need to substitute them out in all occurrences of A2


## Transpose Vertical table to rows

This gives rows of formulas that can be pulled down on the right of the original data
and where delimted, the data is transposed dynamically.
It can then be lifted out using Copy / Paste Values and Sort to remove blank lines

First identify which line becomes the start of the row

=IF(OFFSET(A2,1,0)="Start of Data Marker","Y","")

Then fill across after that to pick up subsequent lines 

=IF($H2="","",IF(COUNTIF(OFFSET($H2,0,0,COLUMNS($T2:T2)+3,1),"Y")<>1,"",IF(INDEX($A2:$A32,COLUMNS($T2:T2)+3)="","",INDEX($A2:$A32,COLUMNS($T2:T2)+3))))


