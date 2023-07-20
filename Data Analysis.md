
# Integration

## Peak email output periods

Sending emails is not always a sign of productivity, 
but sometimes you might want the insight of statistical analysis 
based on when you're sending how many emails.

The basis of this is to Connect Excel to your Office 365 Outlook mailbox 
using Power Query, and filter the emails from the Sent Items folder. 

* Excel Menu / Data / Get Data
  * From Online Services / From Microsoft Exchange Online
  * enter the mailbox address
  * choose the Mail table
  * Load
* Insert Pivot Table
  * Values: ID (Count of)
  * Filters: Folder Path
  * Rows: DateTimeSent

## Chunking

### Group data into time blocks

Three options for grouping transaction times into statistically useful blocks 
are outlined clearly in the the article https://www.excelcampus.com/charts/group-times-in-excel/

* Use the Pivot Field Group options 
  * quick and easy for rounding to months, days, hours, etc
  * Right click on the data field (e.g. row) and choose Group
* Calculate a field before pivot
  * Create a new column before you pivot
  * this could be useful for half hours
* Lookup a table before pivot
  * totally flexible grouping
    * e.g. 2 hourly overnight and half hourly during the day
  * use VLOOKUP with TRUE for closest matching feature


* Excel Menu / Query Tools / Query / Edit
  * Add Column / Custom Column


## Anonymisation

See [Lubuild / data-extraction](https://github.com/artmg/lubuild/blob/master/help/manipulate/data-extraction.md#anonymisation)

## Frequency

### Simple Word Frequency in Excel 

Takes a column of cells each containing a short line of text and returns a frequency count in descending order

- Row 1 column headings 
	- Text Words Frequency Range 
- Range formula in D2:  
	- a2:a1000 
- Words formula in B2: 
	- =UNIQUE(TOCOL(IFERROR(REDUCE(,INDIRECT($D$2),LAMBDA(x,y,VSTACK(TEXTSPLIT(x, " "),TEXTSPLIT(y, " ")))),"-"))) 
- Frequency formula in C2: 
    - =SUM((LEN(INDIRECT($D$2))-LEN(SUBSTITUTE(INDIRECT($D$2),B2,"")))/LEN(B2)) 
- Copy the C cells down 
- Credit for technique to: [https://www.get-digital-help.com/excel-udf-word-frequency/](https://www.get-digital-help.com/excel-udf-word-frequency/)  
- Copy Paste Special Text Only into separate sheet and Sort 

This is a ‘clever’ technique, but as with many Spreadsheet ‘coding via formula’ solutions, very inefficient and it begins to fail once you get over several thousand cells in your range.