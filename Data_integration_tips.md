
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


