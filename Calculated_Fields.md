
These are formulas that can be used in Custom / Calculated fields
in Microsoft Office programmes

## Project

### This Week's Activity

This is to add a Custom Field using a Formula to 
show whether the Activity is due to be in progress This Week. 
This can be valuable for filtering views for reporting. 
It has been validated with MS Project 2010. 

Requirement:

* The Task is assumed to be Active This Week if:
    * It Starts on or before Saturday at the end of the week
    * AND
    * It Ends on or after the Sunday at the start of the week
* It is designed for reporting in advance
    * If run up until Wednesday it reports This Actual Week
    * If run from Thursday onwards it reports as if already in the following week
    * you can change this by changing the  > 4
        * e.g.  > 5  report as if next week from Friday onwards
    * we can use Status Date to allow flexibility on which week to focus on 
* This is designed to be easily adapted for Last Week and Next Week 
    * simply replace the + 0 with 
        *   + 7 for Next Week
        *   - 7 for Last Week

Implementation:

* Set up a new Custom Field
    * MS Project Menu / Project / Properties / Custom Fields
    * Field: Task
    * Type: Flag
    * (pick an empty flag, e.g. Flag1)
    * Rename: Flag1_ActiveThisWeek
    * Custom Attributes: Formula
        * paste in the formula below

Formula:

IIf(
  [Start] <= (7 + 0 + [Status Date] - Weekday([Status Date]) 
  + IIf(Weekday([Status Date]) > 4 ;7;0) )
AND
  [Finish] >= (1 + 0 + [Status Date] - Weekday([Status Date]) 
  + IIf(Weekday([Status Date]) > 4 ;7;0) )
;Yes;No)


Usage: 

* Display the field in a table
    * Open an Activity / Task view, such as Gannt Chart
    * in the top right of the table, click Add New Column
    * Choose Flag1_ActiveThisWeek
    * Right click on new Column Heading / Field Settings
        * Title: This Week

* If the field does not fill correctly as you might expect
    * check the Status Date as below

* Filter using the field
    * Open (or Copy) the view you want
    * MS Project Menu / View / Data / Filter / DropDown / More Filters / Edit
    * Field Name: Flag1_ActiveThisWeek
    * Test: Equals
    * Values: Yes
	* Show related summary rows: check


* We use the Project Status Date to determine which Week we’re focused on
    * It allows us to be more flexible than Current Date
    * This will normally default to be today's date
    * This can be edited in the interface by the Project Manager
        * MS Project Menu / Project / Status / Status Date


#### Finishing last week

Following similar requirements to the Flag above,
here is the formula for a separate flag to indicate all tasks 
that were planned to finish during the previous week.
NB: This does NOT take % Complete into account

Flag2_FinishingLastWeek

IIf(
  [Finish] <= (7 - 7 + [Status Date] - Weekday([Status Date]) 
  + IIf(Weekday([Status Date]) > 4 ;7;0) )
AND
  [Finish] >= (1 - 7 + [Status Date] - Weekday([Status Date]) 
  + IIf(Weekday([Status Date]) > 4 ;7;0) )
;Yes;No)


##### Happing this week OR Finished last week

If you want to extend the "This Week's Activity" calculation 
to include tasks that Finished last week AS WELL, 

IIf(
(
  [Start] <= (7 + 0 + [Status Date] - Weekday([Status Date]) 
  + IIf(Weekday([Status Date]) > 4 ;7;0) )
AND
  [Finish] >= (1 + 0 + [Status Date] - Weekday([Status Date]) 
  + IIf(Weekday([Status Date]) > 4 ;7;0) )
) OR (
  [Finish] <= (7 - 7 + [Status Date] - Weekday([Status Date]) 
  + IIf(Weekday([Status Date]) > 4 ;7;0) )
AND
  [Finish] >= (1 - 7 + [Status Date] - Weekday([Status Date]) 
  + IIf(Weekday([Status Date]) > 4 ;7;0) )
)
;Yes;No)

NB: there may be a shorter formula available by twisting the logic, 
but this one might be simpler to understand and build on


