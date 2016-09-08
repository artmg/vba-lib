Attribute VB_Name = "mod_sps_ItemValueUpdate"


' UpdateListItems in SP 2013 https://msdn.microsoft.com/en-us/library/office/websvclists.lists.updatelistitems.aspx
' example of how to turn the object model code into pure SOAP request string http://stackoverflow.com/q/11092273
' simple example of creating the soap object with late binding https://community.spiceworks.com/topic/478128-soap-request-via-vba-in-excel
' other examples:
' http://stackoverflow.com/questions/22450717/
' http://sharepoint.stackexchange.com/questions/137934/
' http://sharepoint.stackexchange.com/questions/93181/
' how to build SOAP requests to retrieve data http://depressedpress.com/2014/04/05/accessing-sharepoint-lists-with-visual-basic-for-applications/

' introduction to the ShrePoint 2013 REST interface https://msdn.microsoft.com/en-us/library/office/fp142380.aspx
' basics of REST for retrieving and updating SharePoint data https://msdn.microsoft.com/en-us/magazine/dn198245.aspx
' full details of using REST for all List column types, but using ajax code http://www.codeproject.com/Articles/990131/CRUD-operation-to-list-using-SharePoint-Rest-API
' ?? https://msdn.microsoft.com/en-us/library/office/dn567558.aspx



' ## Intro
' After list comparison and analysis, you may end up with a table of values to alter. If the data is in a SharePoint list you want a simple way to make these updates. Because items have last modified and possibly version history, you want to limit the changes to strict necessity
'
' ## Purpose
' Table of items from SharePoint list with ID
' Where a column contains a value, set this on the list item
' Batch up changes to a single line for efficiency and better readability is version history
'
' ## Solutions candidates
' a)  Excel office VBA via web services (SOAP simpler, is REST better supported?)
' b)  PS1 over CSV, but needs object model locally, more suited to admin use
'
' ## Psuedo -code
'For each row (from second to last)
'Assume Col A is ID
'Prepare Request String
'For each column cell - if value
'    Add value update string
'Add terminal string
'    (Optionally) IF there was an update
'        Send the request to server
