* Subject: Access database VBA coding
* Title: Unbound Checkbox Array
* Abstract: This paper contains notes regarding the issue “cannot create an unbound checkbox on continuous forms”. This is a specific technical issue when programming data forms in Microsoft Access 2000, and the paper proposes a solution.
* Date: Last revised Tuesday, 14 March 2006

## Unbound checkbox array
### Why
Non-Access programmers may wonder why we had to do this. Unfortunately there is no simple way to have an unbound Checkbox in a Continuous Form. Try it and when you check it on one record, it gets checked on them all. The technical explanation for this deficiency is “Access creates a single instance of each object in a form” or “has just one set of control properties for all records displayed.

Although the most common solution is “create a temporary table”, we have chosen a slightly more intricate solution. It is more complex to understand - however because it is “in-memory” rather than based upon tables, it has the advantage of being much simpler to instantiate and cleanup in the running application.

### A Function supplies the data from an Array
We “bind” the checkbox to a function, rather than a data source. The function returns the value of the checkbox that is appropriate for any record being displayed, and that value is set and cleared by the checkbox’s update event. We “store” the data in a dynamic array by using the Dictionary object in the “Scripting” Runtime library. This uses key & data pairs, so is more flexible to use than a Collection, because we can access array elements via the string “key” (like the Properties collection). Every record displayed in a form has a unique Bookmark value. We create a member in the Dictionary “array” (using the Bookmark as the key) for each record where the checkbox is set.

### Source references
We developed this technique by research in places like experts-exchange.com and mvps.org\access. People have used similar techniques for conditional formatting in Continous Forms. The sample code at http://www.geocities.com/ragoran_ee/ marked the watershed of us understanding the technique properly - maybe it might help you, too.

