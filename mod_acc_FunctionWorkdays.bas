Attribute VB_Name = "mod_acc_FunctionWorkdays"
Option Compare Database

' credit > http://msdn.microsoft.com/en-us/library/office/dd327646(v=office.12).aspx

Public Function Workdays(ByRef StartDate As Date, _
     ByRef endDate As Date, _
     Optional ByRef strHolidays As String = "Holidays" _
     ) As Integer
    ' Returns the number of workdays between startDate
    ' and endDate inclusive.  Workdays excludes weekends and
    ' holidays. Optionally, pass this function the name of a table
    ' or query as the third argument. If you don't the default
    ' is "Holidays".
    On Error GoTo Workdays_Error
    Dim nWeekdays As Integer
    Dim nHolidays As Integer
    Dim strWhere As String
    
    ' DateValue returns the date part only.
    StartDate = DateValue(StartDate)
    endDate = DateValue(endDate)
    
    nWeekdays = Weekdays(StartDate, endDate)
    If nWeekdays = -1 Then
        Workdays = -1
        GoTo Workdays_Exit
    End If
    
    strWhere = "[Holiday] >= #" & StartDate _
        & "# AND [Holiday] <= #" & endDate & "#"
    
    ' Count the number of holidays.
    nHolidays = DCount(Expr:="[Holiday]", _
        Domain:=strHolidays, _
        Criteria:=strWhere)
    
    Workdays = nWeekdays - nHolidays
    
Workdays_Exit:
    Exit Function
    
Workdays_Error:
    Workdays = -1
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, "Workdays"
    Resume Workdays_Exit
    
End Function

Public Function Weekdays(ByRef StartDate As Date, _
    ByRef endDate As Date _
    ) As Integer
    ' Returns the number of weekdays in the period from startDate
    ' to endDate inclusive. Returns -1 if an error occurs.
    ' If your weekend days do not include Saturday and Sunday and
    ' do not total two per week in number, this function will
    ' require modification.
    On Error GoTo Weekdays_Error
    
    ' The number of weekend days per week.
    Const ncNumberOfWeekendDays As Integer = 2
    
    ' The number of days inclusive.
    Dim varDays As Variant
    
    ' The number of weekend days.
    Dim varWeekendDays As Variant
    
    ' Temporary storage for datetime.
    Dim dtmX As Date
    
    ' If the end date is earlier, swap the dates.
    If endDate < StartDate Then
        dtmX = StartDate
        StartDate = endDate
        endDate = dtmX
    End If
    
    ' Calculate the number of days inclusive (+ 1 is to add back startDate).
    varDays = DateDiff(Interval:="d", _
        date1:=StartDate, _
        date2:=endDate) + 1
    
    ' Calculate the number of weekend days.
    varWeekendDays = (DateDiff(Interval:="ww", _
        date1:=StartDate, _
        date2:=endDate) _
        * ncNumberOfWeekendDays) _
        + IIf(DatePart(Interval:="w", _
        Date:=StartDate) = vbSunday, 1, 0) _
        + IIf(DatePart(Interval:="w", _
        Date:=endDate) = vbSaturday, 1, 0)
    
    ' Calculate the number of weekdays.
    Weekdays = (varDays - varWeekendDays)
    
Weekdays_Exit:
    Exit Function
    
Weekdays_Error:
    Weekdays = -1
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, "Weekdays"
    Resume Weekdays_Exit
End Function



