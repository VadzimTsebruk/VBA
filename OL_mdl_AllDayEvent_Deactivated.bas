Attribute VB_Name = "OL_mdl_AllDayEvent_Deactivated"
Option Explicit
'Source: https://superuser.com/questions/48380/outlook-and-all-day-event-reminders

Sub AllDaySetToZeroFor6months()

    Dim daStart             As Date
    Dim daEnd               As Date
    Dim oCalendar           As Outlook.Folder
    Dim oItems              As Outlook.Items
    Dim oItemsInDateRange   As Outlook.Items
    Dim oFinalItems         As Outlook.Items
    Dim oAppt               As Outlook.AppointmentItem
    Dim strRestriction      As String
    Dim Debuglog            As Variant 'Don't know which datatype suits
    Dim CurrentTitle        As String
    
    '*****************************************************************************************************
    '                                                       CODE
    '*****************************************************************************************************
    
    '1
    'Set the date range for the appointments query -
    'It is set below to start at todays date and
    'end at todays date + 120 days (or 4 months)
    'You can increase or reduce this based on your PCs performance
    
        daStart = Format(Date, "mm/dd/yyyy hh:mm AMPM")
        daEnd = DateAdd("d", 120, daStart)
        daEnd = Format(daEnd, "mm/dd/yyyy hh:mm AMPM")
        Debuglog = "1 Start: " & daStart
        Debuglog = Debuglog & ", " & "1 End: " & daEnd
    
    '2
    'Construct a filter for the next 120-day date range.
        strRestriction = "[Start] >= ‘" & daStart _
        & " 'AND [End] <= ‘" & daEnd & "‘"
        Debuglog = Debuglog & ", " & "2 " & strRestriction
    
    '3
    'The macro obtains the set of appointment items in the default calendar
    'specified by the current Outlook user profile.
    
        Set oCalendar = Application.Session.GetDefaultFolder(olFolderCalendar)
        Set oItems = oCalendar.Items
    
    '4
    'To include recurring appointments, sort by using the Start property.
        oItems.IncludeRecurrences = True
        oItems.Sort "[Start]"
    
    '5
    'Restrict the Items collection for the 1110-day date range.
        Set oFinalItems = oItems.Restrict(strRestriction)
    
    '6
    'Go through each calendar item remaining in turn
    'If it isn’t a full Day event do nothing
    'If it is set Reminder to 0 Minutes.
        oFinalItems.Sort "[Start]"""
        For Each oAppt In oFinalItems
        Debuglog = Debuglog & ", " & "6 " & oAppt.Start & ", " & oAppt.Subject & ", " & oAppt.ReminderMinutesBeforeStart
        CurrentTitle = oAppt.Subject
        If oAppt.AllDayEvent = False Then
        Else
        oAppt.ReminderMinutesBeforeStart = 0
        oAppt.Save
        End If
        Debuglog = Debuglog & ", " & "6 " & oAppt.Start & ", " & oAppt.Subject & ", " & oAppt.ReminderMinutesBeforeStart & vbNewLine & vbNewLine
        Next
        Debuglog = ""

End Sub
