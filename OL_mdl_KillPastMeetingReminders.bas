Attribute VB_Name = "OL_mdl_KillPastMeetingReminders"
Option Explicit
'https://techniclee.wordpress.com/2011/04/01/killing-overdue-reminders-in-outlook/

Private Sub Application_Startup()
    KillOverdueReminders
End Sub
 
Sub KillOverdueReminders()
    ' Purpose: Kills all reminders for past due appointments.'
    ' Usage:   Run at Outlook startup to eliminate reminders on appointments that occur in the past.'
    ' Written: 4/1/2011'
    ' Modified: 11/30/2011'
    ' Author:  David Lee'
    ' Outlook: All versions'
    Const GRACE_PERIOD_MINUTES = 60
    
'    VT 04.01.2019 09:50: Falsche Variablendeklaration vom Author, daher kurzf. geloest durch Variant
'    als Variablentyp
'    Dim olkReminders    As Outlook.Reminder
'    Dim olkReminder     As Outlook.Reminder
'    Dim intCount        As Integer
'    Dim intIndex        As Integer

    Dim olkReminders    As Variant
    Dim olkReminder     As Variant
    Dim intCount        As Variant
    Dim intIndex        As Variant
    
    '*****************************************************************************************************
    '                                                       CODE
    '*****************************************************************************************************
    
    Set olkReminders = Application.Reminders
    
    intCount = olkReminders.Count
    
    For intIndex = intCount To 1 Step -1
        Set olkReminder = olkReminders.Item(intIndex)
        'Please left-click somewhere into your calendar to ensure that you didn't selected for example _
        'e-mails.
        'Otherwise the code will run into an error.
        If olkReminder.Item.Class = olAppointment Then
            If DateAdd("n", GRACE_PERIOD_MINUTES * -1, Now) > olkReminder.NextReminderDate Then
                olkReminders.Remove intIndex
            End If
        End If
    Next
    
    Set olkReminder = Nothing
    
End Sub
