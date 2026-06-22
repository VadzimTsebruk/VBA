Attribute VB_Name = "OL_mdl_OpenPersonalForms"
Option Explicit
'https://www.slipstick.com/outlook/hyperlink-templates/
'https://stackoverflow.com/questions/31483869/determining-selected-outlook-calendar-date-with-vba
 
Sub OpenWorkFreeTemplateUsingSelectedTime()
    
    Dim datStart    As Date
    Dim datEnd      As Date
    Dim oView       As Outlook.View
    Dim oCalView    As Outlook.CalendarView
    Dim oExpl       As Outlook.Explorer
    Dim oFolder     As Outlook.Folder
    Dim oAppt       As Outlook.AppointmentItem
    Const datNull   As Date = #1/1/4501#
    
    ' Obtain the calendar view using
    ' Application.ActiveExplorer.CurrentFolder.CurrentView.
    ' If you use oExpl.CurrentFolder.CurrentView,
    ' this code will not operate as expected.
    Set oExpl = Application.ActiveExplorer
    Set oFolder = Application.ActiveExplorer.CurrentFolder
    Set oView = oExpl.CurrentView
    
    ' Check whether the active explorer is displaying a calendar view.
    If oView.ViewType = olCalendarView Then
        Set oCalView = oExpl.CurrentView
        ' Create the appointment using the values in
        ' the SelectedStartTime and SelectedEndTime properties as
        ' appointment start and end times.
        datStart = oCalView.SelectedStartTime
        datEnd = oCalView.SelectedEndTime
        'Default = "IPM.Appointment" _
        'If you want to use a personal form aka template adjust by choosing your desired folder from _
        C:\Users\XXXX\AppData\Local\Microsoft\FORMS
        Set oAppt = oFolder.Items.Add("IPM.Appointment.WorkFree")
        If datStart <> datNull And datEnd <> datNull Then
            oAppt.Start = datStart
            oAppt.End = datEnd
            oAppt.ReminderSet = True
        End If
        oAppt.Display
    End If
    
End Sub

Sub OpenWorkTentativeTemplateUsingSelectedTime()
    
    Dim datStart    As Date
    Dim datEnd      As Date
    Dim oView       As Outlook.View
    Dim oCalView    As Outlook.CalendarView
    Dim oExpl       As Outlook.Explorer
    Dim oFolder     As Outlook.Folder
    Dim oAppt       As Outlook.AppointmentItem
    Const datNull   As Date = #1/1/4501#
    
    ' Obtain the calendar view using
    ' Application.ActiveExplorer.CurrentFolder.CurrentView.
    ' If you use oExpl.CurrentFolder.CurrentView,
    ' this code will not operate as expected.
    Set oExpl = Application.ActiveExplorer
    Set oFolder = Application.ActiveExplorer.CurrentFolder
    Set oView = oExpl.CurrentView
    
    ' Check whether the active explorer is displaying a calendar view.
    If oView.ViewType = olCalendarView Then
        Set oCalView = oExpl.CurrentView
        ' Create the appointment using the values in
        ' the SelectedStartTime and SelectedEndTime properties as
        ' appointment start and end times.
        datStart = oCalView.SelectedStartTime
        datEnd = oCalView.SelectedEndTime
        'Default = "IPM.Appointment" _
        'If you want to use a personal form aka template adjust by choosing your desired folder from _
        C:\Users\XXXX\AppData\Local\Microsoft\FORMS
        Set oAppt = oFolder.Items.Add("IPM.Appointment.WorkTentative")
        If datStart <> datNull And datEnd <> datNull Then
            oAppt.Start = datStart
            oAppt.End = datEnd
            oAppt.ReminderSet = False
            oAppt.Sensitivity = olNormal
            oAppt.BusyStatus = olTentative
        End If
        oAppt.Display
    End If
    
End Sub

Sub OpenPrivatNextPriorityOOOTemplateUsingSelectedTime()
    
    Dim datStart    As Date
    Dim datEnd      As Date
    Dim oView       As Outlook.View
    Dim oCalView    As Outlook.CalendarView
    Dim oExpl       As Outlook.Explorer
    Dim oFolder     As Outlook.Folder
    Dim oAppt       As Outlook.AppointmentItem
    Const datNull   As Date = #1/1/4501#
    
    ' Obtain the calendar view using
    ' Application.ActiveExplorer.CurrentFolder.CurrentView.
    ' If you use oExpl.CurrentFolder.CurrentView,
    ' this code will not operate as expected.
    Set oExpl = Application.ActiveExplorer
    Set oFolder = Application.ActiveExplorer.CurrentFolder
    Set oView = oExpl.CurrentView
       
    ' Check whether the active explorer is displaying a calendar view.
    If oView.ViewType = olCalendarView Then
        Set oCalView = oExpl.CurrentView
        ' Create the appointment using the values in
        ' the SelectedStartTime and SelectedEndTime properties as
        ' appointment start and end times.
        datStart = oCalView.SelectedStartTime
        datEnd = oCalView.SelectedEndTime
        'Default = "IPM.Appointment" _
        'If you want to use a personal form aka template adjust by choosing your desired folder from _
        C:\Users\XXXX\AppData\Local\Microsoft\FORMS
        Set oAppt = oFolder.Items.Add("IPM.Appointment.PrivatNextPriorityOOO")
        If datStart <> datNull And datEnd <> datNull Then
            oAppt.Start = datStart
            oAppt.End = datEnd
            oAppt.ReminderSet = True
        End If
        oAppt.Display
    End If
    
End Sub

Sub OpenPrivateFreeTemplateUsingSelectedTime()
    
    Dim datStart    As Date
    Dim datEnd      As Date
    Dim oView       As Outlook.View
    Dim oCalView    As Outlook.CalendarView
    Dim oExpl       As Outlook.Explorer
    Dim oFolder     As Outlook.Folder
    Dim oAppt       As Outlook.AppointmentItem
    Const datNull   As Date = #1/1/4501#
    
    ' Obtain the calendar view using
    ' Application.ActiveExplorer.CurrentFolder.CurrentView.
    ' If you use oExpl.CurrentFolder.CurrentView,
    ' this code will not operate as expected.
    Set oExpl = Application.ActiveExplorer
    Set oFolder = Application.ActiveExplorer.CurrentFolder
    Set oView = oExpl.CurrentView
    
    ' Check whether the active explorer is displaying a calendar view.
    If oView.ViewType = olCalendarView Then
        Set oCalView = oExpl.CurrentView
        ' Create the appointment using the values in
        ' the SelectedStartTime and SelectedEndTime properties as
        ' appointment start and end times.
        datStart = oCalView.SelectedStartTime
        datEnd = oCalView.SelectedEndTime
        'Default = "IPM.Appointment" _
        'If you want to use a personal form aka template adjust by choosing your desired folder from _
        C:\Users\XXXX\AppData\Local\Microsoft\FORMS
        Set oAppt = oFolder.Items.Add("IPM.Appointment.PrivatFree")
        If datStart <> datNull And datEnd <> datNull Then
            oAppt.Start = datStart
            oAppt.End = datEnd
            oAppt.ReminderSet = False
        End If
        oAppt.Display
    End If
End Sub


Sub OpenPrivatOoOTemplateUsingSelectedTime()
    
    Dim datStart    As Date
    Dim datEnd      As Date
    Dim oView       As Outlook.View
    Dim oCalView    As Outlook.CalendarView
    Dim oExpl       As Outlook.Explorer
    Dim oFolder     As Outlook.Folder
    Dim oAppt       As Outlook.AppointmentItem
    Dim Category    As String
    Const datNull   As Date = #1/1/4501#
        
    ' Obtain the calendar view using
    ' Application.ActiveExplorer.CurrentFolder.CurrentView.
    ' If you use oExpl.CurrentFolder.CurrentView,
    ' this code will not operate as expected.
    Set oExpl = Application.ActiveExplorer
    Set oFolder = Application.ActiveExplorer.CurrentFolder
    Set oView = oExpl.CurrentView
    
    'Assign category
    Category = "05_Private"
    
    ' Check whether the active explorer is displaying a calendar view.
    If oView.ViewType = olCalendarView Then
        Set oCalView = oExpl.CurrentView
        ' Create the appointment using the values in
        ' the SelectedStartTime and SelectedEndTime properties as
        ' appointment start and end times.
        datStart = oCalView.SelectedStartTime
        datEnd = oCalView.SelectedEndTime
        'Default = "IPM.Appointment" _
        'If you want to use a personal form aka template adjust by choosing your desired folder from _
        C:\Users\XXXX\AppData\Local\Microsoft\FORMS
        Set oAppt = oFolder.Items.Add("IPM.Appointment.PrivatOoO")
            oAppt.BusyStatus = olOutOfOffice
        If datStart <> datNull And datEnd <> datNull Then
            oAppt.Start = datStart
            oAppt.End = datEnd
            oAppt.ReminderSet = True
            oAppt.Categories = Category
            oAppt.Sensitivity = olNormal
        End If
        oAppt.Display
    End If
    
End Sub

Sub OpenAllDayEventFreeTemplateUsingSelectedTime()
    
    Dim datStart    As Date
    Dim datEnd      As Date
    Dim oView       As Outlook.View
    Dim oCalView    As Outlook.CalendarView
    Dim oExpl       As Outlook.Explorer
    Dim oFolder     As Outlook.Folder
    Dim oAppt       As Outlook.AppointmentItem
    Const datNull   As Date = #1/1/4501#
    
    ' Obtain the calendar view using
    ' Application.ActiveExplorer.CurrentFolder.CurrentView.
    ' If you use oExpl.CurrentFolder.CurrentView,
    ' this code will not operate as expected.
    Set oExpl = Application.ActiveExplorer
    Set oFolder = Application.ActiveExplorer.CurrentFolder
    Set oView = oExpl.CurrentView
    
    ' Check whether the active explorer is displaying a calendar view.
    If oView.ViewType = olCalendarView Then
        Set oCalView = oExpl.CurrentView
        ' Create the appointment using the values in
        ' the SelectedStartTime and SelectedEndTime properties as
        ' appointment start and end times.
        datStart = oCalView.SelectedStartTime
        datEnd = oCalView.SelectedEndTime
        'Default = "IPM.Appointment" _
        'If you want to use a personal form aka template adjust by choosing your desired folder from _
        C:\Users\XXXX\AppData\Local\Microsoft\FORMS
        Set oAppt = oFolder.Items.Add("IPM.Appointment.AllDayEventFree")
        If datStart <> datNull And datEnd <> datNull Then
            oAppt.Start = datStart
            oAppt.End = datEnd
            oAppt.ReminderSet = True
        End If
        oAppt.Display
    End If
    
End Sub

Sub OpenAllDayPrivatImportantTemplateUsingSelectedTime()
    
    Dim datStart    As Date
    Dim datEnd      As Date
    Dim oView       As Outlook.View
    Dim oCalView    As Outlook.CalendarView
    Dim oExpl       As Outlook.Explorer
    Dim oFolder     As Outlook.Folder
    Dim oAppt       As Outlook.AppointmentItem
    Const datNull   As Date = #1/1/4501#
    
    ' Obtain the calendar view using
    ' Application.ActiveExplorer.CurrentFolder.CurrentView.
    ' If you use oExpl.CurrentFolder.CurrentView,
    ' this code will not operate as expected.
    Set oExpl = Application.ActiveExplorer
    Set oFolder = Application.ActiveExplorer.CurrentFolder
    Set oView = oExpl.CurrentView
    
    ' Check whether the active explorer is displaying a calendar view.
    If oView.ViewType = olCalendarView Then
        Set oCalView = oExpl.CurrentView
        ' Create the appointment using the values in
        ' the SelectedStartTime and SelectedEndTime properties as
        ' appointment start and end times.
        datStart = oCalView.SelectedStartTime
        datEnd = oCalView.SelectedEndTime
        'Default = "IPM.Appointment" _
        'If you want to use a personal form aka template adjust by choosing your desired folder from _
        C:\Users\XXXX\AppData\Local\Microsoft\FORMS
        Set oAppt = oFolder.Items.Add("IPM.Appointment.PRIVAT_FREE_IMPORTANT")
        If datStart <> datNull And datEnd <> datNull Then
            oAppt.Start = datStart
            oAppt.End = datEnd
            oAppt.ReminderSet = True
        End If
        oAppt.Display
    End If
    
End Sub

Sub Open_Learning_Private_Tentative_Template_Using_Selected_Time()
    
    Dim datStart    As Date
    Dim datEnd      As Date
    Dim oView       As Outlook.View
    Dim oCalView    As Outlook.CalendarView
    Dim oExpl       As Outlook.Explorer
    Dim oFolder     As Outlook.Folder
    Dim oAppt       As Outlook.AppointmentItem
    Const datNull   As Date = #1/1/4501#
    
    ' Obtain the calendar view using
    ' Application.ActiveExplorer.CurrentFolder.CurrentView.
    ' If you use oExpl.CurrentFolder.CurrentView,
    ' this code will not operate as expected.
    Set oExpl = Application.ActiveExplorer
    Set oFolder = Application.ActiveExplorer.CurrentFolder
    Set oView = oExpl.CurrentView
    
    ' Check whether the active explorer is displaying a calendar view.
    If oView.ViewType = olCalendarView Then
        Set oCalView = oExpl.CurrentView
        ' Create the appointment using the values in
        ' the SelectedStartTime and SelectedEndTime properties as
        ' appointment start and end times.
        datStart = oCalView.SelectedStartTime
        datEnd = oCalView.SelectedEndTime
        'Default = "IPM.Appointment" _
        'If you want to use a personal form aka template adjust by choosing your desired folder from _
        C:\Users\XXXX\AppData\Local\Microsoft\FORMS
        Set oAppt = oFolder.Items.Add("IPM.Appointment.LEARNING_TENTATIVE")
        If datStart <> datNull And datEnd <> datNull Then
            oAppt.Start = datStart
            oAppt.End = datEnd
            oAppt.ReminderSet = False
        End If
        oAppt.Display
    End If
    
End Sub

