Attribute VB_Name = "OL_mdl_AddMeetingTime_Busy"
Option Explicit
'Source: http://www.vboffice.net/en/developers/add-appointment-travel-times-to-calendar

Public Sub AddMeetingTimeBusy()
  
    Dim coll As VBA.Collection
    Dim obj As Object
    Dim Appt As Outlook.AppointmentItem
    Dim Travel As Outlook.AppointmentItem
    Dim Items As Outlook.Items
    Dim Before&, After&
    Dim Category$, Subject$
    Dim Availability$
    
    'Block minutes before and after the appointment
    Before = 15
    After = 15
    
'    'Skip this if the default values never change
'    Before = InputBox("Minutes before:", , Before)
'    After = InputBox("Minutes after:", , After)

    If Before = 0 And After = 0 Then Exit Sub
    
    'Assign category
    Category = "04_Work"
    
    Set coll = GetCurrentItems
    If coll.Count = 0 Then Exit Sub
    For Each obj In coll
        If TypeOf obj Is Outlook.AppointmentItem Then
            Set Appt = obj
            If TypeOf Appt.Parent Is Outlook.AppointmentItem Then
              Set Items = Appt.Parent.Parent.Items
            Else
              Set Items = Appt.Parent.Items
            End If
        
            'Use the main appointment's subject
            Subject = Appt.Subject
            
            If Before > 0 Then
              Set Travel = Items.Add
              Travel.Subject = Subject
              Travel.Start = DateAdd("n", -Before, Appt.Start)
              Travel.Duration = Before
              Travel.Categories = Category
              Travel.ReminderSet = True
              Travel.BusyStatus = olBusy
'              Travel.Sensitivity = olPrivate
              Travel.Save
            End If
        
            If After > 0 Then
              Set Travel = Items.Add
              Travel.Subject = Subject
              Travel.Start = Appt.End
              Travel.Duration = After
              Travel.Categories = Category
              Travel.ReminderSet = False
              Travel.BusyStatus = olBusy
'              Travel.Sensitivity = olPrivate
              Travel.Save
            End If
      End If
    Next
  
End Sub

Private Function GetCurrentItems(Optional IsInspector As Boolean) As VBA.Collection
  
    Dim coll As VBA.Collection
    Dim Win As Object
    Dim Sel As Outlook.Selection
    Dim obj As Object
    Dim i&
    
    Set coll = New VBA.Collection
    Set Win = Application.ActiveWindow
    
    If TypeOf Win Is Outlook.Inspector Then
      IsInspector = True
      coll.Add Win.CurrentItem
    Else
      IsInspector = False
      Set Sel = Win.Selection
      If Not Sel Is Nothing Then
        For i = 1 To Sel.Count
          coll.Add Sel(i)
        Next
      End If
    End If
    Set GetCurrentItems = coll
  
End Function


