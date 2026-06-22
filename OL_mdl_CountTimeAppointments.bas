Attribute VB_Name = "OL_mdl_CountTimeAppointments"
Option Explicit
'https://www.extendoffice.com/documents/outlook/3551-outlook-calendar-count-hours-days-weeks.html _
Futher reading for future in order to enhance this macro: _
https://www.howto-outlook.com/howto/timespent.htm#manually

Sub CountTimeSpent()

    Dim oOLApp As Outlook.Application
    Dim oSelection As Outlook.Selection
    Dim oItem As Object
    Dim iDuration As Long
    Dim iTotalWork As Long
    Dim iMileage As Long
    Dim iResult As Integer
    Dim bShowiMileage As Boolean
    Dim oItems As Outlook.Items
    
    bShowiMileage = False
     
    iDuration = 0
    iTotalWork = 0
    iMileage = 0
     
    On Error Resume Next
     
    'Include all recurring calendar items –
    oItems.IncludeRecurrences = True

    Set oOLApp = CreateObject("Outlook.Application")
    Set oSelection = oOLApp.ActiveExplorer.Selection
     
    For Each oItem In oSelection
    If oItem.Class = olAppointment Then
    iDuration = iDuration + oItem.Duration
    iMileage = iMileage + oItem.Mileage
    ElseIf oItem.Class = olTask Then
    iDuration = iDuration + oItem.ActualWork
    iTotalWork = iTotalWork + oItem.TotalWork
    iMileage = iMileage + oItem.Mileage
    ElseIf oItem.Class = Outlook.olJournal Then
    iDuration = iDuration + oItem.Duration
    iMileage = iMileage + oItem.Mileage
    Else
    iResult = MsgBox("Please select some Calendar, Task or Journal items at first!", vbCritical, "Items Time Spent")
    Exit Sub
    End If
    Next
     
    Dim MsgBoxText As String
    MsgBoxText = "Total time spent: " & vbNewLine & iDuration & " minutes"
     
    If iDuration > 60 Then
    MsgBoxText = MsgBoxText & HoursMsg(iDuration)
    End If
     
    If iTotalWork > 0 Then
    MsgBoxText = MsgBoxText & vbNewLine & vbNewLine & "Total work recorded; " & vbNewLine & iTotalWork & " minutes"
     
    If iTotalWork > 60 Then
    MsgBoxText = MsgBoxText & HoursMsg(iTotalWork)
    End If
    End If
     
    If bShowiMileage = True Then
    MsgBoxText = MsgBoxText & vbNewLine & vbNewLine & "Total iMileage; " & iMileage
    End If
     
        iResult = MsgBox(MsgBoxText, vbInformation, "Items Time spent")
     
ExitSub:
    Set oItem = Nothing
    Set oSelection = Nothing
    Set oOLApp = Nothing
    
End Sub
 
Function HoursMsg(TotalMinutes As Long) As String

    Dim iHours As Long
    Dim iMinutes As Long
    iHours = TotalMinutes \ 60
    iMinutes = TotalMinutes Mod 60
    HoursMsg = " (" & iHours & " Hours and " & iMinutes & " Minutes)"

End Function
