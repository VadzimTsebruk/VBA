Attribute VB_Name = "OL_mdl_categorize_appt2learning"
Option Explicit
Public Sub categorize_appointment2learning()
    'https://stackoverflow.com/questions/43919275/outlook-application-filedialog-not-found
    'Open file dialog
'    Dim xlApp As Object
'    Set xlApp = CreateObject("Excel.Application")
'        xlApp.Visible = False
'
'    Dim fd As Office.FileDialog
'    Set fd = xlApp.Application.FileDialog(msoFileDialogFilePicker)
'
'    Dim selectedItem As Variant
'
'    If fd.Show = -1 Then
'        For Each selectedItem In fd.SelectedItems
'            Debug.Print selectedItem
'        Next
'    End If
'
'    Set fd = Nothing
'        xlApp.Quit
'    Set xlApp = Nothing

    'English version
    SendKeys "^{F4}" 'CTRL + F4 | Categorize appointment w/ category 4 which, in my case, is learning.
    SendKeys "%hb1f{ENTER}" 'ALT + H + B1 + F + Enter | Appointment busy status free.
    SendKeys "%{F4}" 'ALT + F4 | Close current window (if FN lock is enabled).
    SendKeys "y" 'Confirm 'Do you want to save your changes' dialog w/ "yes".
End Sub



