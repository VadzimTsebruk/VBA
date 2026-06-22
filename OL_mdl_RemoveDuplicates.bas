Attribute VB_Name = "OL_mdl_RemoveDuplicates"
Option Explicit
'https://www.datanumen.com/blogs/quickly-remove-duplicate-outlook-items-folder-via-vba/

Sub RemoveDuplicateItems()
    Dim objFolder As Folder
    Dim objDictionary As Object
    Dim i As Long
    Dim objItem As Object
    Dim strKey As String

    Set objDictionary = CreateObject("scripting.dictionary")
    'Select a source folder
    Set objFolder = Outlook.Application.Session.PickFolder

    If Not (objFolder Is Nothing) Then
       For i = objFolder.Items.Count To 1 Step -1
           Set objItem = objFolder.Items.Item(i)
 
           Select Case objFolder.DefaultItemType
                  'Check email subject, body and sent time
                  Case olMailItem
                       strKey = objItem.Subject & "," & objItem.Body & "," & objItem.SentOn
                  'Check appointment subject, start time, duration, location and body
                  Case olAppointmentItem
                       strKey = objItem.Subject & "," & objItem.Start & "," & objItem.Duration & "," & objItem.Location & "," & objItem.Body
                  'Check contact full name and email address
                  Case olContactItem
                       strKey = objItem.FullName & "," & objItem.Email1Address & "," & objItem.Email2Address & "," & objItem.Email3Address
                  'Check task subject, start date, due date and body
                  Case olTaskItem
                       strKey = objItem.Subject & "," & objItem.StartDate & "," & objItem.DueDate & "," & objItem.Body
           End Select
 
           strKey = Replace(strKey, ", ", Chr(32))
 
           'Remove the duplicate items
           If objDictionary.Exists(strKey) = True Then
              objItem.Delete
           Else
              objDictionary.Add strKey, True
           End If
       Next i
    End If
End Sub
