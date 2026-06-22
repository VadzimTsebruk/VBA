Attribute VB_Name = "OL_mdl_RemoveCopy"
Option Explicit

Sub OL_Entferne_Kopie_Termine()

    Dim myolApp As Outlook.Application
    Dim calendar As MAPIFolder
    Dim aItem As Object
    Dim iItemsUpdated As Integer
    Dim strTemp As String
    
    Set myolApp = CreateObject("Outlook.Application")
    Set calendar = myolApp.ActiveExplorer.CurrentFolder
    
    iItemsUpdated = 0
    For Each aItem In calendar.Items
        If Mid(aItem.Subject, 1, 7) = "Kopie: " _
        Or Mid(aItem.Subject, 1, 7) = "Kopieren: " Then
            strTemp = Mid(aItem.Subject, 8, Len(aItem.Subject) - 7)
            aItem.Subject = strTemp
            iItemsUpdated = iItemsUpdated + 1
        End If
        aItem.Save
    Next aItem
    
    MsgBox iItemsUpdated & " of " & calendar.Items.Count & " Meetings Updated"

End Sub
