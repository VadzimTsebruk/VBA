Attribute VB_Name = "OL_mdl_CopyConversationID"
Option Explicit

Public Sub Copy_conversation_to_clipboard_GER()

    Dim objMail     As Object   'was earlier Outlook.MailItem
    Dim doClipboard As MSForms.DataObject
    Dim message     As String
         
    'One and ONLY one message must be selected
    'Adjust the title to your language
    If Application.ActiveExplorer.Selection.Count <> 1 Then
        MsgBox _
            Title:="Warnung", _
            prompt:="Bitte waehlen Sie Eine und nur EINE Nachricht aus!", _
            buttons:=vbExclamation
        Exit Sub
    End If
    
    Set doClipboard = New MSForms.DataObject
    Set objMail = Application.ActiveExplorer.Selection.Item(1)
    
    'Adjust the word in brackets [Unterhaltung] to suit your language
    doClipboard.SetText "[Unterhaltung]:=""" & objMail.ConversationTopic & """"
    doClipboard.PutInClipboard
 
End Sub

Public Sub Copy_conversation_to_clipboard_ENG()

    Dim objMail     As Object   'was earlier Outlook.MailItem
    Dim doClipboard As MSForms.DataObject
    Dim message     As String
         
    'One and ONLY one message must be selected
    'Adjust the title to your language
    If Application.ActiveExplorer.Selection.Count <> 1 Then
        MsgBox _
            Title:="Warning", _
            prompt:="Please, select one and ONLY one message!", _
            buttons:=vbExclamation
        Exit Sub
    End If
    
    Set doClipboard = New MSForms.DataObject
    Set objMail = Application.ActiveExplorer.Selection.Item(1)
    
    'Adjust the word in brackets [Conversation] to suit your language
    doClipboard.SetText "[Conversation]:=""" & objMail.ConversationTopic & """"
    doClipboard.PutInClipboard
 
End Sub

