Attribute VB_Name = "OL_mdl_DeleteMultipleAttachment"
Option Explicit
'https://recoverit.wondershare.com/email-tips/remove-attachments-from-outlook.html#:~:text=This%20method%20helps%20to%20delete,would%20be%20deleted%20at%20once.


Public Sub DeleteAttachments2()

Dim coll As VBA.Collection
Dim obj As Object
Dim Atts As Outlook.Attachments
Dim Att As Outlook.Attachment
Dim Sel As Outlook.Selection
Dim i&, Msg$

Set coll = New VBA.Collection

'On Error Resume Next
If TypeOf Application.ActiveWindow Is Outlook.Inspector Then
    coll.Add Application.ActiveInspector.CurrentItem
    Else
        Set Sel = Application.ActiveExplorer.Selection
    For i = 1 To Sel.Count
        coll.Add Sel(i)
    Next
End If

On Error Resume Next
For Each obj In coll
    Set Atts = obj.Attachments
    Msg = ""
    For i = Atts.Count To 1 Step -1
        Msg = Msg & Atts(i).FileName & vbCrLf
    Next
    'Delete confirmation by user required. Otherwise, end sub.
    If MsgBox(Msg, vbYesNo Or vbQuestion, "Delete?") = vbYes Then
        For i = Atts.Count To 1 Step -1
            Atts.Remove i
        Next
            obj.Save
    End If
Next

End Sub
