Attribute VB_Name = "OL_mdl_Forbit_ReplyandForward"
Option Explicit

Sub OL_Forbit_Reply_ReplytoAll_and_Forward_Messages()

    ActiveInspector.CurrentItem.Actions("Antworten").Enabled = False
    ActiveInspector.CurrentItem.Actions("Allen antworten").Enabled = False
    ActiveInspector.CurrentItem.Actions("Weiterleiten").Enabled = False

End Sub
