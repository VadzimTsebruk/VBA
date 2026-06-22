Attribute VB_Name = "OL_mdl_OpenTemplates"
'Option Explicit
'https://www.slipstick.com/outlook/hyperlink-templates/
Dim template As String

Sub OpenKanbanTemplate()
template = Environ("USERPROFILE") & "\AppData\Roaming\Microsoft\Templates\KANBAN.oft"
MakeItem
End Sub

'Sub OpenTemplate2()
'template = Environ("USERPROFILE") & "\Templates\email.oft"
'MakeItem
'End Sub
'
Private Sub MakeItem()
Set newItem = Application.CreateItemFromTemplate(template)
newItem.Display
Set newItem = Nothing
End Sub
