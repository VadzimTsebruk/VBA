VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OL_class_ChangeTextBodyType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'https://www.datanumen.com/blogs/how-to-set-the-default-font-for-the-body-of-contact-appointment-and-task-via-vba/

Public WithEvents objInspectors As Outlook.Inspectors
Attribute objInspectors.VB_VarHelpID = -1
Public WithEvents objInspector As Outlook.Inspector
Attribute objInspector.VB_VarHelpID = -1

Public Sub Application_Startup()
    Set objInspectors = Outlook.Application.Inspectors
End Sub

Public Sub objInspectors_NewInspector(ByVal Inspector As Outlook.Inspector)
    Set objInspector = Inspector
End Sub

Public Sub objInspector_Activate()
    Dim objCurrentItem As Object
    Dim objCurrentInspector As Outlook.Inspector
    Dim objWordDocument As Word.Document
    Dim objWordSelection As Word.Selection
 
    Set objCurrentItem = objInspector.CurrentItem
    Set objCurrentInspector = objCurrentItem.GetInspector
    Set objWordDocument = objCurrentInspector.WordEditor
    Set objWordSelection = objWordDocument.Application.Selection
 
    'You change the font as per your preference
    Select Case objCurrentItem.Class
           Case olContact
                With objWordSelection.Font
                     .Name = "Arial"
                     .ColorIndex = wdAuto
                     .Size = 10
                     .Bold = True
                End With
           Case olAppointment
                With objWordSelection.Font
                     .Name = "Arial"
                     .ColorIndex = wdAuto
                     .Size = 10
                     .Bold = False
                End With
           Case olTask
                With objWordSelection.Font
                     .Name = "Arial"
                     .ColorIndex = wdAuto
                     .Size = 20
                     .Bold = True
                End With
   End Select
End Sub

