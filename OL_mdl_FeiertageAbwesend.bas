Attribute VB_Name = "OL_mdl_FeiertageAbwesend"
Option Explicit
Sub OL_Feiertage_im_Kalender_auf_Abwesend_setzen()

'*********************************************************************************************************
'                                               MACRO SUMMARY
'*********************************************************************************************************
'*!ATTENTION!*
'This macro will set *ALL* holidays to status "out of office" and this includes holidays from other
'countries too.
'Such a scenario is possible when the user imported other hoidays into his/her Outlook
'calendar for reasons of international holiday visibility.
'Thus, before starting this macro the user needs to delete all other country holidays.
    'Firstly, one must switch to the "list" view of the Outlook calendar and filter for the category "holiday"
    'and delete all appearing entries.
    'Secondly, one must switch to the "list" view of the Outlook calendar and filter for his/her country in
    'order to delete possible remaining entries.
'Afterwards, the user can import via the Outlook options again the holidays of his/her country, execute the
'macro so that only his/her country holidays are set to "out of office". Only subsequently the user should
'import the holidays from other countries as the macro was executed and will not set these holidays also
'to "out of office"

    ' Die Anwendung erfolgt im gerade selektierten Kalender
    Dim olFolder As MAPIFolder
    Set olFolder = Application.ActiveExplorer.CurrentFolder
 
    ' Ist kein Kalender selektiert, wird die Operation übersprungen
    If olFolder.DefaultItemType = olAppointmentItem Then
        Dim olApptItem As Outlook.AppointmentItem
 
        For Each olApptItem In olFolder.Items
            With olApptItem
                ' Nur Termine der Kategorie "Feiertag" bearbeiten
                If InStr(.Categories, "Feiertag") Then
                    ' Nur Termine die als "frei" angezeigt werden bearbeiten
                    If .BusyStatus = olFree Then
                        ' Nur NRW-Feiertage als "Abwesend" markieren
                        If InStr(.ConversationTopic, "Heilige Drei Könige") = 0 And _
                           InStr(.ConversationTopic, "Mariä Himmelfahrt") = 0 And _
                           InStr(.ConversationTopic, "Buß- und Bettag") = 0 And _
                           InStr(.ConversationTopic, "Reformationstag") = 0 Then
                            .BusyStatus = olOutOfOffice
                            .Save
                        Else
                            ' Nicht NRW-Feiertage aus der Kategorie "Feiertag" entfernen
                            .Categories = Replace(.Categories, "Feiertag", "")
                            .Save
                        End If
                    End If
                End If
            End With
        Next olApptItem
    End If
    
End Sub
