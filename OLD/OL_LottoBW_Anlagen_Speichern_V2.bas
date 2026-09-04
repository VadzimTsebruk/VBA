Attribute VB_Name = "OL_LottoBW_Anlagen_Speichern_V2"
Option Explicit
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/                                                                                                                 '_/
'_/     #Version 2                                                                                                  '_/
'_/     Makro zum Abspeichern von allen Anlagen aus mehreren Outlook Mails in einem festgelegten Ordner.            '_/
'_/     Gleiche Dateien werden nicht gespeichert                                                                    '_/
'_/     Autor: Vadzim Tsebruk, 08.04.2016                                                                           '_/
'_/-----------------------------------------------------------------------------------------------------------------'_/
'_/     #Version 1                                                                                                  '_/
'_/     Makro zum Abspeichern von allen Anlagen aus mehreren Mails                                                  '_/
'_/     Autor: Vadzim Tsebruk, 02.02.2016                                                                           '_/
'_/                                                                                                                 '_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/                                                                                                                 '_/
'_/     Credits                                                                                                     '_/
'_/     http://www.office-loesung.de/ftopic44816_0_0_asc.php                                                        '_/
'_/                                                                                                                 '_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    
    Sub OL_LottoBW_alle_Anlagen_mehrerer_Mails_in_einem_Ordner_speichern()
    
    Dim olExplorer                      As Explorer
    Dim olFolder                        As MAPIFolder
    Dim olSelection                     As Selection
    Dim olitem                          As MailItem
    
    Dim FSobj                           As Object
    
    Dim strDateiname                    As String
    
    Dim lngAttCount                     As Long
    Dim i                               As Long

    Dim lngAnzahlneuerAnlagen           As Long
    Dim lngAnzahlvorhanderenerAnlagen   As Long
    
    Dim varfcPath                       As Variant
    
    
    
    '######################################################################!!!PRAEMISSEN!!!######################################################################
    'Die Abnahmeprotokolle im Kundenprojekt Lotto BW NewNet2016 basieren auf Version 5
    
    
    
    '################################################################!!!INDIVIDUELLE ANPASSUNG!!!################################################################
    'Const varfcPath As String = "C:\HIERPFADEINTRAGEN\"                       'Durch Entfernung der Auskommentierung besteht die Moeglichkeit einen konstanten _
                                                                                Speicherort fuer die Anlagen festzulegen, sodass man nicht mehr nach dem _
                                                                                Speicherort gefragt wird. _
                                                                                Nach der Entfernung der Auskommentierung fuer einen kosntanten Pfad muss zudem _
                                                                                "varfcPath = Inputbox..." auskommentiert werden
    
2:  varfcPath = InputBox("Speicherort", "Anlagen Speichern unter: ", "E:\Backup_2019-09-24_SanDisk-64GB\Downloads")    'Festlegung des Speicherortes fuer die Anlagen. Laufwerk ist Default auf "C:" _
                                                                                eingestellt kann aber individualisiert werden



    '#########################################################################!!!CODE!!!#########################################################################
    
    '/////Fehlerpruefung\\\\\
    
    'Wenn in User-Abfrage Ordnerpfad leer ist oder Speichervorgang abgebrochen wurde, wird das Makro beendet und der User zur Neueingabe aufgefordert
    If varfcPath = "" Then
        MsgBox "Sie haben entweder keinen Ordnerpfad eingegeben oder auf" & " Abbruch gedrueckt." & _
        " Das Makro wird beendet. Bitte versuchen Sie es spaeter erneut.", vbOKOnly + vbExclamation, "Speichervorgang wurde abgebrochen!"
        Exit Sub
    End If
      
    '/////Jede ausgewaehlte Mail wird in einem vorher festgelegtem Ordner gespeichert. Gleichzeitig Ueberpruefung, ob Datei bereits vorhanden ist oder nicht\\\\\
    Set olExplorer = Application.ActiveExplorer
    Set olFolder = Application.ActiveExplorer.CurrentFolder
    lngAnzahlneuerAnlagen = 0
    lngAnzahlvorhanderenerAnlagen = 0
    
    On Error GoTo ErrMsg
    If Dir(varfcPath, vbDirectory) <> "" Then
        If olFolder.DefaultItemType = olMailItem Then
            Set olSelection = olExplorer.Selection
                For Each olitem In olSelection 'Fuer jede Email die Markiert ist
                    lngAttCount = olitem.Attachments.Count
                        If lngAttCount > 0 Then 'Falls es Anlagen gibt
                            For i = lngAttCount To 1 Step -1 'Fuer jede Anlagen
                                With olitem.Attachments.Item(i)
                                    'Falls Datei exisitiert wird dieselbe Datei nicht noch einmal gespeichert, sondern es wird die naechste Datei ausgewaehlt
                                    If CreateObject("Scripting.FileSystemObject").FileExists(varfcPath & "\" & .FileName) Then
                                        lngAnzahlvorhanderenerAnlagen = lngAnzahlvorhanderenerAnlagen + 1
                                        GoTo 1
                                    Else 'Datei existiert noch nicht
                                        strDateiname = .FileName
                                    End If
                                    On Error GoTo ErrMsg
                                    .SaveAsFile varfcPath & "\" & strDateiname 'Datei speichern
                                End With
                                lngAnzahlneuerAnlagen = lngAnzahlneuerAnlagen + 1
1:                          Next i
                        End If
                Next olitem
        Else
            MsgBox "In diesem Ordner befinden sich keine E-Mail-Nachrichten."
        End If
        
        '/////Meldung, dass neue Anlagen gespeichert bzw., dass bereits vorhandene nicht gespeichert wurden. Wenn keine oder keine neuen Anlagen vorhanden waren _
        erscheint ebenfalls eine entsprechende Mitteilung\\\\\
        
        'Messagebox fuer Anzahl neuer Anlagen
        If lngAnzahlneuerAnlagen < 1 Then
            MsgBox "Keine neuen Anlagen gespeichert.", , "Anzahl neuer Anlagen"
        ElseIf lngAnzahlneuerAnlagen < 2 Then
            MsgBox lngAnzahlneuerAnlagen & " neue Anlage wurde gespeichert.", , "Anzahl neuer Anlagen"
        Else
            MsgBox lngAnzahlneuerAnlagen & " neue Anlagen wurden gespeichert.", , "Anzahl neuer Anlagen"
        End If
        
        'Messagebox fuer Anzahl bereits vorhandener Anlagen. Diese werden nicht nochmals gespeichert
        If lngAnzahlvorhanderenerAnlagen < 1 Then
            MsgBox "Keine bereits vorhandenen Anlagen gefunden.", , "ACHTUNG!"
        ElseIf lngAnzahlvorhanderenerAnlagen < 2 Then
            MsgBox "Da " & lngAnzahlvorhanderenerAnlagen & " Anlage bereits vorhanden war, wurde diese NICHT nocheinmal gespeichert.", vbExclamation, "Anlage bereits vorhanden"
        Else
            MsgBox "Da " & lngAnzahlvorhanderenerAnlagen & " Anlagen bereits vorhanden waren, wurden diese NICHT nocheinmal gespeichert.", vbExclamation, "Anlagen bereits vorhanden"
        End If
        
        Exit Sub
        
    Else
        MsgBox "Der zum Speichern der Anlagen eingetragene Pfad existiert nicht oder der Name dieses Verzeichnisses ist ungueltig:" & vbCr & vbCr & "'" & varfcPath & "'" _
        & vbCr & vbCr & "Bitte versuchen Sie es erneut.", vbCritical, "Speichervorgang wurde abgebrochen!"
        GoTo 2
    End If
        
ErrMsg: MsgBox "Sie haben keinen Zugriff auf das Verzeichnis:" & vbCr & vbCr & "'" & varfcPath & "'" & vbCr & vbCr & _
        "Bitte versuchen Sie es erneut.", vbCritical, "Speichervorgang wurde abgebrochen!"
        GoTo 2

End Sub



