Attribute VB_Name = "OL_mdl_AZ_Sortierung"
Option Explicit

' Original Source: https://techniclee.wordpress.com/2010/08/26/sorting-addressees/
' Updated to prevent Run-time error '91'

Sub SortAddresses()
    Dim colNames As New Collection, _
        intCounter As Integer, _
        intIndex As Integer, _
        olkItem As Object, _
        olkRecipient As Outlook.Recipient, _
        olkAddressee As Outlook.Recipient, _
        varName As Variant

    ' --- FIX FOR ERROR 91: IDENTIFY THE ACTIVE ITEM ---
    If TypeName(Application.ActiveWindow) = "Inspector" Then
        ' If an email window is open
        Set olkItem = Application.ActiveInspector.CurrentItem
    ElseIf TypeName(Application.ActiveWindow) = "Explorer" Then
        ' If you are looking at the Inbox list
        If Application.ActiveExplorer.Selection.Count > 0 Then
            Set olkItem = Application.ActiveExplorer.Selection.Item(1)
        End If
    End If

    ' If no item is found, or the item isn't an email/meeting, exit gracefully
    If olkItem Is Nothing Then
        MsgBox "No email or meeting request selected. Please open or select an item first.", vbInformation
        Exit Sub
    End If
    ' --------------------------------------------------

    ' Start sorting logic
    For intCounter = olkItem.Recipients.Count To 1 Step -1
        Set olkRecipient = olkItem.Recipients.Item(intCounter)
        If colNames.Count > 0 Then
            intIndex = 1
            For Each varName In colNames
                ' Sorting logic: Compare names alphabetically
                If intIndex = colNames.Count Then
                    If LCase(olkRecipient.Name) > LCase(varName) Then
                        colNames.Add olkRecipient, LCase(olkRecipient.Name), , intIndex
                        Exit For
                    Else
                        colNames.Add olkRecipient, LCase(olkRecipient.Name), intIndex
                        Exit For
                    End If
                End If
                
                If LCase(olkRecipient.Name) < LCase(varName) Then
                    colNames.Add olkRecipient, LCase(olkRecipient.Name), intIndex
                    Exit For
                End If
                intIndex = intIndex + 1
            Next
        Else
            colNames.Add olkRecipient, LCase(olkRecipient.Name)
        End If
        olkItem.Recipients.Remove intCounter
    Next

    ' Re-add the sorted recipients
    For Each olkAddressee In colNames
        Set olkRecipient = Session.CreateRecipient(olkAddressee.Name)
        olkRecipient.Resolve
        
        If olkRecipient.Resolved Then
            Set olkRecipient = olkItem.Recipients.Add(olkAddressee.Name)
        Else
            ' Fallback to the email address if the name doesn't resolve
            Set olkRecipient = olkItem.Recipients.Add(olkAddressee.Address)
        End If
        olkRecipient.Type = olkAddressee.Type
    Next

    olkItem.Recipients.ResolveAll

    ' Clean up memory
    Set olkRecipient = Nothing
    Set colNames = Nothing
    Set olkItem = Nothing
    
    MsgBox "Recipients sorted alphabetically!", vbInformation, "Success"
End Sub

