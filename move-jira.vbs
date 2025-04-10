Private WithEvents inboxItems As Outlook.Items

Private Sub Application_Startup()
    Set inboxItems = Session.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub inboxItems_ItemAdd(ByVal Item As Object)
    On Error Resume Next

    Dim mail As Outlook.MailItem
    If TypeName(Item) = "MailItem" Then
        Set mail = Item
        Dim subjectText As String
        subjectText = mail.Subject
        Dim senderAddress As String
        senderAddress = LCase(mail.SenderEmailAddress)

        Dim re As Object
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "\b(AS-\d+|APAC-\d+)\b"
        re.IgnoreCase = True

        If re.Test(subjectText) Then
            Dim matches As Object
            Set matches = re.Execute(subjectText)
            Dim ticketID As String
            ticketID = matches(0).Value

            Dim inbox As Outlook.MAPIFolder
            Set inbox = Session.GetDefaultFolder(olFolderInbox)

            ' Locate the BAU folder
            Dim bauFolder As Outlook.MAPIFolder
            Set bauFolder = Nothing
            Dim subFolder As Outlook.MAPIFolder
            For Each subFolder In inbox.Folders
                If subFolder.Name = "BAU" Then
                    Set bauFolder = subFolder
                    Exit For
                End If
            Next

            If bauFolder Is Nothing Then
                Set bauFolder = inbox ' Fallback to Inbox if BAU isn't found
            End If

            ' Look for or create the ticket folder under BAU
            Dim targetFolder As Outlook.MAPIFolder
            Set targetFolder = FindFolderByName(bauFolder, ticketID)

            If targetFolder Is Nothing Then
                Set targetFolder = bauFolder.Folders.Add(ticketID)
            End If

            mail.Move targetFolder
        ElseIf senderAddress = "jira@jdausteam.atlassian.net" Then
            ' JIRA email with no ticket match
            MoveToBAU mail
        End If
    End If
End Sub

Function FindFolderByName(parentFolder As Outlook.MAPIFolder, folderName As String) As Outlook.MAPIFolder
    Dim subFolder As Outlook.MAPIFolder
    For Each subFolder In parentFolder.Folders
        If subFolder.Name = folderName Then
            Set FindFolderByName = subFolder
            Exit Function
        Else
            Dim result As Outlook.MAPIFolder
            Set result = FindFolderByName(subFolder, folderName)
            If Not result Is Nothing Then
                Set FindFolderByName = result
                Exit Function
            End If
        End If
    Next
    Set FindFolderByName = Nothing
End Function

Sub MoveToBAU(mail As Outlook.MailItem)
    On Error Resume Next
    Dim inbox As Outlook.MAPIFolder
    Set inbox = Session.GetDefaultFolder(olFolderInbox)

    Dim bauFolder As Outlook.MAPIFolder
    Set bauFolder = Nothing

    ' Look for the "BAU" folder in Inbox
    Dim subFolder As Outlook.MAPIFolder
    For Each subFolder In inbox.Folders
        If subFolder.Name = "BAU" Then
            Set bauFolder = subFolder
            Exit For
        End If
    Next

    If Not bauFolder Is Nothing Then
        mail.Move bauFolder
    End If
End Sub
