Imports Microsoft.Office.Interop
Imports Microsoft.VisualBasic

Public Class CustomOutlookHandler
    Sub sendEmail(oApp As Outlook.Application, recipients() As String, subject As String, body As String, Optional debug As Boolean = False)
        Try
            Dim oMsg As Outlook.MailItem
            oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)
            For Each recipient In recipients
                oMsg.Recipients.Add(recipient)
            Next
            oMsg.Subject = subject
            oMsg.Body = body
            oMsg.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            oMsg.Send()
        Catch ex As Exception
            Console.WriteLine(ex.Message.ToString())
        End Try
    End Sub

    Sub sendOptionEmail(oApp As Outlook.Application, recipients() As String, subject As String, votingOptions As String, body As String, Optional debug As Boolean = False)
        Try
            Dim oMsg As Outlook.MailItem
            oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)
            For Each recipient In recipients
                oMsg.Recipients.Add(recipient)
            Next
            oMsg.Subject = subject
            oMsg.Body = body
            oMsg.VotingOptions = votingOptions
            oMsg.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            oMsg.Send()
        Catch ex As Exception
            Console.WriteLine(ex.Message.ToString())
        End Try
    End Sub

    Function readEmails(oApp As Outlook.Application) As Outlook.Items
        Dim oNS As Outlook.NameSpace = oApp.GetNamespace("mapi")
        ' oNS.Logon() ' to do 
        Dim oInbox As Outlook.MAPIFolder = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
        Dim oItems As Outlook.Items = oInbox.Items
        oItems = oItems.Restrict("[Unread] = true")

        oNS.Logoff()

        Return oItems
    End Function

    ''' <summary>
    ''' Gets e-mail address of sender from MailItem
    ''' </summary>
    ''' <param name="oMail"></param>
    ''' <returns></returns>
    Function getEmailAddress(oMail As Outlook.MailItem) As String
        Dim oSender As Outlook.AddressEntry = oMail.Sender
        Dim emailAddr As String = ""
        If oSender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry Or oSender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then
            Dim exchUser As Outlook.ExchangeUser = oSender.GetExchangeUser()
            If exchUser IsNot Nothing Then
                emailAddr = exchUser.PrimarySmtpAddress
            End If
        Else
            emailAddr = oMail.SenderEmailAddress
        End If

        Return emailAddr
    End Function

    ''' <summary>
    ''' Get username of sender from MailItem
    ''' </summary>
    ''' <param name="oMail"></param>
    ''' <returns></returns>
    Function getUsername(oMail As Outlook.MailItem) As String
        Dim oSender As Outlook.AddressEntry = oMail.Sender
        Dim username As String = ""
        If oSender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry Or oSender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then
            Dim exchUser As Outlook.ExchangeUser = oSender.GetExchangeUser()
            If exchUser IsNot Nothing Then
                username = exchUser.Name
            End If
        Else
            username = oMail.SenderName
        End If

        Return username
    End Function


End Class
