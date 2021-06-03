Imports Microsoft.Office.Interop

Public Class CustomOutlookHandler
    Dim acc As Outlook.Account
    Dim email As String
    Dim password As String
    Sub New(arg_acc As Outlook.Account, arg_email As String, arg_pw As String)
        acc = arg_acc
        email = arg_email
        password = arg_pw
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="oApp"></param>
    ''' <param name="recipients"></param>
    ''' <param name="subject"></param>
    ''' <param name="body"></param>
    ''' <param name="debug"></param>
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
            oMsg.SendUsingAccount = acc
            oMsg.Send()
        Catch ex As Exception
            Console.WriteLine(ex.Message.ToString())
        End Try
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="oApp"></param>
    ''' <param name="recipients"></param>
    ''' <param name="subject"></param>
    ''' <param name="votingOptions"></param>
    ''' <param name="body"></param>
    ''' <param name="debug"></param>
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
            oMsg.SendUsingAccount = acc
            oMsg.Send()
        Catch ex As Exception
            Console.WriteLine(ex.Message.ToString())
        End Try
    End Sub

    Function readEmails(oApp As Outlook.Application) As Outlook.Items
        oApp.Session.Logon(email, password)
        Dim oNS As Outlook.NameSpace = oApp.GetNamespace("mapi")
        'Dim oInbox As Outlook.MAPIFolder = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
        Dim oInbox As Outlook.MAPIFolder = oNS.Folders(email).Folders("Inbox")
        Dim oItems As Outlook.Items = oInbox.Items
        oItems = oItems.Restrict("[Unread] = true")

        oNS.Logoff()

        Return oItems
    End Function

    ''' <summary>
    '''     Gets e-mail address of sender from MailItem
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
    '''     Get username of sender from MailItem
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
