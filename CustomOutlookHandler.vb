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
End Class
