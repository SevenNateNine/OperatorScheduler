Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports System.Xml
Imports System.IO

Public Class OperatorMainForm
    Dim conString As String = ""
    Dim designatedEmail As String = ""
    Dim designatedEmailPass As String = ""
    Dim sdaOperator As SqlDataAdapter
    Dim sdaAvailability As SqlDataAdapter
    Dim dsOperator As New DataSet
    Dim dsAvailability As New DataSet
    Dim changes As DataSet
    Dim oApp As Outlook.Application
    Dim coHandler As CustomOutlookHandler
    ' Dim oNameSpace As Outlook.NameSpace = oApp.GetNamespace("mapi")
    Dim acc As Outlook.Account

    Dim startDateIndex As Integer = 4
    Dim endDateIndex As Integer = 6
    Dim isScheduledIndex As Integer = 8
    Private Function TreatXML(xmlValue As String)
        Return System.Text.RegularExpressions.Regex.Replace(xmlValue, "\s+", " ").Trim()
    End Function
    ''' <summary>
    '''     Sets global values to configuration file
    ''' </summary>
    Private Sub ReadFromXML()
        Dim xd As XDocument = XDocument.Load("\\auricle\Communications\OperatorSchedulerApplication\Config\Defaults.xml")
        designatedEmail = TreatXML(xd.<root>.<email>.<email_address>.Value)
        designatedEmailPass = TreatXML(xd.<root>.<email>.<email_password>.Value)
        conString = TreatXML(xd.<root>.<database>.<connection_string>.Value)
    End Sub

    Private Sub New()
        ReadFromXML()
        oApp = New Outlook.Application
        acc = Nothing
        For Each account As Outlook.Account In oApp.Session.Accounts
            If account.SmtpAddress.Equals(designatedEmail, StringComparison.CurrentCultureIgnoreCase) Then
                acc = account
                Exit For
            End If
        Next

        If acc IsNot Nothing Then

        Else
            Throw New Exception(String.Format("Please login to {0} on Outlook before proceeding.", designatedEmail))
        End If

        coHandler = New CustomOutlookHandler(acc, designatedEmail, designatedEmailPass)
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        BindOperatorGrid()
        BindMonthlyCalendarGrid()
        FilterAvailability()

        GetRelevantDocumentation()

        ' Missing
        MonthYearPicker.Format = DateTimePickerFormat.Custom
        MonthYearPicker.CustomFormat = "MM/yyyy"

    End Sub

    Friend Sub RefreshTables()
        BindMonthlyCalendarGrid()
        BindOperatorGrid()
        FilterAvailability()
    End Sub
End Class
