Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Public Class OperatorMainForm
    Dim conString As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\nchan1\source\repos\OperatorScheduler\OpSchedDatabase.mdf;Integrated Security=True;Connect Timeout=30"
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

    Private Sub New()
        oApp = New Outlook.Application
        acc = Nothing
        For Each account As Outlook.Account In oApp.Session.Accounts
            If account.SmtpAddress.Equals("opot@numc.edu", StringComparison.CurrentCultureIgnoreCase) Then
                acc = account
                Exit For
            End If
        Next

        If acc IsNot Nothing Then

        Else
            Throw New Exception("Please login to opot@numc.edu on Outlook before proceeding.")
        End If

        coHandler = New CustomOutlookHandler(acc)
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
