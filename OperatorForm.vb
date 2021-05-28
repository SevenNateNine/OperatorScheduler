Imports System.Data.SqlClient
Imports System
Imports Microsoft.Office.Interop
Imports Microsoft.SqlServer

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

    Dim startDateIndex As Integer = 4
    Dim endDateIndex As Integer = 6
    Dim isScheduledIndex As Integer = 8
    Dim helpText As String = "This application is made "

    Private Sub New()
        oApp = New Outlook.Application
        coHandler = New CustomOutlookHandler()
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

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' 
    ''' OPERATOR TAB
    '''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub BindOperatorGrid()
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand("SELECT * FROM Operator ORDER BY SENIORITY", con)
                cmd.CommandType = CommandType.Text
                sdaOperator = New SqlDataAdapter(cmd)
                dsOperator = New DataSet()
                sdaOperator.Fill(dsOperator, "Operator")
                OperatorDataGridView.DataSource = dsOperator
                OperatorDataGridView.DataMember = "Operator"
                OperatorDataGridView.Columns.Item(0).Visible = False
            End Using
        End Using
    End Sub

    Public Function callBindMCG()
        BindMonthlyCalendarGrid()
    End Function

    Private Sub BindMonthlyCalendarGrid()
        ' Query SHOULD be called in the beginning, not everytime the date changes 
        Dim query As String =
            "SELECT A.ID, O.EmployeeID, O.FirstName, O.LastName, A.StartDate, A.StartTime, A.EndDate, A.EndTime, A.IsScheduled FROM Availability AS A LEFT OUTER JOIN Operator AS O ON A.OperatorID=O.EmployeeID ORDER BY A.StartDate, A.EndDate"
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand(query, con)
                cmd.CommandType = CommandType.Text
                sdaAvailability = New SqlDataAdapter(cmd)
                dsAvailability = New DataSet()
                sdaAvailability.Fill(dsAvailability, "Availability")
                ' Tried to using BindingSource but .Filter was not working -Nathaniel Chan 3/25/2021
                AvailabilityDataGridView.DataSource = dsAvailability
                AvailabilityDataGridView.DataMember = "Availability"
                AvailabilityDataGridView.Columns.Item(0).Visible = False
            End Using
        End Using
        FilterAvailability()
    End Sub

    Private Sub SaveBtn_Click(sender As Object, e As EventArgs) Handles SaveBtn.Click
        Dim tblName = "Operator"
        Dim cmd
        'Dim con As New SqlConnection(conString)
        'con.Open()
        Using con As New SqlConnection(conString)
            cmd = New SqlCommand("UPDATE Operator SET EmployeeID = @EmployeeID, FirstName = @FirstName, LastName = @LastName, Email = @Email, Seniority = @Seniority WHERE ID = @ID", con)
            cmd.Parameters.Add("@EmployeeID", SqlDbType.Int, 5, "EmployeeID")
            cmd.Parameters.Add("@FirstName", SqlDbType.NVarChar, 20, "FirstName")
            cmd.Parameters.Add("@LastName", SqlDbType.NVarChar, 20, "LastName")
            cmd.Parameters.Add("@Email", SqlDbType.NVarChar, 20, "Email")
            cmd.Parameters.Add("@Seniority", SqlDbType.Int, 1, "Seniority")
            cmd.Parameters.Add("@ID", SqlDbType.Int, 1, "ID")
            sdaOperator.UpdateCommand = cmd

            cmd = New SqlCommand("INSERT INTO Operator (EmployeeID, FirstName, LastName, Email, Seniority) VALUES (@EmployeeID, @FirstName, @LastName, @Email, @Seniority)", con)
            cmd.Parameters.Add("@EmployeeID", SqlDbType.Int, 5, "EmployeeID")
            cmd.Parameters.Add("@FirstName", SqlDbType.NVarChar, 20, "FirstName")
            cmd.Parameters.Add("@LastName", SqlDbType.NVarChar, 20, "LastName")
            cmd.Parameters.Add("@Email", SqlDbType.NVarChar, 20, "Email")
            cmd.Parameters.Add("@Seniority", SqlDbType.Int, 1, "Seniority")
            cmd.Parameters.Add("@ID", SqlDbType.Int, 10, "ID")
            sdaOperator.InsertCommand = cmd

            cmd = New SqlCommand("DELETE FROM Operator WHERE ID = @ID", con)
            cmd.Parameters.Add("@ID", SqlDbType.Int, 1, "ID")
            sdaOperator.DeleteCommand = cmd
            sdaOperator.Update(dsOperator, tblName)
        End Using
        'con.Close()
        RefreshTables()

        MessageBox.Show("Changes saved successfully!")
    End Sub

    Friend Sub RefreshTables()
        BindMonthlyCalendarGrid()
        BindOperatorGrid()
        FilterAvailability()
    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' 
    ''' SCHEDULE TAB
    '''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AvailabilityMonthCalendar_DateChanged(sender As Object, e As DateRangeEventArgs) Handles AvailabilityMonthCalendar.DateChanged
        AvailabilityDataGridView.CurrentCell = Nothing
        FilterAvailability()
    End Sub

    Private Sub IsScheduledCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles IsScheduledCheckBox.CheckedChanged
        FilterAvailability()
    End Sub
    Private Sub ShowAllCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ShowAllCheckBox.CheckedChanged
        If ShowAllCheckBox.Checked Then
            AvailabilityMonthCalendar.Enabled = False
            AvailabilityMonthCalendar.SelectionStart = Now()
            AvailabilityMonthCalendar.SelectionEnd = Now()
            FilterAvailability()
        Else
            AvailabilityMonthCalendar.Enabled = True
            FilterAvailability()
        End If

    End Sub

    ''' <summary>
    ''' Returns True if the row has a startTime or endTime within the Monthly Calendar's selected range. 
    ''' </summary>
    ''' <param name="row"></param>
    ''' <returns></returns>
    Private Function IsWithinDateRange(row) As Boolean
        Dim amcSelRange = AvailabilityMonthCalendar.SelectionRange
        Return CDate(row.Cells(startDateIndex).Value).Date >= amcSelRange.Start.Date _
            And CDate(row.Cells(startDateIndex).Value).Date <= amcSelRange.End.Date _
            Or CDate(row.Cells(endDateIndex).Value).Date <= amcSelRange.End.Date _
            And CDate(row.Cells(endDateIndex).Value).Date >= amcSelRange.Start.Date
    End Function

    ''' <summary>
    ''' Returns True if the IsScheduledCheckBox isn't checked or if it is checked AND the isScheduled value of the row is True. 
    ''' </summary>
    ''' <param name="row"></param>
    ''' <returns></returns>
    Private Function IsScheduledHandler(row) As Boolean
        If IsScheduledCheckBox.Checked Then
            Return row.Cells(isScheduledIndex).Value
        Else
            Return True
        End If

    End Function

    ''' <summary>
    ''' Adjusts row visibility depending on whether or not 
    ''' </summary>
    Private Sub FilterAvailability()
        AvailabilityDataGridView.CurrentCell = Nothing
        For Each row As DataGridViewRow In AvailabilityDataGridView.Rows
            If (IsWithinDateRange(row) Or ShowAllCheckBox.Checked) And IsScheduledHandler(row) Then
                AvailabilityDataGridView.Rows.Item(row.Index).Visible = True
            Else
                AvailabilityDataGridView.Rows.Item(row.Index).Visible = False
            End If
            IsScheduledCellsHighlight(row)
        Next
    End Sub

    ''' <summary>
    ''' Opens the Create Availability form on button click
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AddAvailabilityBtn_Click(sender As Object, e As EventArgs) Handles AddAvailabilityBtn.Click
        Dim availabilityFrm As CreateAvailabilityForm = New CreateAvailabilityForm()
        availabilityFrm.Show()
        RefreshTables()
    End Sub

    ''' <summary>
    ''' Highlights availabilities that are scheduled to green
    ''' </summary>
    ''' <param name="row"></param>
    Private Sub IsScheduledCellsHighlight(row)
        If row.Cells(isScheduledIndex).Value Then
            AvailabilityDataGridView.Rows.Item(row.Index).DefaultCellStyle.BackColor = Color.LightGreen
        Else
            AvailabilityDataGridView.Rows.Item(row.Index).DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AvailabilityDataGridView_CellContentChanged(sender As Object, e As DataGridViewCellEventArgs) Handles AvailabilityDataGridView.CellValueChanged
        For Each row As DataGridViewRow In AvailabilityDataGridView.Rows
            IsScheduledCellsHighlight(row)
        Next
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SaveAvailabilityChangesBtn_Click(sender As Object, e As EventArgs) Handles SaveAvalabilityChangesBtn.Click
        Dim tblName = "Availability"
        Dim cmd
        Dim con As New SqlConnection(conString)

        con.Open()
        Dim changeTheFuture As DialogResult = MessageBox.Show("Apply changes to following availabilities?", "Change Confirmation", MessageBoxButtons.YesNo)
        If changeTheFuture = DialogResult.Yes Then
            ' need to keep the date the same, but change the time 
            cmd = New SqlCommand("DECLARE @StartDiff INT; DECLARE @EndDiff INT;
SELECT @StartDiff = DATEDIFF(Day, (SELECT StartDate FROM Availability WHERE ID = @ID), @StartDate); SELECT @EndDiff = DATEDIFF(Day, (SELECT EndDate FROM Availability WHERE ID = @ID), @EndDate);
UPDATE Availability SET StartDate = DATEADD(Day, @StartDiff, StartDate), StartTime = @StartTime, EndDate = DATEADD(Day, @EndDiff, EndDate), EndTime = @EndTime, OperatorID = @OperatorID, IsScheduled = @IsScheduled 
WHERE AvailabilityGroup = (SELECT AvailabilityGroup FROM Availability WHERE Id = @Id) AND Id >= @Id", con)
            cmd.Parameters.Add("@StartDate", SqlDbType.Date, 10, "StartDate")
            cmd.Parameters.Add("@StartTime", SqlDbType.Time, 10, "StartTime")
            cmd.Parameters.Add("@EndDate", SqlDbType.Date, 10, "EndDate")
            cmd.Parameters.Add("@EndTime", SqlDbType.Time, 10, "EndTime")
            cmd.Parameters.Add("@OperatorID", SqlDbType.Int, 10, "EmployeeID")
            cmd.Parameters.Add("@IsScheduled", SqlDbType.Bit, 1, "IsScheduled")
            cmd.Parameters.Add("@ID", SqlDbType.Int, 10, "ID")
            sdaAvailability.UpdateCommand = cmd

            cmd = New SqlCommand("DECLARE @StartDiff INT; DECLARE @EndDiff INT;
SELECT @StartDiff = DATEDIFF(Day, (SELECT StartDate FROM Availability WHERE ID = @ID), @StartDate); SELECT @EndDiff = DATEDIFF(Day, (SELECT EndDate FROM Availability WHERE ID = @ID), @EndDate);
DELETE FROM Availability 
WHERE AvailabilityGroup = (SELECT AvailabilityGroup FROM Availability WHERE Id = @Id) AND Id >= @Id", con)
            cmd.Parameters.Add("@StartDate", SqlDbType.Date, 10, "StartDate")
            cmd.Parameters.Add("@EndDate", SqlDbType.Date, 10, "EndDate")
            cmd.Parameters.Add("@ID", SqlDbType.Int, 5, "ID")
            sdaAvailability.DeleteCommand = cmd
        Else
            cmd = New SqlCommand("UPDATE Availability SET StartDate = @StartDate, StartTime = @StartTime, EndDate = @EndDate, EndTime = @EndTime, OperatorID = @OperatorID, IsScheduled = @IsScheduled WHERE ID = @ID", con)
            cmd.Parameters.Add("@StartDate", SqlDbType.Date, 10, "StartDate")
            cmd.Parameters.Add("@StartTime", SqlDbType.Time, 10, "StartTime")
            cmd.Parameters.Add("@EndDate", SqlDbType.Date, 10, "EndDate")
            cmd.Parameters.Add("@EndTime", SqlDbType.Time, 10, "EndTime")
            cmd.Parameters.Add("@OperatorID", SqlDbType.Int, 10, "EmployeeID")
            cmd.Parameters.Add("@IsScheduled", SqlDbType.Bit, 1, "IsScheduled")
            cmd.Parameters.Add("@ID", SqlDbType.Int, 10, "ID")
            sdaAvailability.UpdateCommand = cmd

            cmd = New SqlCommand("DELETE FROM Availability WHERE ID = @ID", con)
            cmd.Parameters.Add("@ID", SqlDbType.Int, 5, "ID")
            sdaAvailability.DeleteCommand = cmd
        End If

        sdaAvailability.Update(dsAvailability, tblName)
        RefreshTables()
        con.Close()
        MessageBox.Show("Changes saved successfully!")
    End Sub

    Private Sub RefreshButton_Click(sender As Object, e As EventArgs) Handles RefreshButton.Click
        RefreshTables()
    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' 
    ''' MISSING TAB
    '''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''' <summary>
    ''' adds text to rich textbox
    ''' </summary>
    ''' <param name="Message"></param>
    ''' <param name="timestamp"></param>
    ''' Defaults to current time if none specificed
    Private Sub ConsoleAdd(ByVal Message As String, Optional ByVal timestamp As String = "")
        If timestamp.Length = 0 Then
            timestamp = DateTime.Now.ToString()
        End If
        ConsoleRichTextBox.Text += String.Format("{0}    {1}{2}", timestamp, Message, vbLf)
    End Sub

    ''' <summary>
    ''' adds text to rich textbox and adds to documentation table
    ''' </summary>
    ''' <param name="MessageType"></param>
    ''' MessageType represents what kind of message it is. 
    ''' 1 = User Command
    ''' 2 = Email Response
    ''' <param name="Message"></param>
    Private Sub Logger(ByVal Message As String, Optional ByVal MessageType As Integer = 0)
        ' add documentation to query
        Using con As New SqlConnection(conString)
            ' Adds availability to the Database once.
            Dim query As String = "INSERT INTO Documentation (MessageType, Message, DateTimeTarget) VALUES (@MessageType, @Message, @DateTimeTarget)"
            Using cmd As New SqlCommand(query, con)
                With cmd
                    .Connection = con
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.AddWithValue("@MessageType", MessageType)
                    .Parameters.AddWithValue("@Message", Message)
                    .Parameters.AddWithValue("@DateTimeTarget", MonthYearPicker.Value)
                End With
                Try
                    con.Open()
                    cmd.ExecuteNonQuery()
                    con.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "Error")
                End Try
            End Using
        End Using

        ' add to rich text element
        ConsoleAdd(Message)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ActionButton_Click(sender As Object, e As EventArgs) Handles ActionButton.Click
        Select Case ActionComboBox.SelectedItem.ToString()
            ' Slaps a bunch of text in the ConsoleRichTextBox. Is not effected by MonthYear Selection. 
            Case "Help"
                ConsoleAdd(helpText)
            ' Reads e-mail messages and carries out reasonable responses. Is not effected by MonthYear Selection. 
            Case "Refresh email"
                Logger("User requested to read un-read e-mails.", 1)
            ' Send e-mail of missing availabilities to non-outer operators. Is effected by MonthYear Selection. 
            Case "Send email to inner operators"
                StartEmailChain()
            ' Send e-mail of missing availabilities to ALL operators. Is effected by MonthYear Selection. 
            Case "Send email to ALL operators"
        End Select
    End Sub

    Private Sub GetRelevantDocumentation()
        Dim query As String = "SELECT * FROM Documentation WHERE MONTH(DateTimeTarget) = MONTH(@MonthYear) AND YEAR(DateTimeTarget) = YEAR(@MonthYear) ORDER BY TimeStamp"
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand(query, con)
                With cmd
                    .Connection = con
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.AddWithValue("@MonthYear", MonthYearPicker.Value)
                End With
                Try
                    con.Open()
                    ConsoleRichTextBox.Text = ""
                    Dim reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        ConsoleAdd(reader.GetValue(3).ToString(), reader.GetValue(1).ToString().Trim())
                    End While
                    con.Close()
                    Dim opMF As OperatorMainForm = OperatorMainForm.ActiveForm

                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "Error")
                End Try
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Queries messages of relevant time from DB and adds them to textbox
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub FilterButton_Click(sender As Object, e As EventArgs) Handles FilterButton.Click
        GetRelevantDocumentation()
        MessageBox.Show(String.Format("Now showing documentation regarding the {0} schedule.", MonthYearPicker.Value.ToString("MMM yyyy")))
    End Sub

    Private Sub StartEmailChain()
        Dim query As String
        Dim output As New List(Of String)()
        Dim idOutput As New List(Of String)()
        ConsoleAdd("Collecting unassigned shifts.")
        ' Get missing availabilities. If none. End chain.
        query = "SELECT * FROM Availability WHERE OperatorID IS NULL AND MONTH(StartDate) = MONTH(@MonthYear) AND YEAR(StartDate) = YEAR(@MonthYear)"
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand(query, con)
                With cmd
                    .Connection = con
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.AddWithValue("@MonthYear", MonthYearPicker.Value)
                End With
                Try
                    con.Open()
                    Dim reader As SqlDataReader = cmd.ExecuteReader()
                    ' If no records found. Return
                    If Not reader.Read() Then
                        Logger(String.Format("All availabilities of {0}/{1} have been filled!", MonthYearPicker.Value.Month, MonthYearPicker.Value.Year))
                        Return
                    End If
                    While reader.Read()
                        output.Add(String.Format("{0}: {1} {2} - {3} {4}", reader("Id").ToString(), reader("StartDate").ToString().Split(" ")(0), reader("StartTime").ToString(), reader("EndDate").ToString().Split(" ")(0), reader("EndTime").ToString()))
                        idOutput.Add(reader("Id").ToString())
                    End While
                    con.Close()
                    Dim opMF As OperatorMainForm = OperatorMainForm.ActiveForm
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "Error")
                End Try
            End Using
        End Using

        Dim openAvailabilities As String = vbLf
        Dim openIDs As String = "Pass;"
        For Each item As String In output
            openAvailabilities += vbTab + item + vbLf
        Next
        For Each item As String In idOutput
            openIDs += String.Format("{0};", item)
        Next
        openIDs.TrimEnd(CChar(";"))
        ConsoleAdd(openAvailabilities)

        ' For each insider operator create MEC records for the selected month.
        query = "INSERT INTO MonthEmailCheck (OperatorID, Month) SELECT Id, @MonthYear FROM Operator"
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand(query, con)
                With cmd
                    .Connection = con
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.AddWithValue("@MonthYear", MonthYearPicker.Value)
                End With
                Try
                    con.Open()
                    cmd.ExecuteNonQuery()
                    con.Close()
                    Dim opMF As OperatorMainForm = OperatorMainForm.ActiveForm
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "Error")
                End Try
            End Using
        End Using

        Dim opEmail As String = ""
        Dim opFirstName As String = ""
        Dim opLastName As String = ""
        ' Query returns list of INSIDER operators ordered by extra shifts, and seniority. 
        query = "SELECT TOP 1 EmployeeID, FirstName, LastName, Email FROM Operator As O INNER JOIN MonthEmailCheck As M ON O.Id = M.OperatorID WHERE M.GotEmailed = 0 AND MONTH(M.Month) = 5 AND Seniority != -1 AND IsOutside = 0 ORDER BY ExtraShifts ASC, Seniority ASC"
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand(query, con)
                With cmd
                    .Connection = con
                    .CommandType = CommandType.Text
                    .CommandText = query
                End With
                Try
                    con.Open()
                    Dim reader As SqlDataReader = cmd.ExecuteReader()
                    ' If no records found. Return
                    If Not reader.Read() Then
                        ' Clear all relevants MEC records
                        Logger("Error. Please check database.")
                        Return
                    End If
                    opEmail = reader("Email").ToString().Trim()
                    opFirstName = reader("FirstName").ToString().Trim()
                    opLastName = reader("LastName").ToString().Trim()
                    ConsoleAdd(String.Format("{0} {1} is the most senior operator with the least amount of extra shifts.", opFirstName, opLastName))
                    con.Close()
                    Dim opMF As OperatorMainForm = OperatorMainForm.ActiveForm
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "Error")
                End Try
            End Using
        End Using

        ' Send email to first choice. Mark their ME.GotEmailed = True. Ball is then passed to ReadEmail()
        Dim subject As String = String.Format("{0}/{1} Open Availability Selection", MonthYearPicker.Value.Month, MonthYearPicker.Value.Year)
        Dim options As String = openIDs
        Dim body As String = String.Format("Please select an number option that corresponds to the desired availability date that you would like to fill. {0}If you do not want any, select the 'Pass' option. {1}", vbLf, openAvailabilities)
        Logger(String.Format("Sending email offer to {0} {1} using email, {2} ", opFirstName, opLastName, opEmail))
        ' coHandler.sendOptionEmail(oApp, {"nchan1@numc.edu"}, subject, options, body)
    End Sub

    Private Sub ContinueEmailChain()
        ' Get missing availabilities. If none. End chain.

        ' Get the most senior person with the least amount of extra shifts who was not e-mailed already. 
        Dim query As String =
            "SELECT EmployeeID, FirstName, LastName, Email FROM Operator As O INNER JOIN MonthEmailCheck As M ON O.Id = M.OperatorID WHERE M.GotEmailed = 0 AND MONTH(M.Month) = 5 AND Seniority != -1 ORDER BY ExtraShifts, Seniority ASC"
        ' After first iteration, include outside operators.

        ' When query returns 0. Select MonthEmailCheck (MEC) records that correlate to the month and reset GotEmailed to 0. Create MEC records for Outside operators if they don't exist
    End Sub

    Private Sub SendEmailRequest()
        Try
            coHandler.sendOptionEmail(oApp, {"nchan1@numc.edu"}, "options test", "Option1;Option2;Option3;Option4;Option5;Option6", "body")


        Catch error_t As Exception
            MsgBox(error_t.ToString)
        End Try
    End Sub
End Class
