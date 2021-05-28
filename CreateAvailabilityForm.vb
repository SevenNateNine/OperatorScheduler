Imports System.Data.SqlClient

Public Class CreateAvailabilityForm
    Dim conString As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\nchan1\source\repos\OperatorScheduler\OpSchedDatabase.mdf;Integrated Security=True;Connect Timeout=30"
    Dim sdaOperator As SqlDataAdapter
    Dim sdaAvailability As SqlDataAdapter
    Dim dsOperator As New DataTable
    Dim dsAvailability As New DataSet
    Dim changes As DataSet
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        BindComboBox()
    End Sub

    Private Sub BindComboBox()
        Using con As New SqlConnection(conString)
            Dim query As String _
                = "SELECT RTRIM(LTRIM(EmployeeID)) AS EmployeeID, RTRIM(LTRIM(FirstName)) AS FirstName, RTRIM(LTRIM(LastName)) AS LastName FROM Operator"
            Using cmd As New SqlCommand(query, con)
                con.Open()
                cmd.CommandType = CommandType.Text
                sdaOperator = New SqlDataAdapter(cmd)
                dsOperator = New DataTable()
                dsOperator.Load(cmd.ExecuteReader)
                dsOperator.Columns.Add("ID_FullName", GetType(Object), "EmployeeID + ' : ' + FirstName + ' ' + LastName")

                Dim newRow As DataRow = dsOperator.NewRow
                newRow("ID_FullName") = "None"
                newRow("EmployeeID") = "-1"
                dsOperator.Rows.InsertAt(newRow, 0)
                con.Close()

                OperatorComboBox.DataSource = dsOperator
                OperatorComboBox.DisplayMember = "ID_FullName"
                OperatorComboBox.ValueMember = "EmployeeID"

            End Using
        End Using
    End Sub

    Private Sub CreateAvailabilityBtn_Click(sender As Object, e As EventArgs) Handles CreateAvailabilityBtn.Click
        ' check if startDateTime < endDateTime
        Dim startTime = Format(StartTimePicker.Value, "hh:mm:ss tt")
        Dim endTime = Format(EndTimePicker.Value, "hh:mm:ss tt")
        Dim startDate = StartDatePicker.Value.Date
        Dim endDate = EndDatePicker.Value.Date
        Dim operatorID = If(OperatorComboBox.SelectedValue = "-1", Nothing, OperatorComboBox.SelectedValue)
        Dim isScheduled = ScheduledCheckBox.Checked
        Dim repeatedDays() As Integer = {-1, -1, -1, -1, -1, -1, -1}
        ' Check if the Until Date is greater than the starting Date
        If startDate.CompareTo(UntilDateTimePicker.Value) >= 0 Or endDate.CompareTo(UntilDateTimePicker.Value) >= 0 Then
            MessageBox.Show("UntilDateTimePicker must be greater than the selected Date")
            Return
        End If
        Using con As New SqlConnection(conString)
            ' To add repeated availabilities
            If RepeatCheckBox.Checked Then
                Dim query As String
                query = "INSERT INTO Availability (StartDate, StartTime, EndDate, EndTime, "
                If operatorID IsNot Nothing Then
                    query += "OperatorID, "
                End If
                query += "AvailabilityGroup, IsScheduled) VALUES "

                ' what the heck am I doing here? -N8 4/7/21
                ' checking when to repeat this availability. use .isDayOfWeek and if it corresponds to Case then add it to the String? 
                For Each i In DayCheckListBox.CheckedIndices
                    Select Case i
                        ' Sunday
                        Case 0
                            repeatedDays(0) = 0
                        Case 1
                            repeatedDays(1) = 1
                        Case 2
                            repeatedDays(2) = 2
                        Case 3
                            repeatedDays(3) = 3
                        Case 4
                            repeatedDays(4) = 4
                        Case 5
                            repeatedDays(5) = 5
                        'Saturday'
                        Case 6
                            repeatedDays(6) = 6
                    End Select
                Next

                ' From when the availability is first made till the Until DTPicker's value
                Dim initialDate As Date = startDate
                Dim weekValue = If(WeekUpDown.Value > 0, WeekUpDown.Value, 1)
                For i As Integer = startDate.Date.DayOfYear + startDate.Date.Year * 365 To UntilDateTimePicker.Value.DayOfYear + UntilDateTimePicker.Value.Year * 365
                    ' If doing every 1+ weeks. And current date is Sunday, a new week.
                    If WeekUpDown.Value >= 1 And startDate.DayOfWeek = 0 And startDate.CompareTo(initialDate) <> 0 Then
                        startDate = startDate.AddDays(7 * (weekValue - 1))
                        endDate = endDate.AddDays(7 * (weekValue - 1))
                        i += 7 * (weekValue - 1)
                    End If

                    ' If current day is part of the repeated days
                    If repeatedDays.Contains(startDate.DayOfWeek) And startDate.Date.DayOfYear + startDate.Date.Year * 365 < UntilDateTimePicker.Value.DayOfYear + UntilDateTimePicker.Value.Year * 365 Then
                        ' append to value String
                        query += String.Format("({0}, {1}, {2}, {3},", "'" + startDate + "'", "'" + startTime + "'", "'" + endDate + "'", "'" + endTime + "'")
                        If operatorID IsNot Nothing Then
                            query += String.Format("{0}, ", operatorID)
                        End If
                        query += String.Format("(SELECT CASE WHEN COUNT(1) > 0 THEN (Select Top 1 ID From availability ORDER BY Id DESC) ELSE '1' END FROM Availability), {0}),", If(isScheduled, 1, 0))
                    End If

                    startDate = startDate.AddDays(1)
                    endDate = endDate.AddDays(1)
                Next

                ' Remove comma from last record to make query nice.
                query = query.TrimEnd(",")

                ' Run SQL Command.
                Using cmd As New SqlCommand(query, con)
                    With cmd
                        .Connection = con
                        .CommandType = CommandType.Text
                        .CommandText = query
                    End With
                    Try
                        con.Open()
                        cmd.ExecuteNonQuery()
                        con.Close()
                        Close()
                    Catch ex As Exception
                        MessageBox.Show("Line 124: " + ex.Message.ToString(), "Error")
                    End Try
                End Using
            Else
                ' Adds availability to the Database once.
                Dim query As String
                query = "INSERT INTO Availability (startDate, startTime, endDate, endTime,"
                If operatorID IsNot Nothing Then
                    query += "OperatorID,"
                End If
                query += " IsScheduled) VALUES (@startDate, @startTime, @endDate, @endTime,"
                If operatorID IsNot Nothing Then
                    query += "@opID, "
                End If
                query += "@isScheduled)"
                Using cmd As New SqlCommand(query, con)
                    With cmd
                        .Connection = con
                        .CommandType = CommandType.Text
                        .CommandText = query
                        .Parameters.AddWithValue("@startDate", startDate)
                        .Parameters.AddWithValue("@startTime", startTime)
                        .Parameters.AddWithValue("@endDate", endDate)
                        .Parameters.AddWithValue("@endTime", endTime)
                        If operatorID IsNot Nothing Then
                            .Parameters.AddWithValue("@opID", operatorID)
                        End If
                        .Parameters.AddWithValue("@isScheduled", isScheduled)
                    End With
                    Try
                        con.Open()
                        cmd.ExecuteNonQuery()
                        con.Close()
                        MessageBox.Show("Addition was a success!")
                        Me.Close()
                        Dim opMF As OperatorMainForm = OperatorMainForm.ActiveForm
                        opMF.callBindMCG()
                        opMF.RefreshTables()
                    Catch ex As Exception
                        MessageBox.Show(ex.Message.ToString(), "Error")
                    End Try
                End Using
            End If

        End Using

        
    End Sub

    Private Sub RepeatCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles RepeatCheckBox.CheckedChanged
        If RepeatCheckBox.Checked Then
            DayCheckListBox.Enabled = True
            UntilDateTimePicker.Enabled = True
        Else
            DayCheckListBox.Enabled = False
            UntilDateTimePicker.Enabled = False
        End If
    End Sub

    Private Sub WeekUpDown_ValueChanged(sender As Object, e As EventArgs) Handles WeekUpDown.ValueChanged
        If WeekUpDown.Value > 1 Then
            RepeatLabelTwo.Text = "Weeks"
        Else
            RepeatLabelTwo.Text = "Week"
        End If
    End Sub
End Class