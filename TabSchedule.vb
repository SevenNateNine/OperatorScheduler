Imports Microsoft.VisualBasic
Imports System.Data.SqlClient

Partial Public Class OperatorMainForm
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
End Class
