Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Partial Public Class OperatorMainForm
    Dim helpText As String = String.Format("
        Read Unread Emails - Reads and handles unread emails concerning extra shifts requests, then continues email chain. Is NOT affected by MonthYear Selection. 
        Send Email Requests - Sends email concerning unfilled shifts to most senior operator with the least amount of extra shifts. Is affected by MonthYear Selection.
        Reset Extra Shift Count - Resets shift count of all operators in database. Is NOT affected by MonthYear Selection.")

    ''' <summary>
    ''' Adds Message to Console using ConsoleAdd() and inserts Message and MessageType into DB.
    ''' </summary>
    ''' <param name="MessageType"></param>
    ''' MessageType represents what kind of message it is. 
    ''' 0 = Default
    ''' 1 = User Command
    ''' 2 = Email Response
    ''' 8 = Universal Message
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
                    .Parameters.AddWithValue("@DateTimeTarget", If(MessageType = 8, DBNull.Value, MonthYearPicker.Value))
                End With
                Try
                    con.Open()
                    cmd.ExecuteNonQuery()
                    con.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "Logger() Error")
                End Try
            End Using
        End Using

        ' add to rich text element
        ConsoleAdd(Message)
    End Sub

    ''' <summary>
    ''' Queries all relevant documentation of the selected month/year and adds it to the console. 
    ''' </summary>
    Private Sub GetRelevantDocumentation()
        Dim query As String = "SELECT * FROM Documentation WHERE (MONTH(DateTimeTarget) = MONTH(@MonthYear) AND YEAR(DateTimeTarget) = YEAR(@MonthYear)) OR DateTimeTarget IS NULL ORDER BY TimeStamp"
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
    ''' Wrapper for GetRelevantDocumentation(). Informs user of changes via MessageBox. 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub FilterButton_Click(sender As Object, e As EventArgs) Handles FilterButton.Click
        GetRelevantDocumentation()
        MessageBox.Show(String.Format("Now showing documentation regarding the {0} schedule.", MonthYearPicker.Value.ToString("MMM yyyy")))
    End Sub

    ''' <summary>
    ''' Read Unread Emails - Reads e-mail messages and carries out reasonable responses. Is NOT affected by MonthYear Selection. 
    ''' Send Email Requests - Send e-mail of missing availabilities to non-outer operators. Is affected by MonthYear Selection.
    ''' Reset Extra Shift Count - Resets shift count of all operators in database. Is NOT affected by MonthYear Selection. 
    ''' Help - Slaps a bunch of text in the ConsoleRichTextBox. Is NOT affected by MonthYear Selection. 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ActionButton_Click(sender As Object, e As EventArgs) Handles ActionButton.Click
        Select Case ActionComboBox.SelectedItem.ToString()
            ' Slaps a bunch of text in the ConsoleRichTextBox. Is NOT affected by MonthYear Selection. 
            Case "Help"
                ConsoleAdd(helpText)
            ' Reads e-mail messages and carries out reasonable responses. Is NOT affected by MonthYear Selection. 
            Case "Read Unread Emails"
                HandleUnreadEmails()
            ' Send e-mail of missing availabilities to non-outer operators. Is affected by MonthYear Selection. 
            Case "Send Email Requests"
                StartEmailChain()
            ' Resets shift count of all operators in database. Is NOT affected by MonthYear Selection. 
            Case "Reset Extra Shift Count"
                ResetExtraShiftCount()
        End Select
    End Sub
    ''' <summary>
    ''' Gets all Availabilities pertaining to the specific month/year where no Operator is assigned. 
    ''' </summary>
    ''' <returns>
    '''     An empty list if there are no availabilities that meet the criteria.
    '''     Two lists, one containing a visual representation of unassigned shifts, and another containing just the IDs for voting.
    ''' </returns>
    Private Function GetMissingAvailabilities() As List(Of String)()
        Dim output As New List(Of String)()
        Dim idOutput As New List(Of String)()
        Dim query As String = "SELECT * FROM Availability WHERE OperatorID IS NULL AND MONTH(StartDate) = MONTH(@MonthYear) AND YEAR(StartDate) = YEAR(@MonthYear)"
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
                    ' If no records found. Return.
                    If Not reader.HasRows() Then
                        Logger(String.Format("All availabilities of {0}/{1} have been filled!", MonthYearPicker.Value.Month, MonthYearPicker.Value.Year))
                        Return Nothing
                    End If
                    While reader.Read()
                        output.Add(String.Format("{0}: {1} {2} - {3} {4}", reader("Id").ToString(), reader("StartDate").ToString().Split(" ")(0), reader("StartTime").ToString(), reader("EndDate").ToString().Split(" ")(0), reader("EndTime").ToString()))
                        idOutput.Add(reader("Id").ToString())
                    End While
                    con.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "Error")
                End Try
            End Using
        End Using
        Return {output, idOutput}
    End Function
    ''' <summary>
    ''' Wrapper that calls GetMissingAvailabilities() 
    ''' </summary>
    ''' <returns>
    ''' If there are no free availabilities, returns a list containing an empty String.
    ''' Returns two Strings, one containing the options used for voting, the other a visual representation of unassigned shifts. 
    ''' </returns>
    Private Function GetMAWrapper() As String()
        Dim lists As List(Of String)() = GetMissingAvailabilities()
        If lists Is Nothing Then
            Return Nothing
        End If
        Dim output As List(Of String) = lists(0)
        Dim idOutput As List(Of String) = lists(1)

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

        Return {openIDs, openAvailabilities}
    End Function
    ''' <summary>
    ''' All MonthEmailCheck records related to the Month Year are set to False. Used when all MEC records are set to True and there are unassigned shifts remaining. 
    ''' </summary>
    ''' <param name="monthYear"></param>
    Private Sub ClearRelevantMECRecords(ByVal monthYear As String)
        Dim query As String = "UPDATE MonthEmailCheck SET GotEmailed = 0 WHERE MONTH(MonthYear) = MONTH(@MonthYear) AND YEAR(MonthYear) = YEAR(@MonthYear)"
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand(query, con)
                With cmd
                    .Connection = con
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.AddWithValue("@MonthYear", monthYear)
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
    End Sub
    ''' <summary>
    '''     Queries the database to determine which operator will be offered extra shifts next based on seniority and how many extra shifts they currently have accumulated.\
    ''' </summary>
    ''' <param name="monthYear">
    '''     Selects Operators who hasn't been emailed for that monthYear's schedule.
    ''' </param>
    ''' <returns>
    '''     A String array that contains the operator's email, firstname, and lastname
    ''' </returns>
    Private Function GetNextToBeEmailed(ByVal monthYear As String) As String()
        Dim returnArray As String() = {"", "", ""}
        Dim query As String = "SELECT TOP 1 EmployeeID, FirstName, LastName, Email FROM Operator As O INNER JOIN MonthEmailCheck As M ON O.Id = M.OperatorID 
        WHERE M.GotEmailed = 0 AND MONTH(M.MonthYear) = MONTH(@MonthYear) AND YEAR(M.MonthYear) = YEAR(@MonthYear) AND Seniority != -1 AND IsOutside = 0 ORDER BY ExtraShifts ASC, Seniority ASC"
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand(query, con)
                With cmd
                    .Connection = con
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.AddWithValue("@MonthYear", monthYear)
                End With
                Try
                    con.Open()
                    Dim reader As SqlDataReader = cmd.ExecuteReader()
                    ' If no records found. Return
                    If Not reader.Read() Then
                        ' Clear all relevants MEC records
                        ClearRelevantMECRecords(monthYear)
                        Logger("Error. Please check database.")
                        Return Array.Empty(Of String)()
                    End If
                    returnArray(0) = reader("Email").ToString().Trim()
                    returnArray(1) = reader("FirstName").ToString().Trim()
                    returnArray(2) = reader("LastName").ToString().Trim()
                    con.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "Error")
                End Try
            End Using
        End Using

        Return returnArray
    End Function
    ''' <summary>
    '''     Gets the next person to be emailed and send them an option email.
    ''' </summary>
    ''' <param name="openIDs">
    '''     Voting options the operator will use to vote.
    ''' </param>
    ''' <param name="openAvailabilities">
    '''     The schedule in a visual format so the user will know which voting option corresponds to which shift.
    ''' </param>
    ''' <param name="monthYear">
    '''     The schedule the email will be sent for.
    ''' </param>
    Private Sub OutgoingEmailHandler(ByVal openIDs As String, ByVal openAvailabilities As String, ByVal monthYear As String)
        ' Query returns list of INSIDER operators ordered by extra shifts, and seniority. 
        Dim nextEmailed As String() = GetNextToBeEmailed(MonthYearPicker.Value)

        ' Send email to first choice. Mark their ME.GotEmailed = True. Ball is then passed to ReadEmail()
        Dim subject As String = String.Format("{0}/{1} Open Availability Selection", monthYear.Split("/")(0), monthYear.Split("/")(2))
        Dim options As String = openIDs
        Dim body As String = String.Format("Please select an number option that corresponds to the desired availability date that you would like to fill.{0}If you do not want any, select the 'Pass' option. {1}", vbLf, openAvailabilities)
        Logger(String.Format("Sending email offer to {0} {1} using email, {2} ", nextEmailed(1), nextEmailed(2), nextEmailed(0)))

        ' Replace 
        coHandler.sendOptionEmail(oApp, {nextEmailed(0)}, subject, options, body)
    End Sub
    ''' <summary>
    '''     Begins email chain by getting the unassigned shifts, creating MonthEmailCheck records for the given month, then passing the unassigned shifts to OutgoingEmailHandler(). 
    ''' </summary>
    Private Sub StartEmailChain()
        Dim query As String
        ConsoleAdd("Collecting unassigned shifts.")
        ' Get missing availabilities. If none. End chain.
        Dim gmawResults As String() = GetMAWrapper()
        If gmawResults Is Nothing Then
            Return
        End If
        Dim openIDs As String = gmawResults(0)
        Dim openAvailabilities As String = gmawResults(1)


        ' For each INSIDER operator create MEC records for the selected month.
        query = "BEGIN 
	                IF NOT EXISTS (SELECT * FROM MonthEmailCheck WHERE MONTH(MonthYear) = MONTH(@MonthYear) AND YEAR(MonthYear) = YEAR(@MonthYear) )
	                BEGIN 
		                INSERT INTO MonthEmailCheck (OperatorID, MonthYear) SELECT Id, @MonthYear FROM Operator WHERE IsOutside = 0;
                        INSERT INTO MonthEmailCheck(OperatorID, MonthYear, GotEmailed) Select Id, @MonthYear, 1 From Operator WHERE IsOutside = 1;
	                END
                END"
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
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "Error")
                End Try
            End Using
        End Using

        OutgoingEmailHandler(openIDs, openAvailabilities, MonthYearPicker.Value)
    End Sub
    ''' <summary>
    '''     Assigns unassigned shift to operator and increments their extrashift count by 1
    ''' </summary>
    ''' <param name="email"></param>
    ''' <param name="id"></param>
    Private Sub HandleAvailabilityAcceptance(ByVal email As String, ByVal id As String)
        Dim query As String = "BEGIN TRANSACTION;
            UPDATE Availability SET OperatorID = (SELECT TOP 1 EmployeeID FROM Operator WHERE Email = @Email) WHERE Id = @Id;
            UPDATE Operator SET ExtraShifts = ExtraShifts + 1 WHERE EmployeeID = (SELECT TOP 1 EmployeeID FROM Operator WHERE Email = @Email)
            COMMIT;"
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand(query, con)
                With cmd
                    .Connection = con
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.AddWithValue("@Id", id)
                    .Parameters.AddWithValue("@Email", email)
                End With
                Try
                    con.Open()
                    cmd.ExecuteNonQuery()
                    con.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "HandleAvailabilityAcceptance() Error")
                End Try
            End Using
        End Using
    End Sub
    ''' <summary>
    '''     
    ''' </summary>
    ''' <param name="email"></param>
    ''' <param name="monthYear"></param>
    Private Sub UpdateMonthEmailCheck(ByVal email As String, ByVal monthYear As String)
        Dim query As String = "UPDATE MonthEmailCheck SET GotEmailed = 1 
            WHERE OperatorID = (SELECT TOP 1 EmployeeID FROM Operator WHERE Email = @Email) 
            AND MONTH(MonthYear) = MONTH(@MonthYear) AND YEAR(MonthYear) = YEAR(@MonthYear)"
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand(query, con)
                With cmd
                    .Connection = con
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.AddWithValue("@MonthYear", CDate(monthYear))
                    .Parameters.AddWithValue("@Email", email)
                End With
                Try
                    con.Open()
                    cmd.ExecuteNonQuery()
                    con.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "UpdateMonthEmailCheck() Error")
                End Try
            End Using
        End Using
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="monthYear"></param>
    Private Sub ContinueEmailChain(monthYear As String)
        ' Get missing availabilities. 
        Dim gmawResults As String() = GetMAWrapper()
        ' If none. End chain.
        If gmawResults Is Nothing Then
            Return
        End If
        Dim openIDs As String = gmawResults(0)
        Dim openAvailabilities As String = gmawResults(1)

        OutgoingEmailHandler(openIDs, openAvailabilities, monthYear)
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    Private Sub HandleUnreadEmails()
        ' Reads unread emails.
        Logger(String.Format("Received request to read inbox of {0}.", designatedEmail), 8)
        Dim inboxItems As Outlook.Items = coHandler.readEmails(oApp)
        Dim i As Integer
        Dim oMsg As Outlook.MailItem
        Dim msgSubject As String
        Dim msgEmail As String
        Dim msgUsername As String
        Dim msgBody As String

        Dim msgSplit As String()
        Dim msgSplitCode As String
        Dim msgSplitMonthYear As String
        Dim msgSubAsInt As Integer

        Dim consoleMsg As String = ""
        If inboxItems.Count <= 0 Then
            ConsoleAdd("There are no unread emails.")
            Return
        End If
        For i = 1 To inboxItems.Count
            oMsg = inboxItems.Item(i)
            msgSubject = oMsg.Subject
            msgEmail = coHandler.getEmailAddress(oMsg)
            msgUsername = coHandler.getUsername(oMsg)
            msgBody = If(Not IsNothing(oMsg.Body), oMsg.Body, "[No message attached]")
            consoleMsg += String.Format("
                Sender Email: {0} ({1})
                Time Sent: {2}
                Subject: {3}
                Message: {4}
                *********************", msgUsername, msgEmail, oMsg.ReceivedTime, msgSubject, If(msgBody.Length > 25, msgBody.Substring(0, 25) + "...", msgBody))

            msgSplit = Split(msgSubject, ":")
            msgSplitCode = msgSplit(0)

            If msgSplitCode.Equals("Pass") Then
                ConsoleAdd(String.Format("{0} has passed their schedule offer.", msgEmail))
            ElseIf Integer.TryParse(msgSplitCode, msgSubAsInt) Then
                ConsoleAdd(String.Format("{0} has taken the availability with the ID, {1}", msgEmail, msgSplitCode))
                ' Updates availability by ID.
                HandleAvailabilityAcceptance(msgEmail, msgSplitCode)
            Else
                ConsoleAdd("Unknown request. Skip.")
            End If

            If msgSplitCode.Equals("Pass") Or Integer.TryParse(msgSplitCode, msgSubAsInt) Then
                ' Updates corresponding MEC (using msgSplitMonthYear and email to identify) and move to next email
                msgSplitMonthYear = Split(msgSplit(1).Trim())(0)
                UpdateMonthEmailCheck(msgEmail, msgSplitMonthYear)
                ContinueEmailChain(msgSplitMonthYear)
            End If

            ' Marks message as Read to avoid being read again. 
            oMsg.UnRead = False
        Next
        ConsoleAdd(String.Format("Unread messages from inbox: {0}{1}", vbLf, consoleMsg))
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    Private Sub ResetExtraShiftCount()
        Dim query As String = "UPDATE Operator SET ExtraShifts = 0"
        Using con As New SqlConnection(conString)
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
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "ResetExtraShiftCount() Error")
                End Try
            End Using
        End Using
        Logger("Reset Extra Shift count for ALL operators.", 8)
    End Sub



    ''' <summary>
    ''' Returns True if email arg is in the database. False otherwise. 
    ''' </summary>
    ''' <param name="argEmail"></param>
    ''' <param name="debug"></param>
    ''' <returns></returns>
    Function checkEmail(argEmail As String, Optional debug As Boolean = False) As Boolean
        Dim query As String =
            String.Format("SELECT * FROM Operator As O WHERE O.Email = '{0}'", argEmail)
        Using con As New SqlConnection(conString)
            If con.State = ConnectionState.Closed Then
                con.ConnectionString = conString
            End If
            Dim cmd As SqlCommand
            cmd = con.CreateCommand
            cmd.CommandText = query
            con.Open()
            Using sqlRdr As SqlDataReader = cmd.ExecuteReader()
                If sqlRdr.Read() Then
                    Return True
                End If
            End Using
            con.Close()
        End Using
        Return False
    End Function
End Class
