Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Partial Public Class OperatorMainForm
    ''' <summary>
    ''' Adds Message to Console using ConsoleAdd() and inserts Message and MessageType into DB.
    ''' </summary>
    ''' <param name="MessageType"></param>
    ''' MessageType represents what kind of message it is. 
    ''' 0 = Default
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
    ''' "Help" - 
    ''' "Refresh email" - 
    ''' "Send email" - 
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
                HandleUnreadEmails()
            ' Send e-mail of missing availabilities to non-outer operators. Is effected by MonthYear Selection. 
            Case "Send email to inner operators"
                StartEmailChain()
            ' Send e-mail of missing availabilities to ALL operators. Is effected by MonthYear Selection. 
            Case "Send email to ALL operators"
        End Select
    End Sub

    ''' <summary>
    ''' Queries all relevant documentation of the selected month/year and adds it to the console. 
    ''' </summary>
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
    ''' Wrapper for GetRelevantDocumentation(). Informs user of changes via MessageBox. 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub FilterButton_Click(sender As Object, e As EventArgs) Handles FilterButton.Click
        GetRelevantDocumentation()
        MessageBox.Show(String.Format("Now showing documentation regarding the {0} schedule.", MonthYearPicker.Value.ToString("MMM yyyy")))
    End Sub

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
                    If Not reader.Read() Then
                        Logger(String.Format("All availabilities of {0}/{1} have been filled!", MonthYearPicker.Value.Month, MonthYearPicker.Value.Year))
                        Return Array.Empty(Of List(Of String))()
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

    Private Function GetMAWrapper() As String()
        Dim lists As List(Of String)() = GetMissingAvailabilities()
        If lists Is Array.Empty(Of List(Of String))() Then
            Return {""}
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

    Private Sub OutgoingEmailHandler(ByVal openIDs As String, ByVal openAvailabilities As String, ByVal monthYear As String)
        ' Query returns list of INSIDER operators ordered by extra shifts, and seniority. 
        Dim nextEmailed As String() = GetNextToBeEmailed(MonthYearPicker.Value)

        ' Send email to first choice. Mark their ME.GotEmailed = True. Ball is then passed to ReadEmail()
        Dim subject As String = String.Format("{0}/{1} Open Availability Selection", monthYear.Split("/")(0), monthYear.Split("/")(2))
        Dim options As String = openIDs
        Dim body As String = String.Format("Please select an number option that corresponds to the desired availability date that you would like to fill.{0}If you do not want any, select the 'Pass' option. {1}", vbLf, openAvailabilities)
        Logger(String.Format("Sending email offer to {0} {1} using email, {2} ", nextEmailed(1), nextEmailed(2), nextEmailed(0)))

        ' send email here
        coHandler.sendOptionEmail(oApp, {"nchan1@numc.edu"}, subject, options, body)
    End Sub

    Private Sub StartEmailChain()
        Dim query As String
        ConsoleAdd("Collecting unassigned shifts.")
        ' Get missing availabilities. If none. End chain.
        Dim gmawResults As String() = GetMAWrapper()
        Dim openIDs As String = gmawResults(0)
        Dim openAvailabilities As String = gmawResults(1)
        If openIDs.Length = 0 Then
            Return
        End If

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
    ''' 
    ''' </summary>
    Private Sub ContinueEmailChain(monthYear As String)
        ' Get missing availabilities. 
        Dim gmawResults As String() = GetMAWrapper()
        Dim openIDs As String = gmawResults(0)
        Dim openAvailabilities As String = gmawResults(1)
        ' If none. End chain.
        If openIDs.Length = 0 Then
            Return
        End If

        OutgoingEmailHandler(openIDs, openAvailabilities, monthYear)
    End Sub

    Private Sub HandleAvailabilityAcceptance(ByVal email As String, ByVal id As String)
        Dim query As String = "UPDATE Availability SET OperatorID = (SELECT TOP 1 EmployeeID FROM Operator WHERE Email = @Email) WHERE Id = @Id"
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
                    .Parameters.AddWithValue("@MonthYear", monthYear)
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
    Private Sub HandleUnreadEmails()
        ' Reads unread emails.
        Logger("Received request to read inbox.")
        Dim inboxItems As Outlook.Items = coHandler.readEmails(oApp)
        Dim i As Integer
        Dim oMsg As Outlook.MailItem
        Dim msgSubject As String
        Dim msgEmail As String
        Dim msgUsername As String
        Dim msgBody As String

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
                Subject: '{3}'
                Message: '{4}'
                *********************", msgUsername, msgEmail, oMsg.ReceivedTime, msgSubject, msgBody)

            msgSplitCode = Split(msgSubject, ":")(0)
            msgSplitMonthYear = Split(Split(msgSubject, ":")(1).Trim())(0)
            If msgSplitCode.Equals("Pass") Then
                ConsoleAdd(String.Format("{0} has passed their schedule offer.", msgEmail))
            ElseIf Integer.TryParse(msgSplitCode, msgSubAsInt) Then
                ConsoleAdd(String.Format("{0} has taken the availability with the ID, {1}", msgEmail, msgSplitCode))
                ' Updates availability by ID.
                HandleAvailabilityAcceptance(msgEmail, msgSplitCode)
            Else
                ConsoleAdd("Unknown request. Skip.")
            End If
            ' Updates corresponding MEC (using msgSplitMonthYear and email to identify) and move to next email
            UpdateMonthEmailCheck(msgEmail, msgSplitMonthYear)

            ' Marks message as Read to avoid being read again. 
            oMsg.UnRead = False
        Next
        ConsoleAdd(String.Format("Unread messages from inbox: {0}{1}", vbLf, consoleMsg))

        ' Continue email chain
        ' Continue Chain
        ContinueEmailChain(MonthYearPicker.Value)
    End Sub

    Private Sub SendEmailRequest()
        Try
            coHandler.sendOptionEmail(oApp, {"nchan1@numc.edu"}, "options test", "Option1;Option2;Option3;Option4;Option5;Option6", "body")


        Catch error_t As Exception
            MsgBox(error_t.ToString)
        End Try
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
