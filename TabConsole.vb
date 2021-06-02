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
                    ' If no records found. Return.
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

        ' For each INSIDER operator create MEC records for the selected month.
        query = "BEGIN 
	                IF NOT EXISTS (SELECT * FROM MonthEmailCheck WHERE MONTH(MonthYear) = MONTH(@MonthYear) AND YEAR(MonthYear) = YEAR(@MonthYear) )
	                BEGIN 
		                INSERT INTO MonthEmailCheck (OperatorID, MonthYear) SELECT Id, @MonthYear FROM Operator WHERE IsOutside = 0 
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
        query = "SELECT TOP 1 EmployeeID, FirstName, LastName, Email FROM Operator As O INNER JOIN MonthEmailCheck As M ON O.Id = M.OperatorID WHERE M.GotEmailed = 0 AND MONTH(M.MonthYear) = 5 AND Seniority != -1 AND IsOutside = 0 ORDER BY ExtraShifts ASC, Seniority ASC"
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

    ''' <summary>
    ''' 
    ''' </summary>
    Private Sub ContinueEmailChain()
        ' Get missing availabilities. If none. End chain.

        ' Get the most senior person with the least amount of extra shifts who was not e-mailed already. 
        Dim query As String =
            "SELECT EmployeeID, FirstName, LastName, Email FROM Operator As O INNER JOIN MonthEmailCheck As M ON O.Id = M.OperatorID WHERE M.GotEmailed = 0 AND MONTH(M.Month) = 5 AND Seniority != -1 ORDER BY ExtraShifts, Seniority ASC"
        ' After first iteration, include outside operators.

        ' When query returns 0. Select MonthEmailCheck (MEC) records that correlate to the month and reset GotEmailed to 0. Create MEC records for Outside operators if they don't exist
    End Sub

    Private Sub HandleUnreadEmails()
        ' Reads unread emails.
        Logger("Received request to read inbox.")
        Dim inboxItems As Outlook.Items = coHandler.readEmails(oApp)
        Dim i As Integer
        Dim oMsg As Outlook.MailItem
        Dim msgSubject As String
        Dim consoleMsg As String = ""
        For i = 1 To inboxItems.Count
            oMsg = inboxItems.Item(i)
            msgSubject = oMsg.Subject
            consoleMsg += String.Format("
                Sender Email: {0} ({1})
                Time Sent: {2}
                Subject: {3}
                Message: {4}
                *********************", coHandler.getUsername(oMsg), coHandler.getEmailAddress(oMsg), oMsg.ReceivedTime, msgSubject, oMsg.Body)

            ' Switch statement to read msgSubject and process request
            ' Marks message as Read to avoid being read again. 
            'oMsg.UnRead = False
        Next
        ConsoleAdd(consoleMsg)

        ' If it contains "Open Availability Selection" and is prefaced with an id number or "pass" handle appropriately.
        ' Pass: Print in console, mark MEC, done.
        ' ID: Update availability with user EmployeeID, 
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
