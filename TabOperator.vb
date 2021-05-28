Imports Microsoft.VisualBasic
Imports System.Data.SqlClient

Partial Public Class OperatorMainForm
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
End Class
