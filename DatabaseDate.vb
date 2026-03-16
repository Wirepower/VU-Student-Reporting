Imports Microsoft.Data.SqlClient

Module DatabaseDate
    Public Sub UpdateStudentDatabaseLabel()
        ' Database connection string
        Dim connectionString As String = SQLCon.connectionString

        ' SQL query to retrieve the date from the table
        Dim query As String = "SELECT StudentDatabaseDate FROM ElectrotechnologyReports.dbo.Updates WHERE ID = 1;" ' Assuming the ID of the row you want to retrieve is 1

        Try
            ' Create a SqlConnection and SqlCommand objects
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    ' Open the connection
                    connection.Open()

                    ' Execute the SQL command and get the result
                    Dim result As Object = command.ExecuteScalar()

                    ' Check if the result is DBNull.Value
                    If result IsNot DBNull.Value Then
                        ' If not DBNull.Value, display the date in Label37/Label14 and unhide Label36/Label13
                        MainFrm.Label37.Text = CType(result, DateTime).ToString("dd/MM/yyyy")
                        MainFrm.Label36.Visible = True
                        MainFrm.Label37.Visible = True
                        SettingsForm.Label14.Text = CType(result, DateTime).ToString("dd/MM/yyyy")
                        SettingsForm.Label13.Visible = True
                        SettingsForm.Label14.Visible = True
                    Else
                        ' If DBNull.Value, hide Label36/Label13 and Label37/Label14
                        MainFrm.Label36.Visible = False
                        MainFrm.Label37.Visible = False
                        SettingsForm.Label13.Visible = False
                        SettingsForm.Label14.Visible = False
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' Handle exceptions
            MessageBox.Show("Error retrieving date: " & ex.Message)
        End Try
    End Sub
    Public Sub DatabaseStudentUnitDate()
        Try
            ' Construct the SQL query to retrieve the database update date
            Dim query As String = "SELECT DatabaseUpdateDate FROM ElectrotechnologyReports.dbo.Updates WHERE ID = 1"

            ' Create a new SqlConnection object using your connection string
            Using connection As New SqlConnection(SQLCon.connectionString)
                ' Create a new SqlCommand object with the query and connection
                Using command As New SqlCommand(query, connection)
                    ' Open the connection
                    connection.Open()

                    ' Execute the SQL query and get the result
                    Dim result As Object = Command.ExecuteScalar()

                    ' Check if the result is not null
                    If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                        ' Convert the result to DateTime
                        Dim databaseUpdateDate As DateTime = Convert.ToDateTime(result)


                        ' Set the label's text property with the database update date formatted as "dd/MM/yyyy"
                        SettingsForm.Label16.Text = "Database Current as of: " & databaseUpdateDate.ToString("dd/MM/yyyy")

                    Else
                        ' If the result is null or DBNull, display a message indicating no date is available
                        SettingsForm.Label16.Text = "Database Update Date Not Available"
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' Handle any errors
            MessageBox.Show("Error retrieving database update date: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub UpdateProfilingDate()
        ' Database connection string
        Dim connectionString As String = SQLCon.connectionString

        ' SQL query to retrieve the date from the table
        Dim query As String = "SELECT ProfileDBDate FROM ElectrotechnologyReports.dbo.Updates WHERE ID = 1;" ' Assuming the ID of the row you want to retrieve is 1

        Try
            ' Create a SqlConnection and SqlCommand objects
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    ' Open the connection
                    connection.Open()

                    ' Execute the SQL command and get the result
                    Dim result As Object = command.ExecuteScalar()

                    ' Check if the result is DBNull.Value
                    If result IsNot DBNull.Value Then
                        ' If not DBNull.Value, display the date in Label38/Label19 
                        MainFrm.Label39.Text = CType(result, DateTime).ToString("dd/MM/yyyy")
                        MainFrm.Label38.Visible = True
                        MainFrm.Label39.Visible = True
                    Else
                        ' If DBNull.Value, hide Label38/Label19 
                        MainFrm.Label38.Visible = False
                        MainFrm.Label39.Visible = False
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' Handle exceptions
            MessageBox.Show("Error retrieving date: " & ex.Message)
        End Try
    End Sub
End Module
