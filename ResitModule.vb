Imports Microsoft.Data.SqlClient

Module ResitModule
    Public Sub CheckResit(ByVal studentID As String, ByVal resitLabel As Label)
        Dim connectionString As String = SQLCon.connectionString
        Dim query As String = "SELECT Unit, [Resit date] FROM ElectrotechnologyReports.dbo.ElectricalResit WHERE [Student ID] = @studentID"

        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@studentID", studentID)
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.HasRows Then
                    reader.Read()
                    Dim unit As String = reader.GetString(0)
                    Dim resitDate As DateTime = reader.GetDateTime(1)
                    Dim formattedDate As String = resitDate.ToString("dddd, dd MMMM, yyyy")
                    resitLabel.Text = $"Student booked for resit for {unit} on {formattedDate}"
                Else
                    resitLabel.Text = "" ' No resit found for this student
                End If
            End Using
        End Using
    End Sub
End Module
