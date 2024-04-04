Imports Microsoft.Data.SqlClient

Public Class EmailSubjectHelp
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim selectedSubject As String = ComboBox1.SelectedItem.ToString()
        Dim connectionString As String = SQLCon.connectionString
        Dim selectQuery As String = $"SELECT EmailHelp FROM ElectrotechnologyReports.dbo.EmailTemplates WHERE EmailSubject = @Subject"

        Using connection As New SqlConnection(connectionString)
            connection.Open()

            Using command As New SqlCommand(selectQuery, connection)
                command.Parameters.AddWithValue("@Subject", selectedSubject)
                Dim emailHelp As Object = command.ExecuteScalar()

                If emailHelp IsNot Nothing Then
                    Label2.Text = emailHelp.ToString()
                Else
                    Label2.Text = "No help available"
                End If
            End Using
        End Using
    End Sub
End Class