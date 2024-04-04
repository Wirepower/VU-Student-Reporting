Imports Microsoft.Data.SqlClient

Public Class EmailTemplates
    Private Sub EmailTemplates_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PopulateEmailSubjectComboBox()
    End Sub
    Private Sub PopulateEmailSubjectComboBox()
        ' Clear existing items in the ComboBox
        ComboBox1.Items.Clear()

        ' Construct SQL SELECT command to retrieve distinct email subjects
        Dim selectCommand As String = "SELECT DISTINCT EmailSubject FROM ElectrotechnologyReports.dbo.EmailTemplates"

        ' Create SqlConnection and SqlCommand objects
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(selectCommand, connection)
                ' Open connection
                connection.Open()

                ' Execute SQL command and read data
                Using reader As SqlDataReader = command.ExecuteReader()
                    ' Loop through the result set
                    While reader.Read()
                        ' Get email subject
                        Dim emailSubject As String = reader("EmailSubject").ToString()

                        ' Add email subject to ComboBox
                        ComboBox1.Items.Add(emailSubject)
                    End While
                End Using
            End Using
        End Using
    End Sub
    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        Dim selectedEmailSubject As String = ComboBox1.Text

        ' Check if the selected email subject exists in the ComboBox
        If ComboBox1.Items.Contains(selectedEmailSubject) Then
            ' Email subject exists, so hide the ADD button and show the UPDATE and DELETE buttons
            Button2.Visible = False 'ADD button
            Button1.Visible = True ' Update Button
            Button3.Visible = True 'Delete Button
        Else
            ' Email subject doesn't exist, so hide all three buttons
            Button2.Visible = True 'ADD button
            Button1.Visible = False ' Update Button
            Button3.Visible = False 'Delete Button
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ' Get the selected email subject
        Dim selectedEmailSubject As String = ComboBox1.SelectedItem.ToString()

        ' Construct SQL SELECT command to retrieve the EmailBody based on the selected email subject
        Dim selectCommand As String = "SELECT EmailBody FROM ElectrotechnologyReports.dbo.EmailTemplates WHERE EmailSubject = @EmailSubject"

        Try
            ' Create SqlConnection and SqlCommand objects
            Using connection As New SqlConnection(SQLCon.connectionString)
                Using command As New SqlCommand(selectCommand, connection)
                    ' Add parameter for the selected email subject
                    command.Parameters.AddWithValue("@EmailSubject", selectedEmailSubject)

                    ' Open connection
                    connection.Open()

                    ' Execute SQL command and read data
                    Using reader As SqlDataReader = command.ExecuteReader()
                        ' Check if there is data available
                        If reader.Read() Then
                            ' Populate TextBox1 with the EmailBody
                            TextBox1.Text = reader("EmailBody").ToString()
                        Else
                            ' If no data found, clear TextBox1
                            TextBox1.Text = ""
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' Handle any errors
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


        ' Check if the selected email subject exists in the ComboBox
        If ComboBox1.Items.Contains(selectedEmailSubject) Then
            ' Email subject exists, so hide the ADD button and show the UPDATE button
            Button2.Visible = False
            Button1.Visible = True
            Button3.Visible = True
        Else
            ' Email subject doesn't exist, so hide both the ADD and UPDATE buttons
            Button2.Visible = False
            Button1.Visible = False
            Button3.Visible = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Get the values from the TextBox controls
        Dim newEmailSubject As String = ComboBox1.Text
        Dim newEmailBody As String = TextBox1.Text
        ' Construct SQL INSERT command
        Dim insertCommand As String = "INSERT INTO ElectrotechnologyReports.dbo.EmailTemplates (EmailSubject, EmailBody) VALUES (@EmailSubject, @EmailBody)"
        Try
            ' Create SqlConnection and SqlCommand objects
            Using connection As New SqlConnection(SQLCon.connectionString)
                Using command As New SqlCommand(insertCommand, connection)
                    ' Add parameters for the new email subject and body
                    command.Parameters.AddWithValue("@EmailSubject", newEmailSubject)
                    command.Parameters.AddWithValue("@EmailBody", newEmailBody)
                    ' Open connection
                    connection.Open()
                    ' Execute SQL command
                    command.ExecuteNonQuery()
                    ' Refresh the ComboBox after adding a new record
                    PopulateEmailSubjectComboBox()
                End Using
            End Using
            MessageBox.Show("New email template added successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            ' Handle any errors
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Get the values from the TextBox controls
        Dim updatedEmailSubject As String = ComboBox1.Text
        Dim updatedEmailBody As String = TextBox1.Text
        ' Get the selected email subject
        Dim selectedEmailSubject As String = ComboBox1.SelectedItem.ToString()
        ' Construct SQL UPDATE command
        Dim updateCommand As String = "UPDATE ElectrotechnologyReports.dbo.EmailTemplates SET EmailBody = @UpdatedEmailBody WHERE EmailSubject = @SelectedEmailSubject"
        Try
            ' Create SqlConnection and SqlCommand objects
            Using connection As New SqlConnection(SQLCon.connectionString)
                Using command As New SqlCommand(updateCommand, connection)
                    ' Add parameters for the updated email body and selected email subject
                    command.Parameters.AddWithValue("@UpdatedEmailBody", updatedEmailBody)
                    command.Parameters.AddWithValue("@SelectedEmailSubject", selectedEmailSubject)
                    ' Open connection
                    connection.Open()
                    ' Execute SQL command
                    command.ExecuteNonQuery()
                End Using
            End Using
            MessageBox.Show("Email template updated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            ' Handle any errors
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Get the selected email subject
        Dim selectedEmailSubject As String = ComboBox1.SelectedItem.ToString()
        ' Confirmation dialog before deleting
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete the selected email template?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            ' Construct SQL DELETE command
            Dim deleteCommand As String = "DELETE FROM ElectrotechnologyReports.dbo.EmailTemplates WHERE EmailSubject = @EmailSubject"
            Try
                ' Create SqlConnection and SqlCommand objects
                Using connection As New SqlConnection(SQLCon.connectionString)
                    Using command As New SqlCommand(deleteCommand, connection)
                        ' Add parameter for the selected email subject
                        command.Parameters.AddWithValue("@EmailSubject", selectedEmailSubject)
                        ' Open connection
                        connection.Open()
                        ' Execute SQL command
                        command.ExecuteNonQuery()
                        ' Refresh the ComboBox and TextBox after deleting
                        PopulateEmailSubjectComboBox()
                        TextBox1.Clear()
                        ComboBox1.Text = ""
                    End Using
                End Using
                MessageBox.Show("Email template deleted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                ' Handle any errors
                MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ComboBox1.Text = ""  ' Clear ComboBox text first
        TextBox1.Text = ""   ' Clear TextBox text second

        ' Show all buttons
        Button4.Visible = True
        Button3.Visible = True
        Button2.Visible = True
        Button1.Visible = True
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub
End Class