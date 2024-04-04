Imports Microsoft.Data.SqlClient

Public Class LastReportDatePicker
    Private Sub UpdateLastReportDate(studentID As String, reportDate As Date)
        Dim connectionString As String = SQLCon.connectionString
        Dim query As String = "UPDATE ElectrotechnologyReports.dbo.StudentLogs " &
                              "SET LastStudentReportDate = @ReportDate " &
                              "WHERE [Student ID] = @StudentID"

        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@ReportDate", reportDate)
                command.Parameters.AddWithValue("@StudentID", studentID)

                Try
                    connection.Open()
                    Dim rowsAffected As Integer = command.ExecuteNonQuery()
                    If rowsAffected > 0 Then
                        'MessageBox.Show("Last student report date updated successfully.")
                    Else
                        MessageBox.Show("Student ID not found.")
                    End If
                Catch ex As Exception
                    MessageBox.Show("Error updating last report date: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub


    ' Event handler for the OK button click
    Private Sub OkButton_Click(sender As Object, e As EventArgs) Handles OKbutton.Click
        Dim studentID As String = MainFrm.StudentIDLBL.Text
        Dim reportDate As Date = DateTimePicker1.Value

        UpdateLastReportDate(studentID, reportDate)


        ' Close the form
        Me.Close()
    End Sub

    ' Event handler for the Cancel button click
    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
        Me.Close()
    End Sub

    Private Sub LastReportDatePicker_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "It seems this student has never had a previous Student Report before, " & vbCrLf &
            " Please Select the Date When this log started. " & vbCrLf & " 
(This is usually the start Of a term date and/or the First date of class the student was enrolled in)" & vbCrLf & "
This only needs to be set once."
    End Sub
End Class