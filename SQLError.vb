Public Class SQLError
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        My.Settings.SQLConString = TextBox5.Text
        My.Settings.Save()
        TextBox5.Text = My.Settings.SQLConString
        ' Display a message box asking the user if they want to restart the application
        MsgBox("Application will now restart!", MsgBoxStyle.OkOnly)
        RestartApplication()
    End Sub

    Private Sub SQLError_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim SqlConString As String = My.Settings.SQLConString
        TextBox5.Text = SqlConString
    End Sub
    Private Sub RestartApplication()
        Dim applicationPath As String = Application.ExecutablePath
        Dim processInfo As ProcessStartInfo = New ProcessStartInfo(applicationPath)
        Process.Start(processInfo)
        Application.Exit()
    End Sub
End Class