Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Globalization
Imports Microsoft.Data.SqlClient

Module SQLCon
    'Public connectionString As String = "Server=DEVSQLCENTRAL.AD.VU.EDU.AU;Integrated Security=True;Connect Timeout=30;Encrypt=True;Trust Server Certificate=True;Application Intent=ReadWrite;Multi Subnet Failover=False;"
    Public connectionString As String = My.Settings.SQLConString

    ''' <summary>Normalizes student ids from CSV/UI (Excel may emit 4608758.0; must match [Student ID] in StudentLogs / staging).</summary>
    Public Function NormalizeStudentIdForLogs(raw As Object) As String
        If raw Is Nothing OrElse raw Is DBNull.Value Then Return String.Empty
        Dim t As String = Convert.ToString(raw).Trim()
        If t.Length > 0 AndAlso t(0) = ChrW(&HFEFF) Then t = t.Substring(1).Trim()
        Dim d As Double
        If Double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, d) Then
            Return Math.Truncate(d).ToString("0", CultureInfo.InvariantCulture)
        End If
        Return t
    End Function

    Private statusLabel As Label ' Declare a private variable to hold the reference to the status label

    Public Sub InitializeStatusLabel(statusLbl As Label)
        statusLabel = statusLbl ' Assign the label passed from the form to the statusLabel variable
    End Sub

    Public Function GetConnection() As SqlConnection
        Try
            Dim connection As New SqlConnection(connectionString)
            Return connection
        Catch ex As Exception
            MsgBox("Error creating connection: " & vbCrLf & ex.Message & vbCrLf & "Exiting Application!")

            'Environment.Exit(0)
            Return Nothing ' Return null or Nothing to indicate failure
        End Try
    End Function

    Public Sub OpenConnection(ByRef connection As SqlConnection)
        Try
            If connection.State <> ConnectionState.Open Then
                connection.Open()
            End If
            UpdateStatusLabel("Connected", Color.Green)
        Catch ex As Exception
            MsgBox("Error opening connection: " & vbCrLf & ex.Message & vbCrLf & "Exiting Application!")
            UpdateStatusLabel("Error opening connection.", Color.Red)

            'Environment.Exit(0)
        End Try
    End Sub

    Public Sub CloseConnection(ByRef connection As SqlConnection)
        Try
            If connection.State = ConnectionState.Open Then
                connection.Close()
                UpdateStatusLabel("Connection closed", Color.Yellow)
            End If
        Catch ex As Exception
            MsgBox("Error closing connection: " & vbCrLf & ex.Message)
            UpdateStatusLabel("Error closing connection.", Color.Red)
        End Try
    End Sub

    Private Sub UpdateStatusLabel(text As String, color As Color)
        If statusLabel IsNot Nothing Then
            statusLabel.Text = text
            statusLabel.ForeColor = color
        End If
    End Sub
End Module

