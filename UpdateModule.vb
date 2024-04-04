Imports Microsoft.Data.SqlClient
Imports System.Reflection
Imports System.Data.SqlClient

Module UpdateModule
    Private Function GetLatestVersionFromDatabase() As Version
        ' Connection string to your SQL Server database
        Dim connectionString As String = SQLCon.connectionString

        ' Query to retrieve the latest version number from your database
        Dim query As String = "SELECT VersionNumber FROM ElectrotechnologyReports.dbo.Updates WHERE Id = 1" ' Adjust as needed

        Try
            ' Open a connection to the database
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    ' Open the connection
                    connection.Open()

                    ' Execute the query and retrieve the latest version number
                    Dim latestVersionString As String = command.ExecuteScalar()?.ToString()

                    ' Parse the version number string to a Version object
                    Dim latestVersion As New Version(latestVersionString)

                    ' Return the latest version number
                    Return latestVersion
                End Using
            End Using
        Catch ex As Exception
            ' Handle any exceptions (e.g., database connection error)
            ' You can log the error or display an error message to the user
            Return Nothing
        End Try
    End Function

    Public Function IsUpdateAvailable() As Boolean
        ' Get the latest version number from the database
        Dim latestVersion As Version = GetLatestVersionFromDatabase()

        If latestVersion IsNot Nothing Then
            ' Get the installed version number
            Dim installedVersion As Version = Assembly.GetExecutingAssembly().GetName().Version

            ' Compare the installed version with the latest version
            If latestVersion > installedVersion Then
                ' An update is available
                Return True
            End If
        End If

        ' No update available
        Return False
    End Function

End Module
