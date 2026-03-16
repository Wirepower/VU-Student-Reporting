Imports Microsoft.Data.SqlClient
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Reflection
Imports System.Text.Json

Public Class GitHubReleaseInfo
    Public Property TagName As String
    Public Property Name As String
    Public Property HtmlUrl As String
    Public Property PublishedAt As DateTime?
End Class

Public Class GitHubUpdateCheckResult
    Public Property IsSuccessful As Boolean
    Public Property IsUpdateAvailable As Boolean
    Public Property CurrentTag As String
    Public Property LatestRelease As GitHubReleaseInfo
    Public Property ErrorMessage As String
End Class

Module UpdateModule
    Private Const GitHubOwner As String = "Wirepower"
    Private Const GitHubRepo As String = "VU-Student-Reporting"
    Public Const CurrentReleaseTag As String = "v2.0sql"

    Private Function GetLatestVersionFromDatabase() As Version
        Dim connectionString As String = SQLCon.connectionString
        Dim query As String = "SELECT VersionNumber FROM ElectrotechnologyReports.dbo.Updates WHERE Id = 1"

        Try
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    connection.Open()
                    Dim latestVersionString As String = command.ExecuteScalar()?.ToString()

                    If String.IsNullOrWhiteSpace(latestVersionString) Then
                        Return Nothing
                    End If

                    Return New Version(latestVersionString)
                End Using
            End Using
        Catch
            Return Nothing
        End Try
    End Function

    ' Legacy SQL-based update check kept for backward compatibility/fallback.
    Public Function IsUpdateAvailable() As Boolean
        Dim latestVersion As Version = GetLatestVersionFromDatabase()

        If latestVersion IsNot Nothing Then
            Dim installedVersion As Version = Assembly.GetExecutingAssembly().GetName().Version
            Return latestVersion > installedVersion
        End If

        Return False
    End Function

    Public Async Function CheckForGitHubUpdateAsync() As Task(Of GitHubUpdateCheckResult)
        Dim apiUrl As String = $"https://api.github.com/repos/{GitHubOwner}/{GitHubRepo}/releases/latest"

        Try
            Using client As New HttpClient()
                client.DefaultRequestHeaders.UserAgent.Clear()
                client.DefaultRequestHeaders.UserAgent.Add(New ProductInfoHeaderValue("StudentAttendanceReporting", "1.0"))
                client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/vnd.github+json"))

                Using response As HttpResponseMessage = Await client.GetAsync(apiUrl).ConfigureAwait(False)
                    If Not response.IsSuccessStatusCode Then
                        Return New GitHubUpdateCheckResult With {
                            .IsSuccessful = False,
                            .CurrentTag = CurrentReleaseTag,
                            .ErrorMessage = $"GitHub check failed ({CInt(response.StatusCode)} {response.ReasonPhrase})."
                        }
                    End If

                    Dim jsonPayload As String = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                    Using document As JsonDocument = JsonDocument.Parse(jsonPayload)
                        Dim root As JsonElement = document.RootElement

                        Dim latestTag As String = GetString(root, "tag_name")
                        Dim releaseName As String = GetString(root, "name")
                        Dim htmlUrl As String = GetString(root, "html_url")
                        Dim publishedAtValue As DateTime? = Nothing
                        Dim publishedAtRaw As String = GetString(root, "published_at")

                        Dim parsedDate As DateTime
                        If DateTime.TryParse(publishedAtRaw, parsedDate) Then
                            publishedAtValue = parsedDate
                        End If

                        Dim latestRelease As New GitHubReleaseInfo With {
                            .TagName = latestTag,
                            .Name = releaseName,
                            .HtmlUrl = htmlUrl,
                            .PublishedAt = publishedAtValue
                        }

                        Return New GitHubUpdateCheckResult With {
                            .IsSuccessful = True,
                            .IsUpdateAvailable = IsNewerReleaseTag(CurrentReleaseTag, latestTag),
                            .CurrentTag = CurrentReleaseTag,
                            .LatestRelease = latestRelease
                        }
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Return New GitHubUpdateCheckResult With {
                .IsSuccessful = False,
                .CurrentTag = CurrentReleaseTag,
                .ErrorMessage = $"GitHub check failed: {ex.Message}"
            }
        End Try
    End Function

    Private Function GetString(root As JsonElement, propertyName As String) As String
        Dim prop As JsonElement
        If root.TryGetProperty(propertyName, prop) Then
            If prop.ValueKind = JsonValueKind.String Then
                Return prop.GetString()
            End If
        End If

        Return String.Empty
    End Function

    Private Function IsNewerReleaseTag(currentTag As String, latestTag As String) As Boolean
        If String.IsNullOrWhiteSpace(latestTag) Then
            Return False
        End If

        Return Not String.Equals(currentTag.Trim(), latestTag.Trim(), StringComparison.OrdinalIgnoreCase)
    End Function
End Module
