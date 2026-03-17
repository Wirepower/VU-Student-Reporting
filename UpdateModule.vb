Imports Microsoft.Data.SqlClient
Imports System.IO
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Reflection
Imports System.Text.Json
Imports System.Text.RegularExpressions

Public Class GitHubReleaseInfo
    Public Property TagName As String
    Public Property Name As String
    Public Property HtmlUrl As String
    Public Property PublishedAt As DateTime?
    Public Property Body As String
    Public Property Assets As New List(Of GitHubReleaseAssetInfo)
End Class

Public Class GitHubReleaseAssetInfo
    Public Property Name As String
    Public Property DownloadUrl As String
    Public Property SizeBytes As Long
End Class

Public Class GitHubUpdateCheckResult
    Public Property IsSuccessful As Boolean
    Public Property IsUpdateAvailable As Boolean
    Public Property IsMandatory As Boolean
    Public Property MandatoryReason As String
    Public Property CurrentTag As String
    Public Property LatestRelease As GitHubReleaseInfo
    Public Property PreferredAsset As GitHubReleaseAssetInfo
    Public Property ErrorMessage As String
End Class

Public Class GitHubUpdateInstallResult
    Public Property IsSuccessful As Boolean
    Public Property DownloadedFilePath As String
    Public Property ErrorMessage As String
End Class

Friend Class OtaPolicy
    Public Property MinRequiredTag As String
    Public Property ForceUpdate As Boolean
    Public Property PreferredAssetName As String
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
                        Dim releaseBody As String = GetString(root, "body")
                        Dim assets As List(Of GitHubReleaseAssetInfo) = ParseAssets(root)

                        Dim parsedDate As DateTime
                        If DateTime.TryParse(publishedAtRaw, parsedDate) Then
                            publishedAtValue = parsedDate
                        End If

                        Dim policy As OtaPolicy = ParseOtaPolicy(releaseBody)

                        Dim latestRelease As New GitHubReleaseInfo With {
                            .TagName = latestTag,
                            .Name = releaseName,
                            .HtmlUrl = htmlUrl,
                            .PublishedAt = publishedAtValue,
                            .Body = releaseBody,
                            .Assets = assets
                        }

                        Dim preferredAsset As GitHubReleaseAssetInfo = SelectPreferredAsset(assets, policy.PreferredAssetName)
                        Dim updateAvailable As Boolean = IsNewerReleaseTag(CurrentReleaseTag, latestTag)
                        Dim mandatory As Boolean = IsMandatoryUpdate(CurrentReleaseTag, latestTag, updateAvailable, policy)
                        Dim mandatoryReason As String = GetMandatoryReason(mandatory, CurrentReleaseTag, latestTag, policy)

                        Return New GitHubUpdateCheckResult With {
                            .IsSuccessful = True,
                            .IsUpdateAvailable = updateAvailable,
                            .IsMandatory = mandatory,
                            .MandatoryReason = mandatoryReason,
                            .CurrentTag = CurrentReleaseTag,
                            .LatestRelease = latestRelease,
                            .PreferredAsset = preferredAsset
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

    Public Async Function DownloadAndLaunchGitHubUpdateAsync(checkResult As GitHubUpdateCheckResult) As Task(Of GitHubUpdateInstallResult)
        If checkResult Is Nothing OrElse checkResult.PreferredAsset Is Nothing Then
            Return New GitHubUpdateInstallResult With {
                .IsSuccessful = False,
                .ErrorMessage = "No downloadable release asset is available for this update."
            }
        End If

        If String.IsNullOrWhiteSpace(checkResult.PreferredAsset.DownloadUrl) Then
            Return New GitHubUpdateInstallResult With {
                .IsSuccessful = False,
                .ErrorMessage = "Release asset URL is missing."
            }
        End If

        Try
            Dim assetName As String = Path.GetFileName(checkResult.PreferredAsset.Name)
            If String.IsNullOrWhiteSpace(assetName) Then
                assetName = "StudentAttendanceReporting-update.bin"
            End If

            Dim targetFolder As String = Path.Combine(
                Path.GetTempPath(),
                "StudentAttendanceReporting",
                "updates",
                MakeSafePathPart(checkResult.LatestRelease?.TagName)
            )
            Directory.CreateDirectory(targetFolder)

            Dim targetFile As String = Path.Combine(targetFolder, assetName)
            Dim shouldDownload As Boolean = True

            If File.Exists(targetFile) AndAlso checkResult.PreferredAsset.SizeBytes > 0 Then
                Dim existingLength As Long = New FileInfo(targetFile).Length
                shouldDownload = (existingLength <> checkResult.PreferredAsset.SizeBytes)
            End If

            If shouldDownload Then
                Using client As New HttpClient()
                    client.DefaultRequestHeaders.UserAgent.Clear()
                    client.DefaultRequestHeaders.UserAgent.Add(New ProductInfoHeaderValue("StudentAttendanceReporting", "1.0"))
                    client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/octet-stream"))

                    Using response As HttpResponseMessage = Await client.GetAsync(checkResult.PreferredAsset.DownloadUrl, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)
                        response.EnsureSuccessStatusCode()
                        Using sourceStream As Stream = Await response.Content.ReadAsStreamAsync().ConfigureAwait(False)
                            Using destinationStream As New FileStream(targetFile, FileMode.Create, FileAccess.Write, FileShare.None)
                                Await sourceStream.CopyToAsync(destinationStream).ConfigureAwait(False)
                            End Using
                        End Using
                    End Using
                End Using
            End If

            Dim startInfo As New ProcessStartInfo With {
                .FileName = targetFile,
                .UseShellExecute = True
            }
            Process.Start(startInfo)

            Return New GitHubUpdateInstallResult With {
                .IsSuccessful = True,
                .DownloadedFilePath = targetFile
            }
        Catch ex As Exception
            Return New GitHubUpdateInstallResult With {
                .IsSuccessful = False,
                .ErrorMessage = ex.Message
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

    Private Function ParseAssets(root As JsonElement) As List(Of GitHubReleaseAssetInfo)
        Dim assets As New List(Of GitHubReleaseAssetInfo)()

        Dim assetsElement As JsonElement
        If root.TryGetProperty("assets", assetsElement) AndAlso assetsElement.ValueKind = JsonValueKind.Array Then
            For Each assetElement As JsonElement In assetsElement.EnumerateArray()
                Dim asset As New GitHubReleaseAssetInfo With {
                    .Name = GetString(assetElement, "name"),
                    .DownloadUrl = GetString(assetElement, "browser_download_url"),
                    .SizeBytes = GetLong(assetElement, "size")
                }

                If Not String.IsNullOrWhiteSpace(asset.DownloadUrl) Then
                    assets.Add(asset)
                End If
            Next
        End If

        Return assets
    End Function

    Private Function GetLong(root As JsonElement, propertyName As String) As Long
        Dim prop As JsonElement
        If root.TryGetProperty(propertyName, prop) AndAlso prop.ValueKind = JsonValueKind.Number Then
            Dim result As Long
            If prop.TryGetInt64(result) Then
                Return result
            End If
        End If

        Return 0
    End Function

    Private Function ParseOtaPolicy(releaseBody As String) As OtaPolicy
        Dim policy As New OtaPolicy()

        If String.IsNullOrWhiteSpace(releaseBody) Then
            Return policy
        End If

        For Each line As String In releaseBody.Split({ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
            Dim trimmed As String = line.Trim()
            If String.IsNullOrWhiteSpace(trimmed) Then
                Continue For
            End If

            If trimmed.StartsWith("#") OrElse trimmed.StartsWith("//") Then
                Continue For
            End If

            Dim key As String = ""
            Dim value As String = ""
            Dim separatorIndex As Integer = trimmed.IndexOf("="c)

            If separatorIndex < 0 Then
                separatorIndex = trimmed.IndexOf(":"c)
            End If

            If separatorIndex > 0 Then
                key = trimmed.Substring(0, separatorIndex).Trim().ToLowerInvariant()
                value = trimmed.Substring(separatorIndex + 1).Trim()
            End If

            If String.IsNullOrEmpty(key) Then
                Continue For
            End If

            Select Case key
                Case "min_required_tag", "min-required-tag", "minimum_required_tag"
                    policy.MinRequiredTag = value
                Case "force_update", "force-update", "mandatory_update", "mandatory-update"
                    policy.ForceUpdate = IsTrueValue(value)
                Case "asset_name", "asset-name", "preferred_asset"
                    policy.PreferredAssetName = value
            End Select
        Next

        Return policy
    End Function

    Private Function IsTrueValue(value As String) As Boolean
        If String.IsNullOrWhiteSpace(value) Then
            Return False
        End If

        Select Case value.Trim().ToLowerInvariant()
            Case "true", "1", "yes", "y"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Function SelectPreferredAsset(assets As List(Of GitHubReleaseAssetInfo), configuredAssetName As String) As GitHubReleaseAssetInfo
        If assets Is Nothing OrElse assets.Count = 0 Then
            Return Nothing
        End If

        If Not String.IsNullOrWhiteSpace(configuredAssetName) Then
            For Each asset In assets
                If String.Equals(asset.Name, configuredAssetName.Trim(), StringComparison.OrdinalIgnoreCase) Then
                    Return asset
                End If
            Next
        End If

        Dim extensionPriority As String() = {".msi", ".msixbundle", ".msix", ".exe", ".zip"}
        For Each extension In extensionPriority
            For Each asset In assets
                If asset.Name IsNot Nothing AndAlso asset.Name.EndsWith(extension, StringComparison.OrdinalIgnoreCase) Then
                    Return asset
                End If
            Next
        Next

        Return assets(0)
    End Function

    Private Function IsMandatoryUpdate(currentTag As String, latestTag As String, updateAvailable As Boolean, policy As OtaPolicy) As Boolean
        If policy Is Nothing Then
            Return False
        End If

        If Not String.IsNullOrWhiteSpace(policy.MinRequiredTag) Then
            Dim comparison As Integer
            If TryCompareReleaseTags(currentTag, policy.MinRequiredTag, comparison) Then
                If comparison < 0 Then
                    Return True
                End If
            ElseIf Not String.Equals(currentTag, policy.MinRequiredTag, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        End If

        Return policy.ForceUpdate AndAlso updateAvailable
    End Function

    Private Function GetMandatoryReason(isMandatory As Boolean, currentTag As String, latestTag As String, policy As OtaPolicy) As String
        If Not isMandatory Then
            Return String.Empty
        End If

        If policy IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(policy.MinRequiredTag) Then
            Return $"Current release {currentTag} is below minimum required {policy.MinRequiredTag}."
        End If

        Return $"Release {latestTag} is marked as mandatory."
    End Function

    Private Function IsNewerReleaseTag(currentTag As String, latestTag As String) As Boolean
        If String.IsNullOrWhiteSpace(latestTag) Then
            Return False
        End If

        If String.Equals(currentTag.Trim(), latestTag.Trim(), StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        Dim comparison As Integer
        If TryCompareReleaseTags(currentTag, latestTag, comparison) Then
            Return comparison < 0
        End If

        Return True
    End Function

    Private Function TryCompareReleaseTags(leftTag As String, rightTag As String, ByRef comparison As Integer) As Boolean
        comparison = 0
        Dim leftVersion As Version = ExtractNumericVersion(leftTag)
        Dim rightVersion As Version = ExtractNumericVersion(rightTag)

        If leftVersion Is Nothing OrElse rightVersion Is Nothing Then
            Return False
        End If

        comparison = leftVersion.CompareTo(rightVersion)
        Return True
    End Function

    Private Function ExtractNumericVersion(tag As String) As Version
        If String.IsNullOrWhiteSpace(tag) Then
            Return Nothing
        End If

        Dim match As Match = Regex.Match(tag, "(\d+(?:\.\d+)*)")
        If Not match.Success Then
            Return Nothing
        End If

        Dim rawParts() As String = match.Groups(1).Value.Split("."c)
        Dim normalizedParts As New List(Of String)()

        For Each part In rawParts
            Dim parsed As Integer
            If Integer.TryParse(part, parsed) Then
                normalizedParts.Add(parsed.ToString())
            End If
        Next

        If normalizedParts.Count = 0 Then
            Return Nothing
        End If

        If normalizedParts.Count = 1 Then
            normalizedParts.Add("0")
        End If

        While normalizedParts.Count > 4
            normalizedParts.RemoveAt(normalizedParts.Count - 1)
        End While

        Try
            Return New Version(String.Join(".", normalizedParts))
        Catch
            Return Nothing
        End Try
    End Function

    Private Function MakeSafePathPart(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then
            Return "unknown"
        End If

        Dim safe As String = value
        For Each invalidChar As Char In Path.GetInvalidFileNameChars()
            safe = safe.Replace(invalidChar, "_"c)
        Next

        If String.IsNullOrWhiteSpace(safe) Then
            Return "unknown"
        End If

        Return safe
    End Function
End Module
