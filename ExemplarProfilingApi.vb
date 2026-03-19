Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Reflection
Imports System.Text
Imports System.Text.Json

Public Class ExemplarProfileLookupResult
    Public Property IsSuccessful As Boolean
    Public Property IsConfigured As Boolean
    Public Property UserId As String
    Public Property StatusText As String
    Public Property DetailText As String
    Public Property TotalCards As Integer?
    Public Property CompletedCards As Integer?
    Public Property PendingCards As Integer?
    Public Property ErrorMessage As String
    ''' <summary>For colored UI: Cards not submitted/Outstanding count.</summary>
    Public Property MissingWeeks As Integer?
    ''' <summary>For colored UI: Cards Submitted (Not verified) count.</summary>
    Public Property SubmittedNotVerified As Integer?
    ''' <summary>For colored UI: Cards submitted (Employer Verified) count.</summary>
    Public Property SubmittedEmployerVerified As Integer?
    ''' <summary>For colored UI: Last card date as dd/MM/yyyy or empty.</summary>
    Public Property LastCardFormatted As String
    Public Property RawCardsSummaryJson As String
End Class

Public Class ExemplarQualificationUpdateResult
    Public Property IsSuccessful As Boolean
    Public Property ErrorMessage As String
    Public Property ResponseBody As String
End Class

Public Class ExemplarUnitProgressItem
    Public Property UnitCode As String
    Public Property Percentage As Decimal?
    Public Property Status As String
    Public Property TotalCards As Integer?
    Public Property CompletedCards As Integer?
    ' Raw JSON for the per-unit experience cards list from:
    ' /api/v1/users/:id/qualifications/:qualificationId/units/:unitCode/progression/cards
    Public Property ExperienceCardsRawJson As String
    ' Raw JSON for the per-unit demonstration cards list from the same endpoint.
    Public Property DemonstrationCardsRawJson As String
    ' Raw JSON payload from /progression/cards for this unit (debugging / inspection).
    Public Property ProgressionCardsEndpointRawJson As String

    ' Percentages as calculated by the qualification payload:
    ' calculation.demonstration_progression.unit_demonstration_statuses[*].percentage
    ' and (if present) calculation.experience_progression.unit_experience_statuses[*].percentage
    Public Property ExperienceCardsPercentage As Decimal?
    Public Property DemonstrationCardsPercentage As Decimal?
End Class

Public Class ExemplarUnitProgressResult
    Public Property IsSuccessful As Boolean
    Public Property IsConfigured As Boolean
    Public Property UserId As String
    Public Property QualificationId As String
    Public Property Units As New Dictionary(Of String, ExemplarUnitProgressItem)(StringComparer.OrdinalIgnoreCase)
    Public Property ErrorMessage As String
    Public Property RawApiJson As String
    Public Property RawQualificationsCatalogJson As String
    Public Property RawQualificationsJson As String
End Class

Public Class ExemplarQualificationsCatalogResult
    Public Property IsSuccessful As Boolean
    Public Property IsConfigured As Boolean
    Public Property ErrorMessage As String
    Public Property RawJson As String
End Class

Friend Class CardCountsResult
    Public Property Success As Boolean
    Public Property TotalCards As Integer
    Public Property CompletedCards As Integer
    Public Property ExperienceCardsRawJson As String
    Public Property DemonstrationCardsRawJson As String
    Public Property ProgressionCardsEndpointRawJson As String
End Class

Module ExemplarProfilingApi
    ''' <summary>Production API. Use the production login JAR with this URL. For staging, set EXEMPLAR_API_BASE_URL and use the staging JAR.</summary>
    Private Const DefaultBaseUrl As String = "https://api.profiling.exemplarsystems.com.au"
    Private ReadOnly ApiClient As New HttpClient() With {
        .Timeout = TimeSpan.FromSeconds(20)
    }
    Private CachedToken As String
    Private ReadOnly TokenRefreshLock As New Object()
    ''' <summary>Set when token refresh fails so the error message can be shown to the user.</summary>
    Private LastTokenRefreshError As String
    ''' <summary>Last token refresh debug info (non-secret).</summary>
    Private LastTokenRefreshBaseUrl As String
    Private LastTokenRefreshJarPath As String
    Private LastTokenRefreshUsernameUsed As String
    ''' <summary>When EXEMPLAR_DEBUG_LIST_RESPONSES=1, responses are collected here and written to ExemplarApiResponseFields.txt.</summary>
    Private DebugResponseList As New List(Of Tuple(Of String, String))

    ' Hardcoded Exemplar login: the app uses these to get a Bearer token (via the embedded or external JAR).
    ' Optionally, env vars EXEMPLAR_API_USERNAME / EXEMPLAR_API_PASSWORD override these if set (e.g. by IT).
    Private Const ExemplarApiUsername As String = "electrotechnology.admin@vu.edu.au"
    Private Const ExemplarApiPassword As String = "VUapi12345!"


    Private Class ExemplarUserCandidate
        Public Property Id As String
        Public Property FirstName As String
        Public Property LastName As String
        Public Property Email As String
    End Class

    Public Function IsConfigured() As Boolean
        Return Not String.IsNullOrWhiteSpace(GetBearerToken())
    End Function

    ''' <summary>When not configured, returns the reason (e.g. "Login JAR not found", "Java not found") so the UI can show it instead of the generic env-var message.</summary>
    Public Function GetNotConfiguredReason() As String
        If String.IsNullOrWhiteSpace(CachedToken) AndAlso String.IsNullOrWhiteSpace(LastTokenRefreshError) Then
            GetBearerToken()
        End If
        If Not String.IsNullOrWhiteSpace(CachedToken) Then
            Return ""
        End If
        If Not String.IsNullOrWhiteSpace(LastTokenRefreshError) Then
            Return LastTokenRefreshError
        End If
        Return "Set EXEMPLAR_API_TOKEN, or ensure ExemplarLogin.jar is in the app folder and Java is installed (or add a jre folder next to the app)."
    End Function

    Public Function GetConfiguredQualificationId() As String
        If Not String.IsNullOrWhiteSpace(My.Settings.ExemplarQualificationId) Then
            Return My.Settings.ExemplarQualificationId.Trim()
        End If

        Dim envQualification As String = Environment.GetEnvironmentVariable("EXEMPLAR_QUALIFICATION_ID")
        If String.IsNullOrWhiteSpace(envQualification) Then
            Return ""
        End If

        Return envQualification.Trim()
    End Function

    Public Async Function GetQualificationsCatalogAsync() As Task(Of ExemplarQualificationsCatalogResult)
        Dim token As String = GetBearerToken()
        If String.IsNullOrWhiteSpace(token) Then
            Return New ExemplarQualificationsCatalogResult With {
                .IsSuccessful = False,
                .IsConfigured = False,
                .ErrorMessage = "Profiling API credentials are not configured."
            }
        End If

        Try
            Dim url As String = $"{GetBaseUrl()}/api/v1/qualifications"
            Using doc As JsonDocument = Await SendJsonRequestAsync(HttpMethod.Get, url, token, Nothing)
                Return New ExemplarQualificationsCatalogResult With {
                    .IsSuccessful = True,
                    .IsConfigured = True,
                    .RawJson = doc.RootElement.GetRawText()
                }
            End Using
        Catch ex As Exception
            Return New ExemplarQualificationsCatalogResult With {
                .IsSuccessful = False,
                .IsConfigured = True,
                .ErrorMessage = ex.Message
            }
        End Try
    End Function

    Public Function GetQualificationIdSourceDescription() As String
        Dim settingValue As String = If(My.Settings.ExemplarQualificationId, "").Trim()
        If Not String.IsNullOrWhiteSpace(settingValue) Then
            Return "My.Settings.ExemplarQualificationId = " & settingValue
        End If

        Dim envQualification As String = If(Environment.GetEnvironmentVariable("EXEMPLAR_QUALIFICATION_ID"), "").Trim()
        If Not String.IsNullOrWhiteSpace(envQualification) Then
            Return "EXEMPLAR_QUALIFICATION_ID = " & envQualification
        End If

        Return "No qualification ID found in My.Settings.ExemplarQualificationId or EXEMPLAR_QUALIFICATION_ID."
    End Function

    Public Sub SetBearerToken(token As String)
        If String.IsNullOrWhiteSpace(token) Then
            CachedToken = Nothing
            Return
        End If

        Dim normalized As String = SanitizeTokenForHeader(token)
        If normalized.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase) Then
            normalized = normalized.Substring("Bearer ".Length).Trim()
        End If

        CachedToken = normalized
    End Sub

    ''' <summary>Removes newlines and trims so the token is safe for HTTP Authorization header.</summary>
    Private Function SanitizeTokenForHeader(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""
        Dim firstLine As String = value.Split(New Char() {vbCr, vbLf})(0)
        Return firstLine.Trim().Replace(Chr(13), "").Replace(Chr(10), "")
    End Function

    Public Async Function LookupStudentProfileAsync(firstName As String, lastName As String, email As String) As Task(Of ExemplarProfileLookupResult)
        Dim token As String = GetBearerToken()
        If String.IsNullOrWhiteSpace(token) Then
            Return New ExemplarProfileLookupResult With {
                .IsSuccessful = False,
                .IsConfigured = False,
                .StatusText = "Not configured",
                .DetailText = "Profiling API credentials are not configured."
            }
        End If

        Try
            If String.Equals(Environment.GetEnvironmentVariable("EXEMPLAR_DEBUG_LIST_RESPONSES"), "1", StringComparison.OrdinalIgnoreCase) Then
                DebugResponseList.Clear()
            End If
            Dim student As ExemplarUserCandidate = Await FindStudentAsync(token, firstName, lastName, email)
            If student Is Nothing Then
                Return New ExemplarProfileLookupResult With {
                    .IsSuccessful = False,
                    .IsConfigured = True,
                    .StatusText = "Student not found",
                    .DetailText = "No matching Exemplar student was found for the selected record."
                }
            End If

            Dim cardsUrl As String = $"{GetBaseUrl()}/api/v1/users/{Uri.EscapeDataString(student.Id)}/cards/summary"
            Using cardsDoc As JsonDocument = Await SendJsonRequestAsync(HttpMethod.Get, cardsUrl, token, Nothing)
                Dim rawCardsJson As String = cardsDoc.RootElement.GetRawText()
                ' Optional: set EXEMPLAR_DEBUG_SAVE_JSON=1 to write the raw cards/summary response to a file so you can see what the API returns and choose which fields to display.
                If String.Equals(Environment.GetEnvironmentVariable("EXEMPLAR_DEBUG_SAVE_JSON"), "1", StringComparison.OrdinalIgnoreCase) Then
                    Try
                        Dim debugPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExemplarCardsSummary_sample.json")
                        File.WriteAllText(debugPath, rawCardsJson, Encoding.UTF8)
                    Catch
                        ' Ignore
                    End Try
                End If
                ' Use the exact API paths you identified for the UI display
                Dim missingWeeks As Integer? = GetIntegerByPath(cardsDoc.RootElement, "card_submission.missing_weeks")  ' Cards not submitted/Outstanding
                Dim submittedNotVerified As Integer? = GetIntegerByPath(cardsDoc.RootElement, "card_status_counts[0].count")  ' Cards Submitted (Not verified)
                Dim submittedEmployerVerified As Integer? = GetIntegerByPath(cardsDoc.RootElement, "card_status_counts[1].count")  ' Cards submitted (Employer Verified)
                ' Fallback: API may return card_status_counts in different order; resolve by status name
                If (Not submittedNotVerified.HasValue OrElse Not submittedEmployerVerified.HasValue) Then
                    GetCardStatusCountsFromArray(cardsDoc.RootElement, submittedNotVerified, submittedEmployerVerified)
                End If

                Dim missingNum As String = If(missingWeeks.HasValue, missingWeeks.Value.ToString(), "?")
                Dim notVerifiedNum As String = If(submittedNotVerified.HasValue, submittedNotVerified.Value.ToString(), "?")
                Dim employerVerifiedNum As String = If(submittedEmployerVerified.HasValue, submittedEmployerVerified.Value.ToString(), "?")
                Dim detail As String = $"Cards not submitted/Outstanding: {missingNum} | Cards Submitted (Not verified): {notVerifiedNum} | Cards submitted (Employer Verified): {employerVerifiedNum}"

                ' Last card timestamp in dd/MM/yyyy format
                Dim lastCardFormatted As String = ""
                Dim lastCardTimestampRaw As String = GetStringByPath(cardsDoc.RootElement, "last_card_timestamp")
                If String.IsNullOrWhiteSpace(lastCardTimestampRaw) Then lastCardTimestampRaw = GetStringByPath(cardsDoc.RootElement, "card_submission.last_card_timestamp")
                If Not String.IsNullOrWhiteSpace(lastCardTimestampRaw) Then
                    Dim dt As DateTime
                    If DateTime.TryParse(lastCardTimestampRaw, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, dt) Then
                        lastCardFormatted = dt.ToString("dd/MM/yyyy")
                        detail &= " | Last card: " & lastCardFormatted
                    Else
                        detail &= " | Last card: " & lastCardTimestampRaw
                    End If
                End If

                ' Optional: EXEMPLAR_DEBUG_LIST_RESPONSES=1 writes API response fields to ExemplarApiResponseFields.txt
                If String.Equals(Environment.GetEnvironmentVariable("EXEMPLAR_DEBUG_LIST_RESPONSES"), "1", StringComparison.OrdinalIgnoreCase) AndAlso DebugResponseList.Count > 0 Then
                    Try
                        Dim sb As New StringBuilder()
                        For Each t As Tuple(Of String, String) In DebugResponseList
                            sb.AppendLine("========== API: " & t.Item1 & " ==========")
                            Try
                                Using doc As JsonDocument = JsonDocument.Parse(t.Item2)
                                    Dim lines As New List(Of String)()
                                    FlattenJsonToLines(doc.RootElement, "", lines)
                                    For Each line As String In lines
                                        sb.AppendLine(line)
                                    Next
                                End Using
                            Catch
                                sb.AppendLine("(raw): " & t.Item2)
                            End Try
                            sb.AppendLine()
                        Next
                        Dim outPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExemplarApiResponseFields.txt")
                        File.WriteAllText(outPath, sb.ToString(), Encoding.UTF8)
                    Catch
                        ' Ignore
                    End Try
                End If

                Dim totalCards As Integer? = If(missingWeeks.HasValue AndAlso submittedNotVerified.HasValue AndAlso submittedEmployerVerified.HasValue,
                    missingWeeks.Value + submittedNotVerified.Value + submittedEmployerVerified.Value, Nothing)
                Return New ExemplarProfileLookupResult With {
                    .IsSuccessful = True,
                    .IsConfigured = True,
                    .UserId = student.Id,
                    .StatusText = "Connected",
                    .DetailText = detail,
                    .TotalCards = totalCards,
                    .CompletedCards = If(submittedNotVerified.HasValue AndAlso submittedEmployerVerified.HasValue, submittedNotVerified.Value + submittedEmployerVerified.Value, Nothing),
                    .PendingCards = missingWeeks,
                    .MissingWeeks = missingWeeks,
                    .SubmittedNotVerified = submittedNotVerified,
                    .SubmittedEmployerVerified = submittedEmployerVerified,
                    .LastCardFormatted = lastCardFormatted,
                    .RawCardsSummaryJson = rawCardsJson
                }
            End Using
        Catch ex As Exception
            Dim detail As String = "Unable to retrieve profiling summary."
            If Not String.IsNullOrWhiteSpace(ex.Message) Then
                detail &= " " & ex.Message
            End If
            If ex.InnerException IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(ex.InnerException.Message) Then
                detail &= " (" & ex.InnerException.Message & ")"
            End If
            Return New ExemplarProfileLookupResult With {
                .IsSuccessful = False,
                .IsConfigured = True,
                .StatusText = "Error",
                .DetailText = detail,
                .ErrorMessage = ex.Message
            }
        End Try
    End Function

    Public Async Function GetStudentUnitProgressAsync(firstName As String, lastName As String, email As String, qualificationId As String, unitCodes As IEnumerable(Of String), Optional fallbackQualificationIds As IEnumerable(Of String) = Nothing) As Task(Of ExemplarUnitProgressResult)
        Dim token As String = GetBearerToken()
        If String.IsNullOrWhiteSpace(token) Then
            Return New ExemplarUnitProgressResult With {
                .IsSuccessful = False,
                .IsConfigured = False,
                .ErrorMessage = "Profiling API credentials are not configured."
            }
        End If

        Dim candidateIds As New List(Of String)()
        If Not String.IsNullOrWhiteSpace(qualificationId) Then candidateIds.Add(qualificationId.Trim())
        Dim configuredQualificationId As String = GetConfiguredQualificationId()
        If Not String.IsNullOrWhiteSpace(configuredQualificationId) Then candidateIds.Add(configuredQualificationId)
        If fallbackQualificationIds IsNot Nothing Then
            For Each candidate As String In fallbackQualificationIds
                If Not String.IsNullOrWhiteSpace(candidate) Then candidateIds.Add(candidate.Trim())
            Next
        End If
        candidateIds = candidateIds.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
        If candidateIds.Count = 0 Then
            Return New ExemplarUnitProgressResult With {
                .IsSuccessful = False,
                .IsConfigured = True,
                .ErrorMessage = "Qualification ID is not configured."
            }
        End If

        Try
            Dim student As ExemplarUserCandidate = Await FindStudentAsync(token, firstName, lastName, email)
            If student Is Nothing Then
                Return New ExemplarUnitProgressResult With {
                    .IsSuccessful = False,
                    .IsConfigured = True,
                    .ErrorMessage = "No matching Exemplar student was found."
                }
            End If

            Dim requestedCodes As New HashSet(Of String)(
                unitCodes.Where(Function(c) Not String.IsNullOrWhiteSpace(c)).
                Select(Function(c) c.Trim().ToUpperInvariant()),
                StringComparer.OrdinalIgnoreCase
            )

            Dim result As New ExemplarUnitProgressResult With {
                .IsSuccessful = True,
                .IsConfigured = True,
                .UserId = student.Id,
                .QualificationId = candidateIds(0)
            }

            Dim qualificationsCatalogUrl As String = $"{GetBaseUrl()}/api/v1/qualifications"
            Try
                Using qualificationsCatalogDoc As JsonDocument = Await SendJsonRequestAsync(HttpMethod.Get, qualificationsCatalogUrl, token, Nothing)
                    result.RawQualificationsCatalogJson = qualificationsCatalogDoc.RootElement.GetRawText()
                End Using
            Catch
                ' Keep going if the catalog endpoint is restricted in this tenant.
            End Try

            Dim qualificationsUrl As String = $"{GetBaseUrl()}/api/v1/users/{Uri.EscapeDataString(student.Id)}/qualifications"
            Try
                Using qualificationsDoc As JsonDocument = Await SendJsonRequestAsync(HttpMethod.Get, qualificationsUrl, token, Nothing)
                    result.RawQualificationsJson = qualificationsDoc.RootElement.GetRawText()
                End Using
            Catch
                ' Some tenants may not expose the list endpoint; keep going with the detail endpoint.
            End Try

            Dim lastError As String = ""
            Dim matched As Boolean = False
            For Each candidateQualificationId As String In candidateIds
                Dim qualificationUrl As String = $"{GetBaseUrl()}/api/v1/users/{Uri.EscapeDataString(student.Id)}/qualifications/{Uri.EscapeDataString(candidateQualificationId)}"
                Try
                    Using qualificationDoc As JsonDocument = Await SendJsonRequestAsync(HttpMethod.Get, qualificationUrl, token, Nothing)
                        result.RawApiJson = qualificationDoc.RootElement.GetRawText()
                        result.QualificationId = candidateQualificationId
                        CollectUnitProgress(qualificationDoc.RootElement, result.Units, requestedCodes, 0)
                        CollectExperienceAndDemonstrationPercentages(qualificationDoc.RootElement, result.Units, requestedCodes)
                        matched = True
                    End Using
                    Exit For
                Catch ex As Exception
                    lastError = ex.Message
                End Try
            Next

            If Not matched Then
                Return New ExemplarUnitProgressResult With {
                    .IsSuccessful = False,
                    .IsConfigured = True,
                    .ErrorMessage = If(String.IsNullOrWhiteSpace(lastError), "Unable to retrieve qualification details.", lastError)
                }
            End If

            For Each requestedCode In requestedCodes
                If Not result.Units.ContainsKey(requestedCode) Then
                    result.Units(requestedCode) = New ExemplarUnitProgressItem With {
                        .UnitCode = requestedCode
                    }
                End If
            Next

            For Each code In requestedCodes
                Dim unitItem As ExemplarUnitProgressItem = result.Units(code)
                Dim countsResult As CardCountsResult = Await TryGetContributingCardCountsAsync(token, student.Id, result.QualificationId, code)
                If countsResult IsNot Nothing Then
                    ' Always store the raw endpoint/cards lists for debugging and card-type percentages.
                    unitItem.ExperienceCardsRawJson = countsResult.ExperienceCardsRawJson
                    unitItem.DemonstrationCardsRawJson = countsResult.DemonstrationCardsRawJson
                    unitItem.ProgressionCardsEndpointRawJson = countsResult.ProgressionCardsEndpointRawJson

                    ' Only store Total/Completed and compute unit overall % when the endpoint returned
                    ' the status/count fields we use for the approximation.
                    If countsResult.Success Then
                        unitItem.TotalCards = countsResult.TotalCards
                        unitItem.CompletedCards = countsResult.CompletedCards

                        If Not unitItem.Percentage.HasValue AndAlso countsResult.TotalCards > 0 Then
                            unitItem.Percentage = Math.Round(CDec((countsResult.CompletedCards / countsResult.TotalCards) * 100D), 2)
                        End If
                    End If
                End If
            Next

            Return result
        Catch ex As Exception
            Return New ExemplarUnitProgressResult With {
                .IsSuccessful = False,
                .IsConfigured = True,
                .ErrorMessage = ex.Message
            }
        End Try
    End Function

    Public Async Function UpdateQualificationStatusAsync(userId As String, qualificationId As String, statusValue As String) As Task(Of ExemplarQualificationUpdateResult)
        Dim token As String = GetBearerToken()
        If String.IsNullOrWhiteSpace(token) Then
            Return New ExemplarQualificationUpdateResult With {
                .IsSuccessful = False,
                .ErrorMessage = "Profiling API credentials are not configured."
            }
        End If

        If String.IsNullOrWhiteSpace(userId) OrElse String.IsNullOrWhiteSpace(qualificationId) Then
            Return New ExemplarQualificationUpdateResult With {
                .IsSuccessful = False,
                .ErrorMessage = "User ID and Qualification ID are required."
            }
        End If

        Dim normalizedStatus As String = statusValue?.Trim().ToUpperInvariant()
        Dim allowed As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "ACTIVE",
            "INACTIVE",
            "COMPLETED",
            "WITHDRAWN"
        }

        If Not allowed.Contains(normalizedStatus) Then
            Return New ExemplarQualificationUpdateResult With {
                .IsSuccessful = False,
                .ErrorMessage = "Invalid status value."
            }
        End If

        Dim endpointUrl As String = $"{GetBaseUrl()}/api/v1/users/{Uri.EscapeDataString(userId)}/qualifications/{Uri.EscapeDataString(qualificationId)}"
        Dim payload As String = JsonSerializer.Serialize(New With {.status = normalizedStatus})

        Try
            Using request As New HttpRequestMessage(HttpMethod.Put, endpointUrl)
                request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", token)
                request.Headers.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
                request.Content = New StringContent(payload, Encoding.UTF8, "application/json")

                Using response As HttpResponseMessage = Await ApiClient.SendAsync(request).ConfigureAwait(False)
                    Dim body As String = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                    If response.IsSuccessStatusCode Then
                        Return New ExemplarQualificationUpdateResult With {
                            .IsSuccessful = True,
                            .ResponseBody = body
                        }
                    End If

                    Return New ExemplarQualificationUpdateResult With {
                        .IsSuccessful = False,
                        .ErrorMessage = $"{CInt(response.StatusCode)} {response.ReasonPhrase}",
                        .ResponseBody = body
                    }
                End Using
            End Using
        Catch ex As Exception
            Return New ExemplarQualificationUpdateResult With {
                .IsSuccessful = False,
                .ErrorMessage = ex.Message
            }
        End Try
    End Function

    Private Async Function FindStudentAsync(token As String, firstName As String, lastName As String, email As String) As Task(Of ExemplarUserCandidate)
        Dim searchUrl As String = $"{GetBaseUrl()}/api/v1/users?roles=STUDENT&firstName={WebUtility.UrlEncode(firstName)}&lastName={WebUtility.UrlEncode(lastName)}"
        If Not String.IsNullOrWhiteSpace(email) Then
            searchUrl &= "&email=" & WebUtility.UrlEncode(email.Trim())
        End If

        Using doc As JsonDocument = Await SendJsonRequestAsync(HttpMethod.Get, searchUrl, token, Nothing)
            Dim candidates As New List(Of ExemplarUserCandidate)()
            CollectUserCandidates(doc.RootElement, candidates, 0)
            Return SelectBestCandidate(candidates, firstName, lastName, email)
        End Using
    End Function

    Private Async Function SendJsonRequestAsync(method As HttpMethod, url As String, token As String, payload As String) As Task(Of JsonDocument)
        Dim doc As JsonDocument = Await SendJsonRequestCoreAsync(method, url, token, payload).ConfigureAwait(False)
        If doc Is Nothing Then
            ' 401 or failure: try once to refresh token (run JAR) and retry
            ClearCachedToken()
            ' Force a token refresh from the login JAR so we don't keep reusing an
            ' old bearer token from settings/env when credentials have changed.
            Dim newToken As String = TryRefreshTokenFromJar()
            If Not String.IsNullOrWhiteSpace(newToken) Then
                SetBearerToken(newToken)
                doc = Await SendJsonRequestCoreAsync(method, url, CachedToken, payload).ConfigureAwait(False)
            End If
        End If
        If doc Is Nothing Then
            Dim msg As String
            If Not String.IsNullOrWhiteSpace(LastTokenRefreshError) Then
                msg = "Exemplar API request failed or returned 401 and token refresh did not succeed. " & LastTokenRefreshError
            Else
                ' We got a new token but the API still returned 401 = JAR and API mismatch
                msg = "Exemplar API returned 401 after token refresh (token rejected by API). " &
                      "baseUrl=" & If(String.IsNullOrWhiteSpace(LastTokenRefreshBaseUrl), "(unknown)", LastTokenRefreshBaseUrl) &
                      " | jar=" & If(String.IsNullOrWhiteSpace(LastTokenRefreshJarPath), "(unknown)", LastTokenRefreshJarPath)
            End If
            Throw New HttpRequestException(msg)
        End If
        Return doc
    End Function

    Private Async Function SendJsonRequestCoreAsync(method As HttpMethod, url As String, token As String, payload As String) As Task(Of JsonDocument)
        Using request As New HttpRequestMessage(method, url)
            request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", token)
            request.Headers.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
            If payload IsNot Nothing Then
                request.Content = New StringContent(payload, Encoding.UTF8, "application/json")
            End If

            Using response As HttpResponseMessage = Await ApiClient.SendAsync(request).ConfigureAwait(False)
                Dim body As String = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                If response.StatusCode = HttpStatusCode.Unauthorized Then
                    Return Nothing
                End If
                response.EnsureSuccessStatusCode()
                If String.Equals(Environment.GetEnvironmentVariable("EXEMPLAR_DEBUG_LIST_RESPONSES"), "1", StringComparison.OrdinalIgnoreCase) Then
                    DebugResponseList.Add(Tuple.Create(url, body))
                End If
                Return JsonDocument.Parse(body)
            End Using
        End Using
    End Function

    Private Async Function TryGetContributingCardCountsAsync(token As String, userId As String, qualificationId As String, unitCode As String) As Task(Of CardCountsResult)
        Dim url As String = $"{GetBaseUrl()}/api/v1/users/{Uri.EscapeDataString(userId)}/qualifications/{Uri.EscapeDataString(qualificationId)}/units/{Uri.EscapeDataString(unitCode)}/progression/cards"

        Try
            Using doc As JsonDocument = Await SendJsonRequestAsync(HttpMethod.Get, url, token, Nothing)
                Dim progressionCardsRawJson As String = doc.RootElement.GetRawText()
                Dim statusCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                CollectStatusCounts(doc.RootElement, statusCounts, 0)

                ' Extract the per-unit cards lists for debug/UI inspection.
                Dim experienceCardsRaw As String = TryGetFirstCardListRawJson(doc.RootElement,
                    New String() {"experience_cards", "experienceCards"})
                Dim demonstrationCardsRaw As String = TryGetFirstCardListRawJson(doc.RootElement,
                    New String() {"demonstration_cards", "demonstrationCards"})

                If statusCounts.Count = 0 Then
                    ' Endpoint may still include the experience/demonstration cards lists,
                    ' but not the status/count fields we use for TotalCards/CompletedCards.
                    Return New CardCountsResult With {
                        .Success = False,
                        .TotalCards = 0,
                        .CompletedCards = 0,
                        .ExperienceCardsRawJson = experienceCardsRaw,
                        .DemonstrationCardsRawJson = demonstrationCardsRaw,
                        .ProgressionCardsEndpointRawJson = progressionCardsRawJson
                    }
                End If

                Dim totalCards As Integer = statusCounts.Values.Sum()
                Dim completedCards As Integer = statusCounts.
                    Where(Function(kvp) kvp.Key.IndexOf("APPROVED", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                      kvp.Key.IndexOf("COMPLETE", StringComparison.OrdinalIgnoreCase) >= 0).
                    Sum(Function(kvp) kvp.Value)

                Return New CardCountsResult With {
                    .Success = True,
                    .TotalCards = totalCards,
                    .CompletedCards = completedCards,
                    .ExperienceCardsRawJson = experienceCardsRaw,
                    .DemonstrationCardsRawJson = demonstrationCardsRaw,
                    .ProgressionCardsEndpointRawJson = progressionCardsRawJson
                }
            End Using
        Catch
            Return New CardCountsResult With {.Success = False}
        End Try
    End Function

    Private Function TryGetFirstCardListRawJson(root As JsonElement, possiblePropertyNames As IEnumerable(Of String)) As String
        Dim wantsExperience As Boolean = possiblePropertyNames.Any(Function(n) n.IndexOf("experience", StringComparison.OrdinalIgnoreCase) >= 0)
        Dim wantsDemonstration As Boolean = possiblePropertyNames.Any(Function(n) n.IndexOf("demonstration", StringComparison.OrdinalIgnoreCase) >= 0)
        Return TryGetFirstCardListRawJsonInternal(root, possiblePropertyNames, wantsExperience, wantsDemonstration, False, False, 0)
    End Function

    ' Context-aware extraction:
    ' Some tenants wrap the card arrays like:
    '   { "experience": { "cards": [ ... ] }, "demonstration": { "cards": [ ... ] } }
    ' In that shape, the key "cards" has no "experience" substring, so we must track whether we're
    ' already inside an "experience" or "demonstration" branch.
    Private Function TryGetFirstCardListRawJsonInternal(element As JsonElement,
                                                        possiblePropertyNames As IEnumerable(Of String),
                                                        wantsExperience As Boolean,
                                                        wantsDemonstration As Boolean,
                                                        inExperienceContext As Boolean,
                                                        inDemonstrationContext As Boolean,
                                                        depth As Integer) As String
        If depth > 20 Then Return ""

        Select Case element.ValueKind
            Case JsonValueKind.Object
                For Each prop As JsonProperty In element.EnumerateObject()
                    ' Exact key match first (covers array or object-wrapped).
                    If possiblePropertyNames.Any(Function(n) String.Equals(n, prop.Name, StringComparison.OrdinalIgnoreCase)) Then
                        Return prop.Value.GetRawText()
                    End If

                    Dim propLower As String = prop.Name.ToLowerInvariant()
                    Dim nextInExperience As Boolean = inExperienceContext OrElse (wantsExperience AndAlso propLower.Contains("experience"))
                    Dim nextInDemonstration As Boolean = inDemonstrationContext OrElse (wantsDemonstration AndAlso propLower.Contains("demonstration"))

                    ' If we find an array inside the relevant context, return it.
                    If prop.Value.ValueKind = JsonValueKind.Array Then
                        If (wantsExperience AndAlso nextInExperience) OrElse (wantsDemonstration AndAlso nextInDemonstration) Then
                            Return prop.Value.GetRawText()
                        End If
                    End If

                    Dim nested As String = TryGetFirstCardListRawJsonInternal(prop.Value,
                                                                              possiblePropertyNames,
                                                                              wantsExperience,
                                                                              wantsDemonstration,
                                                                              nextInExperience,
                                                                              nextInDemonstration,
                                                                              depth + 1)
                    If Not String.IsNullOrWhiteSpace(nested) Then Return nested
                Next

            Case JsonValueKind.Array
                ' If the whole node is an array and we're already in context, accept it.
                If (wantsExperience AndAlso inExperienceContext) OrElse (wantsDemonstration AndAlso inDemonstrationContext) Then
                    Return element.GetRawText()
                End If

                For Each item As JsonElement In element.EnumerateArray()
                    Dim nested As String = TryGetFirstCardListRawJsonInternal(item,
                                                                              possiblePropertyNames,
                                                                              wantsExperience,
                                                                              wantsDemonstration,
                                                                              inExperienceContext,
                                                                              inDemonstrationContext,
                                                                              depth + 1)
                    If Not String.IsNullOrWhiteSpace(nested) Then Return nested
                Next
        End Select

        Return ""
    End Function

    Private Sub CollectStatusCounts(element As JsonElement, output As Dictionary(Of String, Integer), depth As Integer)
        If depth > 12 Then
            Return
        End If

        Select Case element.ValueKind
            Case JsonValueKind.Object
                Dim status As String = GetStringByPossibleKeys(element, New String() {"status", "card_status", "state"})
                Dim count As Integer
                If Not String.IsNullOrWhiteSpace(status) AndAlso TryGetIntegerByPossibleKeys(element, New String() {"count", "total", "value"}, count) Then
                    If output.ContainsKey(status) Then
                        output(status) += count
                    Else
                        output(status) = count
                    End If
                End If

                For Each prop As JsonProperty In element.EnumerateObject()
                    CollectStatusCounts(prop.Value, output, depth + 1)
                Next

            Case JsonValueKind.Array
                For Each item As JsonElement In element.EnumerateArray()
                    CollectStatusCounts(item, output, depth + 1)
                Next
        End Select
    End Sub

    Private Sub CollectUnitProgress(element As JsonElement, output As Dictionary(Of String, ExemplarUnitProgressItem), requestedCodes As HashSet(Of String), depth As Integer)
        If depth > 12 Then
            Return
        End If

        Select Case element.ValueKind
            Case JsonValueKind.Object
                Dim codeCandidate As String = GetStringByPossibleKeys(element, New String() {"unit_code", "unitCode", "unit", "unit_id", "unitId"})
                If LooksLikeUnitCode(codeCandidate) Then
                    Dim normalizedCode As String = codeCandidate.Trim().ToUpperInvariant()
                    If requestedCodes Is Nothing OrElse requestedCodes.Contains(normalizedCode) Then
                        If Not output.ContainsKey(normalizedCode) Then
                            output(normalizedCode) = New ExemplarUnitProgressItem With {
                                .UnitCode = normalizedCode
                            }
                        End If

                        Dim percentageValue As Decimal
                        If TryGetDecimalByPossibleKeys(element, New String() {"percentage", "completion_percentage", "completionPercentage", "progress_percentage", "progressPercentage"}, percentageValue) Then
                            output(normalizedCode).Percentage = percentageValue
                        End If

                        Dim statusValue As String = GetStringByPossibleKeys(element, New String() {"status", "state", "completion_status"})
                        If Not String.IsNullOrWhiteSpace(statusValue) Then
                            output(normalizedCode).Status = statusValue
                        End If
                    End If
                End If

                For Each prop As JsonProperty In element.EnumerateObject()
                    CollectUnitProgress(prop.Value, output, requestedCodes, depth + 1)
                Next

            Case JsonValueKind.Array
                For Each item As JsonElement In element.EnumerateArray()
                    CollectUnitProgress(item, output, requestedCodes, depth + 1)
                Next
        End Select
    End Sub

    ' Parses qualification payload:
    ' - calculation.demonstration_progression.unit_demonstration_statuses[*].percentage
    ' - calculation.experience_progression.unit_experience_statuses[*].percentage (if present)
    Private Sub CollectExperienceAndDemonstrationPercentages(root As JsonElement,
                                                              output As Dictionary(Of String, ExemplarUnitProgressItem),
                                                              requestedCodes As HashSet(Of String))
        Try
            CollectUnitPercentFromStatusArray(root,
                                               output,
                                               requestedCodes,
                                               "unit_demonstration_statuses",
                                               Function(item As JsonElement) True,
                                               Sub(unitCode As String, pct As Decimal)
                                                   Dim normalized As String = unitCode.Trim().ToUpperInvariant()
                                                   If output.ContainsKey(normalized) Then
                                                       output(normalized).DemonstrationCardsPercentage = pct
                                                   End If
                                               End Sub)

            ' Experience may not exist on all qualification payloads.
            CollectUnitPercentFromStatusArray(root,
                                               output,
                                               requestedCodes,
                                               "unit_experience_statuses",
                                               Function(item As JsonElement) True,
                                               Sub(unitCode As String, pct As Decimal)
                                                   Dim normalized As String = unitCode.Trim().ToUpperInvariant()
                                                   If output.ContainsKey(normalized) Then
                                                       output(normalized).ExperienceCardsPercentage = pct
                                                   End If
                                               End Sub)
        Catch
            ' Best-effort parsing; UI still shows N/A if unavailable.
        End Try
    End Sub

    Private Sub CollectUnitPercentFromStatusArray(root As JsonElement,
                                                    output As Dictionary(Of String, ExemplarUnitProgressItem),
                                                    requestedCodes As HashSet(Of String),
                                                    statusArrayPropertyName As String,
                                                    predicate As Func(Of JsonElement, Boolean),
                                                    applyPercent As Action(Of String, Decimal),
                                                    Optional depth As Integer = 0)
        If depth > 20 Then Return

        Select Case root.ValueKind
            Case JsonValueKind.Object
                For Each prop As JsonProperty In root.EnumerateObject()
                    Dim propName = prop.Name
                    If String.Equals(propName, statusArrayPropertyName, StringComparison.OrdinalIgnoreCase) Then
                        If prop.Value.ValueKind = JsonValueKind.Array Then
                            For Each item In prop.Value.EnumerateArray()
                                If Not predicate(item) Then Continue For

                                Dim unitCode As String = GetStringByPossibleKeys(item, New String() {"unit_code", "unitCode", "unit", "unit_id", "unitId"})
                                If String.IsNullOrWhiteSpace(unitCode) Then
                                    unitCode = GetStringFromNestedId(item)
                                End If

                                If String.IsNullOrWhiteSpace(unitCode) Then Continue For

                                Dim pct As Decimal
                                If TryGetDecimalByPossibleKeys(item, New String() {"percentage", "completion_percentage", "completionPercentage", "progress_percentage", "progressPercentage"}, pct) Then
                                    Dim normalized As String = unitCode.Trim().ToUpperInvariant()
                                    If requestedCodes Is Nothing OrElse requestedCodes.Contains(normalized) Then
                                        applyPercent(unitCode, pct)
                                    End If
                                End If
                            Next
                        End If
                    Else
                        CollectUnitPercentFromStatusArray(prop.Value, output, requestedCodes, statusArrayPropertyName, predicate, applyPercent, depth + 1)
                    End If
                Next

            Case JsonValueKind.Array
                For Each item In root.EnumerateArray()
                    CollectUnitPercentFromStatusArray(item, output, requestedCodes, statusArrayPropertyName, predicate, applyPercent, depth + 1)
                Next
        End Select
    End Sub

    Private Function GetStringFromNestedId(item As JsonElement) As String
        If item.ValueKind <> JsonValueKind.Object Then Return ""
        For Each prop As JsonProperty In item.EnumerateObject()
            If String.Equals(prop.Name, "id", StringComparison.OrdinalIgnoreCase) AndAlso prop.Value.ValueKind = JsonValueKind.Object Then
                Return GetStringByPossibleKeys(prop.Value, New String() {"unit_code", "unitCode", "unit"})
            End If
        Next
        Return ""
    End Function

    Private Function LooksLikeUnitCode(value As String) As Boolean
        If String.IsNullOrWhiteSpace(value) Then
            Return False
        End If

        Dim candidate As String = value.Trim().ToUpperInvariant()
        If candidate.Contains(" "c) Then
            Return False
        End If

        If candidate.Length < 5 Then
            Return False
        End If

        Dim hasLetter As Boolean = candidate.Any(Function(c) Char.IsLetter(c))
        Dim hasDigit As Boolean = candidate.Any(Function(c) Char.IsDigit(c))
        Return hasLetter AndAlso hasDigit
    End Function

    Private Sub CollectUserCandidates(element As JsonElement, output As List(Of ExemplarUserCandidate), depth As Integer)
        If depth > 12 Then
            Return
        End If

        Select Case element.ValueKind
            Case JsonValueKind.Object
                Dim id As String = GetStringByPossibleKeys(element, New String() {"id", "user_id", "userId"})
                Dim firstName As String = GetStringByPossibleKeys(element, New String() {"first_name", "firstName", "given_name", "givenName"})
                Dim lastName As String = GetStringByPossibleKeys(element, New String() {"last_name", "lastName", "family_name", "familyName"})
                Dim email As String = GetStringByPossibleKeys(element, New String() {"email", "student_email", "personal_email"})

                If Not String.IsNullOrWhiteSpace(id) AndAlso
                    (Not String.IsNullOrWhiteSpace(firstName) OrElse Not String.IsNullOrWhiteSpace(lastName) OrElse Not String.IsNullOrWhiteSpace(email)) Then
                    output.Add(New ExemplarUserCandidate With {
                        .Id = id,
                        .FirstName = firstName,
                        .LastName = lastName,
                        .Email = email
                    })
                End If

                For Each prop As JsonProperty In element.EnumerateObject()
                    CollectUserCandidates(prop.Value, output, depth + 1)
                Next

            Case JsonValueKind.Array
                For Each item As JsonElement In element.EnumerateArray()
                    CollectUserCandidates(item, output, depth + 1)
                Next
        End Select
    End Sub

    Private Function SelectBestCandidate(candidates As List(Of ExemplarUserCandidate), firstName As String, lastName As String, email As String) As ExemplarUserCandidate
        If candidates Is Nothing OrElse candidates.Count = 0 Then
            Return Nothing
        End If

        Dim normalizedEmail As String = Normalize(email)
        Dim normalizedFirst As String = Normalize(firstName)
        Dim normalizedLast As String = Normalize(lastName)

        If Not String.IsNullOrWhiteSpace(normalizedEmail) Then
            For Each candidate In candidates
                If Normalize(candidate.Email) = normalizedEmail Then
                    Return candidate
                End If
            Next
        End If

        If Not String.IsNullOrWhiteSpace(normalizedFirst) AndAlso Not String.IsNullOrWhiteSpace(normalizedLast) Then
            For Each candidate In candidates
                If Normalize(candidate.FirstName) = normalizedFirst AndAlso Normalize(candidate.LastName) = normalizedLast Then
                    Return candidate
                End If
            Next
        End If

        Return candidates(0)
    End Function

    Private Function FindFirstInteger(element As JsonElement, keys As IEnumerable(Of String)) As Integer?
        Dim keySet As New HashSet(Of String)(keys.Select(Function(k) k.ToLowerInvariant()))
        Dim foundValue As Integer
        If TryFindIntByKeyRecursive(element, keySet, foundValue, 0) Then
            Return foundValue
        End If

        Return Nothing
    End Function

    Private Function FindFirstString(element As JsonElement, keys As IEnumerable(Of String)) As String
        Dim keySet As New HashSet(Of String)(keys.Select(Function(k) k.ToLowerInvariant()))
        Dim foundValue As String = Nothing
        If TryFindStringByKeyRecursive(element, keySet, foundValue, 0) Then
            Return foundValue
        End If

        Return ""
    End Function

    Private Function TryFindIntByKeyRecursive(element As JsonElement, keys As HashSet(Of String), ByRef value As Integer, depth As Integer) As Boolean
        If depth > 12 Then
            Return False
        End If

        Select Case element.ValueKind
            Case JsonValueKind.Object
                For Each prop As JsonProperty In element.EnumerateObject()
                    Dim propName As String = prop.Name.ToLowerInvariant()
                    If keys.Contains(propName) Then
                        Dim parsed As Integer
                        If TryParseInteger(prop.Value, parsed) Then
                            value = parsed
                            Return True
                        End If
                    End If

                    If TryFindIntByKeyRecursive(prop.Value, keys, value, depth + 1) Then
                        Return True
                    End If
                Next

            Case JsonValueKind.Array
                For Each item As JsonElement In element.EnumerateArray()
                    If TryFindIntByKeyRecursive(item, keys, value, depth + 1) Then
                        Return True
                    End If
                Next
        End Select

        Return False
    End Function

    Private Function TryFindStringByKeyRecursive(element As JsonElement, keys As HashSet(Of String), ByRef value As String, depth As Integer) As Boolean
        If depth > 12 Then
            Return False
        End If

        Select Case element.ValueKind
            Case JsonValueKind.Object
                For Each prop As JsonProperty In element.EnumerateObject()
                    Dim propName As String = prop.Name.ToLowerInvariant()
                    If keys.Contains(propName) Then
                        Dim parsed As String = TryParseString(prop.Value)
                        If Not String.IsNullOrWhiteSpace(parsed) Then
                            value = parsed
                            Return True
                        End If
                    End If

                    If TryFindStringByKeyRecursive(prop.Value, keys, value, depth + 1) Then
                        Return True
                    End If
                Next

            Case JsonValueKind.Array
                For Each item As JsonElement In element.EnumerateArray()
                    If TryFindStringByKeyRecursive(item, keys, value, depth + 1) Then
                        Return True
                    End If
                Next
        End Select

        Return False
    End Function

    ''' <summary>Flattens a JsonElement to "path = value" lines so you can see every field name and value. Used when EXEMPLAR_DEBUG_LIST_RESPONSES=1.</summary>
    Private Sub FlattenJsonToLines(element As JsonElement, prefix As String, output As List(Of String), Optional depth As Integer = 0)
        If depth > 15 Then Return
        Select Case element.ValueKind
            Case JsonValueKind.Object
                For Each prop As JsonProperty In element.EnumerateObject()
                    FlattenJsonToLines(prop.Value, If(String.IsNullOrEmpty(prefix), prop.Name, prefix & "." & prop.Name), output, depth + 1)
                Next
            Case JsonValueKind.Array
                Dim i As Integer = 0
                For Each item As JsonElement In element.EnumerateArray()
                    FlattenJsonToLines(item, prefix & "[" & i & "]", output, depth + 1)
                    i += 1
                Next
            Case JsonValueKind.String
                Dim s As String = element.GetString()
                If s Is Nothing Then s = ""
                output.Add((If(String.IsNullOrEmpty(prefix), "(root)", prefix) & " = " & s).Replace(vbCr, " ").Replace(vbLf, " "))
            Case JsonValueKind.Number
                Dim n As Long
                If element.TryGetInt64(n) Then
                    output.Add((If(String.IsNullOrEmpty(prefix), "(root)", prefix) & " = " & n.ToString()))
                Else
                    output.Add((If(String.IsNullOrEmpty(prefix), "(root)", prefix) & " = " & element.GetRawText()))
                End If
            Case JsonValueKind.True, JsonValueKind.False
                output.Add((If(String.IsNullOrEmpty(prefix), "(root)", prefix) & " = " & element.GetBoolean().ToString()))
            Case JsonValueKind.Null
                output.Add((If(String.IsNullOrEmpty(prefix), "(root)", prefix) & " = (null)"))
        End Select
    End Sub

    ''' <summary>Finds a cards/data/items array and counts items by status: submitted (SUBMITTED/COMPLETED/APPROVED) vs outstanding (anything else).</summary>
    Private Sub CountSubmittedAndOutstandingFromArray(element As JsonElement, ByRef submitted As Integer, ByRef outstanding As Integer, Optional depth As Integer = 0)
        If depth > 8 Then Return
        Select Case element.ValueKind
            Case JsonValueKind.Array
                For Each item As JsonElement In element.EnumerateArray()
                    ' Only objects can have a status property; skip strings/numbers
                    If item.ValueKind <> JsonValueKind.Object Then Continue For
                    Dim status As String = GetStringByPossibleKeys(item, New String() {"status", "state", "cardStatus", "completionStatus"})
                    If Not String.IsNullOrWhiteSpace(status) Then
                        Dim s As String = status.Trim().ToUpperInvariant()
                        If s = "SUBMITTED" OrElse s = "COMPLETED" OrElse s = "APPROVED" OrElse s = "DONE" Then
                            submitted += 1
                        Else
                            outstanding += 1
                        End If
                    End If
                Next
            Case JsonValueKind.Object
                For Each prop As JsonProperty In element.EnumerateObject()
                    CountSubmittedAndOutstandingFromArray(prop.Value, submitted, outstanding, depth + 1)
                Next
        End Select
    End Sub

    Private Function TryGetIntegerByPossibleKeys(element As JsonElement, keys As IEnumerable(Of String), ByRef value As Integer) As Boolean
        For Each key In keys
            Dim prop As JsonElement
            If element.TryGetProperty(key, prop) Then
                Return TryParseInteger(prop, value)
            End If
        Next

        Return False
    End Function

    Private Function TryGetDecimalByPossibleKeys(element As JsonElement, keys As IEnumerable(Of String), ByRef value As Decimal) As Boolean
        For Each key In keys
            Dim prop As JsonElement
            If element.TryGetProperty(key, prop) Then
                Return TryParseDecimal(prop, value)
            End If
        Next

        Return False
    End Function

    Private Function TryParseInteger(value As JsonElement, ByRef result As Integer) As Boolean
        Select Case value.ValueKind
            Case JsonValueKind.Number
                If value.TryGetInt32(result) Then
                    Return True
                End If

                Dim asLong As Long
                If value.TryGetInt64(asLong) AndAlso asLong <= Integer.MaxValue AndAlso asLong >= Integer.MinValue Then
                    result = CInt(asLong)
                    Return True
                End If

            Case JsonValueKind.String
                Return Integer.TryParse(value.GetString(), result)
        End Select

        Return False
    End Function

    Private Function TryParseDecimal(value As JsonElement, ByRef result As Decimal) As Boolean
        Select Case value.ValueKind
            Case JsonValueKind.Number
                Return value.TryGetDecimal(result)
            Case JsonValueKind.String
                Return Decimal.TryParse(value.GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, result) OrElse
                    Decimal.TryParse(value.GetString(), NumberStyles.Any, CultureInfo.CurrentCulture, result)
        End Select

        Return False
    End Function

    Private Function TryParseString(value As JsonElement) As String
        Select Case value.ValueKind
            Case JsonValueKind.String
                Return value.GetString()
            Case JsonValueKind.Number, JsonValueKind.True, JsonValueKind.False
                Return value.ToString()
            Case Else
                Return ""
        End Select
    End Function

    ''' <summary>Fills notVerified and employerVerified from card_status_counts array by matching status/name (when index-based paths fail).</summary>
    Private Sub GetCardStatusCountsFromArray(root As JsonElement, ByRef notVerified As Integer?, ByRef employerVerified As Integer?)
        Dim arr As JsonElement
        If Not root.TryGetProperty("card_status_counts", arr) OrElse arr.ValueKind <> JsonValueKind.Array Then
            Return
        End If
        For Each item As JsonElement In arr.EnumerateArray()
            If item.ValueKind <> JsonValueKind.Object Then Continue For
            Dim count As Integer? = Nothing
            Dim countProp As JsonElement
            If item.TryGetProperty("count", countProp) Then
                Dim c As Integer
                If TryParseInteger(countProp, c) Then count = c
            End If
            If Not count.HasValue Then Continue For
            Dim status As String = GetStringByPossibleKeys(item, New String() {"status", "name", "status_name", "type"})
            If String.IsNullOrWhiteSpace(status) Then Continue For
            Dim s As String = status.Trim().ToUpperInvariant()
            If (s.Contains("NOT_VERIFIED") OrElse s.Contains("SUBMITTED") OrElse s = "PENDING") AndAlso Not notVerified.HasValue Then
                notVerified = count
            ElseIf (s.Contains("EMPLOYER_VERIFIED") OrElse s.Contains("VERIFIED") OrElse s = "APPROVED") AndAlso Not employerVerified.HasValue Then
                employerVerified = count
            End If
        Next
    End Sub

    ''' <summary>Gets an integer by path, e.g. "card_submission.missing_weeks" or "card_status_counts[0].count".</summary>
    Private Function GetIntegerByPath(root As JsonElement, path As String) As Integer?
        If String.IsNullOrWhiteSpace(path) Then Return Nothing
        Dim segments As String() = path.Split("."c)
        Dim current As JsonElement = root
        For Each seg As String In segments
            seg = seg.Trim()
            If current.ValueKind <> JsonValueKind.Object AndAlso current.ValueKind <> JsonValueKind.Array Then Return Nothing
            Dim propName As String = seg
            Dim arrayIndex As Integer? = Nothing
            Dim bracketStart As Integer = seg.IndexOf("["c)
            If bracketStart >= 0 Then
                Dim bracketEnd As Integer = seg.IndexOf("]"c)
                If bracketEnd > bracketStart Then
                    propName = seg.Substring(0, bracketStart)
                    Dim indexStr As String = seg.Substring(bracketStart + 1, bracketEnd - bracketStart - 1)
                    Dim idx As Integer
                    If Integer.TryParse(indexStr.Trim(), idx) Then arrayIndex = idx
                End If
            End If
            If current.ValueKind = JsonValueKind.Object Then
                If Not current.TryGetProperty(propName, current) Then Return Nothing
            End If
            If arrayIndex.HasValue Then
                If current.ValueKind <> JsonValueKind.Array Then Return Nothing
                Dim i As Integer = 0
                Dim found As Boolean = False
                For Each item As JsonElement In current.EnumerateArray()
                    If i = arrayIndex.Value Then
                        current = item
                        found = True
                        Exit For
                    End If
                    i += 1
                Next
                If Not found Then Return Nothing
            End If
        Next
        Dim result As Integer
        If TryParseInteger(current, result) Then Return result
        Return Nothing
    End Function

    ''' <summary>Gets a string by path, e.g. "last_card_timestamp" or "card_submission.last_card_timestamp".</summary>
    Private Function GetStringByPath(root As JsonElement, path As String) As String
        If String.IsNullOrWhiteSpace(path) Then Return ""
        Dim segments As String() = path.Split("."c)
        Dim current As JsonElement = root
        For Each seg As String In segments
            seg = seg.Trim()
            If current.ValueKind <> JsonValueKind.Object AndAlso current.ValueKind <> JsonValueKind.Array Then Return ""
            Dim propName As String = seg
            Dim arrayIndex As Integer? = Nothing
            Dim bracketStart As Integer = seg.IndexOf("["c)
            If bracketStart >= 0 Then
                Dim bracketEnd As Integer = seg.IndexOf("]"c)
                If bracketEnd > bracketStart Then
                    propName = seg.Substring(0, bracketStart)
                    Dim indexStr As String = seg.Substring(bracketStart + 1, bracketEnd - bracketStart - 1)
                    Dim idx As Integer
                    If Integer.TryParse(indexStr.Trim(), idx) Then arrayIndex = idx
                End If
            End If
            If current.ValueKind = JsonValueKind.Object Then
                If Not current.TryGetProperty(propName, current) Then Return ""
            End If
            If arrayIndex.HasValue Then
                If current.ValueKind <> JsonValueKind.Array Then Return ""
                Dim i As Integer = 0
                Dim found As Boolean = False
                For Each item As JsonElement In current.EnumerateArray()
                    If i = arrayIndex.Value Then
                        current = item
                        found = True
                        Exit For
                    End If
                    i += 1
                Next
                If Not found Then Return ""
            End If
        Next
        Dim s As String = TryParseString(current)
        Return If(s, "")
    End Function

    Private Function GetStringByPossibleKeys(element As JsonElement, keys As IEnumerable(Of String)) As String
        ' TryGetProperty requires element to be Object; string/array/number would throw
        If element.ValueKind <> JsonValueKind.Object Then
            Return ""
        End If
        For Each key As String In keys
            Dim prop As JsonElement
            If element.TryGetProperty(key, prop) Then
                Dim s As String = TryParseString(prop)
                If s IsNot Nothing Then Return s
            End If
        Next

        Return ""
    End Function

    Private Function GetBaseUrl() As String
        If Not String.IsNullOrWhiteSpace(My.Settings.ExemplarApiBaseUrl) Then
            Return My.Settings.ExemplarApiBaseUrl.Trim().TrimEnd("/"c)
        End If

        Dim envBase As String = Environment.GetEnvironmentVariable("EXEMPLAR_API_BASE_URL")
        If String.IsNullOrWhiteSpace(envBase) Then
            Return DefaultBaseUrl
        End If

        Return envBase.Trim().TrimEnd("/"c)
    End Function

    Private Function GetBearerToken() As String
        If Not String.IsNullOrWhiteSpace(CachedToken) Then
            Return CachedToken
        End If

        Dim token As String = My.Settings.ExemplarApiToken
        If String.IsNullOrWhiteSpace(token) Then
            token = Environment.GetEnvironmentVariable("EXEMPLAR_API_TOKEN")
        End If
        If String.IsNullOrWhiteSpace(token) Then
            token = Environment.GetEnvironmentVariable("EXEMPLAR_BEARER_TOKEN")
        End If

        If String.IsNullOrWhiteSpace(token) Then
            ' Auto-refresh: if username/password and JAR are configured, run login JAR to get a token
            token = TryRefreshTokenFromJar()
        End If

        If String.IsNullOrWhiteSpace(token) Then
            Return ""
        End If

        SetBearerToken(token)
        Return CachedToken
    End Function

    ''' <summary>Runs the Exemplar login JAR with username/password (hardcoded above; env vars override) and returns the Bearer token. Uses embedded JAR if present, else EXEMPLAR_LOGIN_JAR_PATH or app directory.</summary>
    Private Function TryRefreshTokenFromJar() As String
        LastTokenRefreshError = ""
        LastTokenRefreshBaseUrl = GetBaseUrl()
        LastTokenRefreshJarPath = ""
        LastTokenRefreshUsernameUsed = ""
        ' Credentials: hardcoded ExemplarApiUsername / ExemplarApiPassword; env vars override if set
        Dim username As String = Environment.GetEnvironmentVariable("EXEMPLAR_API_USERNAME")?.Trim()
        Dim password As String = Environment.GetEnvironmentVariable("EXEMPLAR_API_PASSWORD")
        If String.IsNullOrWhiteSpace(username) Then username = ExemplarApiUsername?.Trim()
        If String.IsNullOrWhiteSpace(password) Then password = ExemplarApiPassword
        LastTokenRefreshUsernameUsed = If(username, "").Trim()
        If String.IsNullOrWhiteSpace(username) OrElse String.IsNullOrWhiteSpace(password) Then
            LastTokenRefreshError = "Username or password not set."
            Return ""
        End If

        Dim jarPath As String = GetLoginJarPath()
        LastTokenRefreshJarPath = jarPath
        If String.IsNullOrWhiteSpace(jarPath) Then
            LastTokenRefreshError = "Login JAR not found (add ExemplarLogin.jar or ExemplarLoginDev.jar to the app folder, or set EXEMPLAR_LOGIN_JAR_PATH)."
            Return ""
        End If

        SyncLock TokenRefreshLock
            Try
                Dim javaExe As String = GetJavaExecutablePath()
                If String.IsNullOrWhiteSpace(javaExe) Then
                    LastTokenRefreshError = "Java not found. Install Java or add a jre folder next to the app."
                    Return ""
                End If
                Using proc As New Process()
                    proc.StartInfo.FileName = javaExe
                    proc.StartInfo.Arguments = "-jar """ & jarPath & """"
                    proc.StartInfo.UseShellExecute = False
                    proc.StartInfo.RedirectStandardOutput = True
                    proc.StartInfo.RedirectStandardError = True
                    proc.StartInfo.CreateNoWindow = True
                    proc.StartInfo.StandardOutputEncoding = Encoding.UTF8
                    proc.StartInfo.EnvironmentVariables("username") = username
                    proc.StartInfo.EnvironmentVariables("password") = password
                    Try
                        proc.Start()
                    Catch ex As System.ComponentModel.Win32Exception When ex.NativeErrorCode = 2
                        LastTokenRefreshError = "Java not found. Install Java (e.g. from https://adoptium.net), add it to PATH or set JAVA_HOME, or place a jre folder next to this app."
                        Return ""
                    End Try
                    Dim output As String = proc.StandardOutput.ReadToEnd()
                    Dim errOut As String = proc.StandardError.ReadToEnd()
                    proc.WaitForExit(15000)
                    If Not proc.HasExited Then
                        LastTokenRefreshError = "Login JAR timed out."
                        Return ""
                    End If
                    If proc.ExitCode <> 0 Then
                        LastTokenRefreshError = "Login JAR failed (exit " & proc.ExitCode.ToString() & "). " & If(String.IsNullOrWhiteSpace(errOut), "Check username/password and staging vs production.", errOut.Trim())
                        Return ""
                    End If
                    ' Take first line only and strip any newlines (header values must not contain CR/LF)
                    Dim token As String = SanitizeTokenForHeader(output)
                    ' The login JAR may output either the raw token or "Bearer <token>".
                    ' Normalize so callers can always send it as: Authorization: Bearer <token>
                    If token.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase) Then
                        token = token.Substring("Bearer ".Length).Trim()
                    End If
                    ' If auth failed, the jar may print an error line (e.g. "Auth failed....")
                    ' which we must not treat as a bearer token.
                    If String.IsNullOrWhiteSpace(token) _
                        OrElse token.Equals("null", StringComparison.OrdinalIgnoreCase) _
                        OrElse token.StartsWith("Auth failed", StringComparison.OrdinalIgnoreCase) _
                        OrElse token.IndexOf("Exception", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        ' Keep the raw jar line as the error detail.
                        LastTokenRefreshError = "Login JAR authentication failed: " & token
                        Return ""
                    End If

                    ' Basic sanity-check: most bearer tokens are JWT-like (contain 2 dots).
                    ' If it doesn't look like a token, treat it as failure.
                    Dim dotCount As Integer = token.Count(Function(ch) ch = "."c)
                    If token.Length < 20 OrElse dotCount < 2 Then
                        LastTokenRefreshError = "Login JAR returned unexpected output (not a token). FirstLine=" & token
                        Return ""
                    End If
                    If Not String.IsNullOrWhiteSpace(token) Then
                        Return token
                    End If
                    LastTokenRefreshError = "Login JAR returned no token. Check username/password and that you use the staging JAR for staging."
                End Using
            Catch ex As Exception
                LastTokenRefreshError = "Token refresh error: " & ex.Message
                Return ""
            End Try
        End SyncLock

        Return ""
    End Function

    ''' <summary>Returns path to java.exe: bundled JRE (app dir or parent dirs), then JAVA_HOME, then common install folders (Java 25, Microsoft, Adoptium, etc.), then "java" (PATH).</summary>
    Private Function GetJavaExecutablePath() As String
        Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
        ' 1. Bundled JRE: next to exe, then parent folders (so jre in project root is found when running from bin\Debug\net8.0-windows)
        Dim dirsToCheck As New List(Of String) From {baseDir}
        Dim parent As String = Path.GetDirectoryName(baseDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
        For i As Integer = 0 To 4
            If String.IsNullOrEmpty(parent) Then Exit For
            dirsToCheck.Add(parent)
            parent = Path.GetDirectoryName(parent)
        Next
        For Each dir As String In dirsToCheck
            If String.IsNullOrEmpty(dir) OrElse Not Directory.Exists(dir) Then Continue For
            For Each rel As String In New String() {"jre\bin\java.exe", "runtime\bin\java.exe", "jdk\bin\java.exe"}
                Dim exe As String = Path.Combine(dir, rel)
                If File.Exists(exe) Then Return exe
            Next
        Next
        ' 2. JAVA_HOME
        Dim javaHome As String = Environment.GetEnvironmentVariable("JAVA_HOME")?.Trim()
        If Not String.IsNullOrEmpty(javaHome) Then
            Dim exe As String = Path.Combine(javaHome, "bin", "java.exe")
            If File.Exists(exe) Then Return exe
            exe = Path.Combine(javaHome, "jre", "bin", "java.exe")
            If File.Exists(exe) Then Return exe
        End If
        ' 3. Common install locations (Java folder, Microsoft OpenJDK, Eclipse Adoptium - includes Java 25)
        For Each base As String In New String() {
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
        }
            For Each folderName As String In New String() {"Java", "Microsoft", "Eclipse Adoptium", "Adoptium", "Amazon Corretto"}
                Dim javaDir As String = Path.Combine(base, folderName)
                If Not Directory.Exists(javaDir) Then Continue For
                Try
                    For Each subDir As String In Directory.GetDirectories(javaDir)
                        Dim exe As String = Path.Combine(subDir, "bin", "java.exe")
                        If File.Exists(exe) Then Return exe
                        exe = Path.Combine(subDir, "jre", "bin", "java.exe")
                        If File.Exists(exe) Then Return exe
                    Next
                Catch
                    ' Ignore access or path errors
                End Try
            Next
        Next
        ' 4. Fall back to system Java (must be on PATH)
        Return "java"
    End Function

    ''' <summary>True when EXEMPLAR_API_BASE_URL is set (staging/dev); then app uses ExemplarLoginDev.jar if present.</summary>
    Private Function IsStagingApi() As Boolean
        Dim baseUrl As String = ""
        If Not String.IsNullOrWhiteSpace(My.Settings.ExemplarApiBaseUrl) Then
            baseUrl = My.Settings.ExemplarApiBaseUrl
        Else
            baseUrl = Environment.GetEnvironmentVariable("EXEMPLAR_API_BASE_URL")
        End If

        If String.IsNullOrWhiteSpace(baseUrl) Then Return False

        ' Only treat it as staging when the effective base URL actually indicates staging/dev.
        ' This avoids using the staging JAR for production API tokens.
        Dim v As String = baseUrl.Trim().ToLowerInvariant()
        Return v = "staging" OrElse v.Contains("staging") OrElse v.Contains("dev")
    End Function

    ''' <summary>Returns path to the login JAR: embedded (ExemplarLogin.jar = production, ExemplarLoginDev.jar = staging), then EXEMPLAR_LOGIN_JAR_PATH, then app directory. Creates temp file if extracted from assembly.</summary>
    Private Function GetLoginJarPath() As String
        Dim useStagingJar As Boolean = IsStagingApi()

        ' 1. Embedded JAR: prefer ExemplarLoginDev.jar when staging, ExemplarLogin.jar when production
        Dim asm As Assembly = Assembly.GetExecutingAssembly()
        Dim resourceName As String = Nothing
        Dim preferredSuffix As String = If(useStagingJar, "ExemplarLoginDev.jar", "ExemplarLogin.jar")
        Dim fallbackSuffix As String = If(useStagingJar, "ExemplarLogin.jar", "ExemplarLoginDev.jar")
        For Each name As String In asm.GetManifestResourceNames()
            If Not name.EndsWith(".jar", StringComparison.OrdinalIgnoreCase) Then Continue For
            If name.EndsWith(preferredSuffix, StringComparison.OrdinalIgnoreCase) Then
                resourceName = name
                Exit For
            End If
            If resourceName Is Nothing AndAlso (name.EndsWith(fallbackSuffix, StringComparison.OrdinalIgnoreCase) OrElse name.Contains("ExemplarLogin") OrElse name.Contains("eprofiling")) Then
                resourceName = name
            End If
        Next
        If Not String.IsNullOrEmpty(resourceName) Then
            Try
                Dim tempDir As String = Path.Combine(Path.GetTempPath(), "ExemplarProfiling")
                Directory.CreateDirectory(tempDir)
                Dim tempFileName As String = If(useStagingJar, "ExemplarLoginDev.jar", "ExemplarLogin.jar")
                Dim tempPath As String = Path.Combine(tempDir, tempFileName)
                Using stream As Stream = asm.GetManifestResourceStream(resourceName)
                    If stream IsNot Nothing Then
                        Using fs As FileStream = File.Create(tempPath)
                            stream.CopyTo(fs)
                        End Using
                        Return tempPath
                    End If
                End Using
            Catch
                ' Fall through to external JAR
            End Try
        End If

        ' 2. Explicit path from env
        Dim envPath As String = Environment.GetEnvironmentVariable("EXEMPLAR_LOGIN_JAR_PATH")?.Trim()
        If Not String.IsNullOrWhiteSpace(envPath) AndAlso File.Exists(envPath) Then
            Return envPath
        End If

        ' 3. App directory, then parent folders (so JAR in project root is found when running from bin\Debug\net8.0-windows)
        Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
        Dim dirsToCheck As New List(Of String) From {baseDir}
        Dim parent As String = Path.GetDirectoryName(baseDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
        For i As Integer = 0 To 4
            If String.IsNullOrEmpty(parent) Then Exit For
            dirsToCheck.Add(parent)
            parent = Path.GetDirectoryName(parent)
        Next
        For Each dir As String In dirsToCheck
            If String.IsNullOrEmpty(dir) OrElse Not Directory.Exists(dir) Then Continue For
            Dim devPath As String = Path.Combine(dir, "ExemplarLoginDev.jar")
            Dim prodPath As String = Path.Combine(dir, "ExemplarLogin.jar")
            If useStagingJar Then
                If File.Exists(devPath) Then Return devPath
                If File.Exists(prodPath) Then Return prodPath
            Else
                If File.Exists(prodPath) Then Return prodPath
                If File.Exists(devPath) Then Return devPath
            End If
            Dim legacyStaging As String = Path.Combine(dir, "eprofiling-user-login-1.0-staging.jar")
            Dim legacyProd As String = Path.Combine(dir, "eprofiling-user-login-1.0.jar")
            If File.Exists(legacyStaging) Then Return legacyStaging
            If File.Exists(legacyProd) Then Return legacyProd
        Next

        Return ""
    End Function

    ''' <summary>Clears the cached token so the next call will try env or JAR again. Call when API returns 401.</summary>
    Public Sub ClearCachedToken()
        SyncLock TokenRefreshLock
            CachedToken = Nothing
        End SyncLock
    End Sub

    Private Function Normalize(value As String) As String
        If value Is Nothing Then
            Return ""
        End If

        Return value.Trim().ToLowerInvariant()
    End Function
End Module
