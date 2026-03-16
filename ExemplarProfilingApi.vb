Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
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
End Class

Public Class ExemplarQualificationUpdateResult
    Public Property IsSuccessful As Boolean
    Public Property ErrorMessage As String
    Public Property ResponseBody As String
End Class

Module ExemplarProfilingApi
    Private Const DefaultBaseUrl As String = "https://api.profiling.exemplarsystems.com.au"
    Private ReadOnly ApiClient As New HttpClient() With {
        .Timeout = TimeSpan.FromSeconds(20)
    }
    Private CachedToken As String

    Private Class ExemplarUserCandidate
        Public Property Id As String
        Public Property FirstName As String
        Public Property LastName As String
        Public Property Email As String
    End Class

    Public Function IsConfigured() As Boolean
        Return Not String.IsNullOrWhiteSpace(GetBearerToken())
    End Function

    Public Sub SetBearerToken(token As String)
        If String.IsNullOrWhiteSpace(token) Then
            CachedToken = Nothing
            Return
        End If

        Dim normalized As String = token.Trim()
        If normalized.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase) Then
            normalized = normalized.Substring("Bearer ".Length).Trim()
        End If

        CachedToken = normalized
    End Sub

    Public Async Function LookupStudentProfileAsync(firstName As String, lastName As String, email As String) As Task(Of ExemplarProfileLookupResult)
        Dim token As String = GetBearerToken()
        If String.IsNullOrWhiteSpace(token) Then
            Return New ExemplarProfileLookupResult With {
                .IsSuccessful = False,
                .IsConfigured = False,
                .StatusText = "Not configured",
                .DetailText = "Set environment variable EXEMPLAR_API_TOKEN to enable profiling lookups."
            }
        End If

        Try
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
                Dim total As Integer? = FindFirstInteger(cardsDoc.RootElement, New String() {"total_cards", "totalCards", "card_count", "cards_count", "total"})
                Dim completed As Integer? = FindFirstInteger(cardsDoc.RootElement, New String() {"completed_cards", "completedCards", "completed"})
                Dim pending As Integer? = FindFirstInteger(cardsDoc.RootElement, New String() {"pending_cards", "pendingCards", "pending", "incomplete"})
                Dim status As String = FindFirstString(cardsDoc.RootElement, New String() {"status", "qualification_status", "progress_status"})

                Dim detail As String
                If total.HasValue AndAlso completed.HasValue Then
                    detail = $"{completed.Value}/{total.Value} cards completed"
                ElseIf total.HasValue Then
                    detail = $"Total cards: {total.Value}"
                Else
                    detail = "Card summary loaded."
                End If

                If pending.HasValue Then
                    detail &= $" | Pending: {pending.Value}"
                End If

                If Not String.IsNullOrWhiteSpace(status) Then
                    detail &= $" | Status: {status}"
                End If

                Return New ExemplarProfileLookupResult With {
                    .IsSuccessful = True,
                    .IsConfigured = True,
                    .UserId = student.Id,
                    .StatusText = "Connected",
                    .DetailText = detail,
                    .TotalCards = total,
                    .CompletedCards = completed,
                    .PendingCards = pending
                }
            End Using
        Catch ex As Exception
            Return New ExemplarProfileLookupResult With {
                .IsSuccessful = False,
                .IsConfigured = True,
                .StatusText = "Error",
                .DetailText = "Unable to retrieve profiling summary.",
                .ErrorMessage = ex.Message
            }
        End Try
    End Function

    Public Async Function UpdateQualificationStatusAsync(userId As String, qualificationId As String, statusValue As String) As Task(Of ExemplarQualificationUpdateResult)
        Dim token As String = GetBearerToken()
        If String.IsNullOrWhiteSpace(token) Then
            Return New ExemplarQualificationUpdateResult With {
                .IsSuccessful = False,
                .ErrorMessage = "EXEMPLAR_API_TOKEN is not configured."
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

        Using doc As JsonDocument = Await SendJsonRequestAsync(HttpMethod.Get, searchUrl, token, Nothing)
            Dim candidates As New List(Of ExemplarUserCandidate)()
            CollectUserCandidates(doc.RootElement, candidates, 0)
            Return SelectBestCandidate(candidates, firstName, lastName, email)
        End Using
    End Function

    Private Async Function SendJsonRequestAsync(method As HttpMethod, url As String, token As String, payload As String) As Task(Of JsonDocument)
        Using request As New HttpRequestMessage(method, url)
            request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", token)
            request.Headers.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
            If payload IsNot Nothing Then
                request.Content = New StringContent(payload, Encoding.UTF8, "application/json")
            End If

            Using response As HttpResponseMessage = Await ApiClient.SendAsync(request).ConfigureAwait(False)
                Dim body As String = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                response.EnsureSuccessStatusCode()
                Return JsonDocument.Parse(body)
            End Using
        End Using
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

    Private Function GetStringByPossibleKeys(element As JsonElement, keys As IEnumerable(Of String)) As String
        For Each key In keys
            Dim prop As JsonElement
            If element.TryGetProperty(key, prop) Then
                Return TryParseString(prop)
            End If
        Next

        Return ""
    End Function

    Private Function GetBaseUrl() As String
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

        Dim token As String = Environment.GetEnvironmentVariable("EXEMPLAR_API_TOKEN")
        If String.IsNullOrWhiteSpace(token) Then
            token = Environment.GetEnvironmentVariable("EXEMPLAR_BEARER_TOKEN")
        End If

        If String.IsNullOrWhiteSpace(token) Then
            Return ""
        End If

        SetBearerToken(token)
        Return CachedToken
    End Function

    Private Function Normalize(value As String) As String
        If value Is Nothing Then
            Return ""
        End If

        Return value.Trim().ToLowerInvariant()
    End Function
End Module
