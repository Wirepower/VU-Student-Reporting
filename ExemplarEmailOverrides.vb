Imports Microsoft.Data.SqlClient

''' <summary>
''' In-memory cache plus persistence in ElectrotechnologyReports.dbo.ExemplarProfilingStudentDB
''' (StudentID, ExemplarProfilingEmail).
''' </summary>
Module ExemplarEmailOverrides
    Private Const ProfilingEmailTable As String = "ElectrotechnologyReports.dbo.ExemplarProfilingStudentDB"

    Private ReadOnly OverridesByStudentId As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    ''' <summary>Student IDs for which we already queried SQL and found no usable row.</summary>
    Private ReadOnly DbMissStudentIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    Public Function GetOverride(studentId As String) As String
        If String.IsNullOrWhiteSpace(studentId) Then
            Return ""
        End If

        Dim key As String = studentId.Trim()
        If OverridesByStudentId.ContainsKey(key) Then
            Return OverridesByStudentId(key)
        End If

        If DbMissStudentIds.Contains(key) Then
            Return ""
        End If

        Dim loadErr As Boolean
        Dim fromDb As String = TryLoadEmailFromDatabase(key, loadErr)
        If loadErr Then
            Return ""
        End If
        If Not String.IsNullOrWhiteSpace(fromDb) Then
            OverridesByStudentId(key) = fromDb
            Return fromDb
        End If

        DbMissStudentIds.Add(key)
        Return ""
    End Function

    ''' <summary>Clears cached lookup so the next GetOverride reads SQL again (e.g. after switching students).</summary>
    Public Sub InvalidateCacheForStudent(studentId As String)
        If String.IsNullOrWhiteSpace(studentId) Then
            Return
        End If
        Dim key As String = studentId.Trim()
        If OverridesByStudentId.ContainsKey(key) Then
            OverridesByStudentId.Remove(key)
        End If
        DbMissStudentIds.Remove(key)
    End Sub

    Public Sub SetOverride(studentId As String, email As String)
        If String.IsNullOrWhiteSpace(studentId) Then
            Return
        End If

        Dim key As String = studentId.Trim()
        Dim normalized As String = If(email, "").Trim()

        If String.IsNullOrWhiteSpace(normalized) Then
            If OverridesByStudentId.ContainsKey(key) Then
                OverridesByStudentId.Remove(key)
            End If
            DbMissStudentIds.Add(key)
            TryDeleteFromDatabase(key)
            Return
        End If

        OverridesByStudentId(key) = normalized
        DbMissStudentIds.Remove(key)
        TryUpsertDatabase(key, normalized)
    End Sub

    Private Function TryLoadEmailFromDatabase(studentId As String, ByRef hadSqlError As Boolean) As String
        hadSqlError = False
        Try
            Const sql As String = "SELECT ExemplarProfilingEmail FROM " & ProfilingEmailTable & " WHERE StudentID = @StudentID"
            Using connection As New SqlConnection(SQLCon.connectionString)
                Using command As New SqlCommand(sql, connection)
                    command.Parameters.AddWithValue("@StudentID", studentId)
                    connection.Open()
                    Dim result As Object = command.ExecuteScalar()
                    If result Is Nothing OrElse result Is DBNull.Value Then
                        Return ""
                    End If
                    Dim s As String = Convert.ToString(result).Trim()
                    Return s
                End Using
            End Using
        Catch ex As Exception
            hadSqlError = True
            Try
                Dim debugPath As String = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExemplarDbOverrideErrors.txt")
                System.IO.File.AppendAllText(debugPath,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} LOAD failed for StudentID='{studentId}': {ex.Message}" & Environment.NewLine)
            Catch
                ' Ignore debug logging failures.
            End Try
            Return ""
        End Try
    End Function

    Private Sub TryUpsertDatabase(studentId As String, email As String)
        Try
            Dim sql As String =
                "MERGE " & ProfilingEmailTable & " AS T " &
                "USING (SELECT @StudentID AS StudentID) AS S " &
                "ON T.StudentID = S.StudentID " &
                "WHEN MATCHED THEN UPDATE SET ExemplarProfilingEmail = @Email " &
                "WHEN NOT MATCHED BY TARGET THEN INSERT (StudentID, ExemplarProfilingEmail) VALUES (@StudentID, @Email);"
            Using connection As New SqlConnection(SQLCon.connectionString)
                Using command As New SqlCommand(sql, connection)
                    command.Parameters.AddWithValue("@StudentID", studentId)
                    command.Parameters.AddWithValue("@Email", email)
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End Using
        Catch
            ' Non-fatal: in-memory override still applies for this session.
        End Try
    End Sub

    Private Sub TryDeleteFromDatabase(studentId As String)
        Try
            Const sql As String = "DELETE FROM " & ProfilingEmailTable & " WHERE StudentID = @StudentID"
            Using connection As New SqlConnection(SQLCon.connectionString)
                Using command As New SqlCommand(sql, connection)
                    command.Parameters.AddWithValue("@StudentID", studentId)
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End Using
        Catch
        End Try
    End Sub
End Module
