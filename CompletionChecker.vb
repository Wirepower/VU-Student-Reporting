'Imports System.Data.SqlClient
Imports Microsoft.Data.SqlClient

Module CompletionChecker
    Public columnsToCheck() As String = {"UEECO0023", "UEECD0007", "UEECD0019", "UEECD0020", "UEECD0051", "UEECD0046", "UEECD0044", "UEEEL0021", "UEEEL0019", "UEERE0001", "UEEEL0023", "UEEEL0020", "UEEEL0025", "UEEEL0024", "UEEEL0008", "UEEEL0009", "UEEEL0010", "UEEDV0005", "UEEDV0008", "UEEEL0003", "UEEEL0018", "UEEEL0005", "UEECD0016", "UEEEL0047", "HLTAID009", "UETDRRF004", "UEEEL0014", "UEEEL0012", "UEEEL0039"}

    Public correctOrder() As String = {"UEECO0023", "UEECD0007", "UEECD0019", "UEECD0020", "UEECD0051", "UEECD0046", "UEECD0044", "UEEEL0021", "UEEEL0019", "UEERE0001", "UEEEL0023", "UEEEL0020", "UEEEL0025", "UEEEL0024", "UEEEL0008", "UEEEL0009", "UEEEL0010", "UEEDV0005", "UEEDV0008", "UEEEL0003", "UEEEL0018", "UEEEL0005", "UEECD0016", "UEEEL0047", "HLTAID009", "UETDRRF004", "UEEEL0014", "UEEEL0012", "UEEEL0039"}

    Public Sub UpdateLabelsFromDatabase(studentID As String)
        ' Connection string to your database
        Dim connectionString As String = "Server=DEVSQLCENTRAL.AD.VU.EDU.AU;Integrated Security=True;Connect Timeout=30;Encrypt=True;Trust Server Certificate=True;Application Intent=ReadWrite;Multi Subnet Failover=False;"

        ' Query to retrieve the data for the specified student
        Dim query As String = "SELECT * FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@StudentID", studentID)
                    Dim reader As SqlDataReader = command.ExecuteReader()

                    ' Check if the student exists in the database
                    If reader.Read() Then
                        ' Extract the unit columns from the database reader
                        Dim checkedUnits As New List(Of String)()

                        For Each columnName As String In columnsToCheck
                            Dim columnIndex As Integer = reader.GetOrdinal(columnName)
                            Dim value As Boolean = Convert.ToBoolean(reader(columnIndex))
                            If value Then
                                checkedUnits.Add(columnName)
                            End If
                        Next

                        ' Check if all columns are set to true
                        Dim allCompleted As Boolean = True
                        For Each column As String In columnsToCheck
                            Dim columnIndex As Integer = reader.GetOrdinal(column)
                            Dim value As Boolean = Convert.ToBoolean(reader(columnIndex))
                            If Not value Then
                                allCompleted = False
                                Exit For
                            End If
                        Next
                        MainFrm.UnitAlertLbl.Text = ""
                        StudentUnits.UnitAlertLbl1.Text = ""

                        ' Check if the checked units are in the correct order
                        Dim allSequential As Boolean = CheckSequentialUnits(checkedUnits, correctOrder)

                        ' Update labels based on the analysis
                        If allCompleted Then
                            MainFrm.UnitAlertLbl.Text = "UEE30820 - Completed"
                            StudentUnits.UnitAlertLbl1.Text = "UEE30820 - Completed"
                        ElseIf checkedUnits.Count = 0 Then
                            MainFrm.UnitAlertLbl.Text = ""
                            StudentUnits.UnitAlertLbl1.Text = ""

                        ElseIf Not allSequential Then
                            MainFrm.UnitAlertLbl.Text = "Student Is missing pre-requisite units"
                            StudentUnits.UnitAlertLbl1.Text = "Student Is missing pre-requisite units"
                        End If
                    Else
                        ' Student not found in the database
                        MainFrm.UnitAlertLbl.Text = ""
                        StudentUnits.UnitAlertLbl1.Text = ""
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error updating labels from database: " & ex.Message)
        End Try
    End Sub

    Public Function CheckSequentialUnits(units As List(Of String), correctOrder As String()) As Boolean
        Dim index As Integer = 0

        ' Check if the checked units are in the correct order
        For Each unit As String In units
            If unit <> correctOrder(index) Then
                Return False
            End If
            index += 1
        Next

        Return True
    End Function


    Private Sub InsertNewStudentRow(studentID As String)
        ' Connection string to your database
        Dim connectionString As String = "Server=DEVSQLCENTRAL.AD.VU.EDU.AU;Integrated Security=True;Connect Timeout=30;Encrypt=True;Trust Server Certificate=True;Application Intent=ReadWrite;Multi Subnet Failover=False;"

        ' Query to insert a new row for the student
        Dim query As String = "INSERT INTO ElectrotechnologyReports.dbo.StudentLogs ([Student ID]) VALUES (@StudentID)"

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@StudentID", studentID)
                    command.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error inserting new student row: " & ex.Message)
        End Try
    End Sub

    Public Sub LoadCheckBoxStates(studentID As String)
        ' Connection string to your database
        Dim connectionString As String = "Server=DEVSQLCENTRAL.AD.VU.EDU.AU;Integrated Security=True;Connect Timeout=30;Encrypt=True;Trust Server Certificate=True;Application Intent=ReadWrite;Multi Subnet Failover=False;"

        ' Query to retrieve the data for the specified student
        Dim query As String = "SELECT UEECO0023, UEECD0007, UEECD0019, UEECD0020, UEECD0051, UEECD0046, UEECD0044, UEEEL0021, UEEEL0019, UEERE0001, UEEEL0023, UEEEL0020, UEEEL0025, UEEEL0024, UEEEL0008, UEEEL0009, UEEEL0010, UEEDV0005, UEEDV0008, UEEEL0003, UEEEL0018, UEEEL0005, UEECD0016, UEEEL0047, HLTAID009, UETDRRF004, UEEEL0012, UEEEL0014, UEEEL0039 FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@StudentID", studentID)
                    Dim reader As SqlDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        ' Student ID exists, update checkbox states based on database values
                        For i As Integer = 0 To columnsToCheck.Length - 1
                            Dim columnName As String = columnsToCheck(i)
                            Dim checkboxName As String = "checkbox" & (i + 1)
                            Dim checkbox As CheckBox = TryCast(StudentUnits.Controls(checkboxName), CheckBox)

                            If checkbox IsNot Nothing Then
                                Dim columnIndex As Integer = reader.GetOrdinal(columnName)
                                checkbox.Checked = Convert.ToBoolean(reader(columnIndex))
                            End If
                        Next
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading checkbox states: " & ex.Message)
        End Try
    End Sub
    Public Sub UpdateDatabase(studentID As String, columnName As String, newValue As Boolean)
        ' Connection string to your database
        Dim connectionString As String = "Server=DEVSQLCENTRAL.AD.VU.EDU.AU;Integrated Security=True;Connect Timeout=30;Encrypt=True;Trust Server Certificate=True;Application Intent=ReadWrite;Multi Subnet Failover=False;"

        ' Check if the student ID exists in the database
        Dim queryCheckExistence As String = "SELECT COUNT(*) FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"
        Dim studentExists As Boolean = False

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Using command As New SqlCommand(queryCheckExistence, connection)
                    command.Parameters.AddWithValue("@StudentID", studentID)
                    Dim count As Integer = Convert.ToInt32(command.ExecuteScalar())
                    If count > 0 Then
                        studentExists = True
                    End If
                End Using
            End Using

            If Not studentExists Then
                ' Insert a new row for the student
                InsertNewStudentRow(studentID)
            End If

            ' Query to update the database
            Dim queryUpdate As String = "UPDATE ElectrotechnologyReports.dbo.StudentLogs SET " & columnName & " = @Value WHERE [Student ID] = @StudentID"

            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Using command As New SqlCommand(queryUpdate, connection)
                    command.Parameters.AddWithValue("@Value", newValue)
                    command.Parameters.AddWithValue("@StudentID", studentID)
                    command.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error updating database: " & ex.Message)
        End Try

    End Sub
End Module




