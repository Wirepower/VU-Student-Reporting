Option Explicit On
Imports Microsoft.Data.SqlClient
'Imports MyNamespace
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Net
Imports System.Net.Mail
Imports Student_Attendance_Reporting
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.SharePoint.Client
Imports Microsoft.SharePoint.ApplicationPages
Imports Azure
Imports Microsoft.SharePoint.Mobile.Controls
Imports Microsoft.Office.Server.UserProfiles

Module SendOutlookEmail
    Dim emailedstudent As String
    Dim rangstudent As String
    Dim othercontact As String
    Dim othertext As String
    Dim LastStudentReportDate As String

    ''' <summary>Inserted between SQL EmailBody and JPEG signature when template is Exemplar Profiling Outstanding Alert.</summary>
    Private Function BuildExemplarProfilingSummaryHtml() As String
        Dim missing As String = If(String.IsNullOrWhiteSpace(MainFrm.ProfilingMissingValLbl.Text), "?", WebUtility.HtmlEncode(MainFrm.ProfilingMissingValLbl.Text))
        Dim notVerified As String = If(String.IsNullOrWhiteSpace(MainFrm.ProfilingNotVerifiedValLbl.Text), "?", WebUtility.HtmlEncode(MainFrm.ProfilingNotVerifiedValLbl.Text))
        Dim employerVerified As String = If(String.IsNullOrWhiteSpace(MainFrm.ProfilingEmployerVerifiedValLbl.Text), "?", WebUtility.HtmlEncode(MainFrm.ProfilingEmployerVerifiedValLbl.Text))
        Dim lastCard As String = If(String.IsNullOrWhiteSpace(MainFrm.ProfilingLastCardValLbl.Text), "?", WebUtility.HtmlEncode(MainFrm.ProfilingLastCardValLbl.Text))
        Return "<BR><BR><b>Exemplar profiling (API)</b><BR>" &
            "Cards not submitted/Outstanding: " & missing & "<BR>" &
            "Cards Submitted (Not verified): " & notVerified & "<BR>" &
            "Cards submitted (Employer Verified): " & employerVerified & "<BR>" &
            "Last Card Submission: " & lastCard & "<BR>"
    End Function

    Sub SendOutlookEmail(studentID As Double, studentFirstname As String, studentSurname As String, studentEmail As String, employerFirstname As String, employerSurname As String, employerBusinessName As String, employerEmail As String)
        Dim OutApp As Object
        Dim OutMail As Object
        Dim body As String
        Dim AdminEmail As String
        Dim ApptrainEmail As String
        Dim FrankOffer As String
        Dim Trades As String

        ' Retrieve email addresses from the EmailSettings table
        GetEmailAddresses(AdminEmail, ApptrainEmail, FrankOffer, Trades)

        ' Determine the template based on the selected value in ComboBox1
        Dim selectedTemplate As String = MainFrm.ComboBox12.Text

        ' Get the email body for the selected template
        body = GenerateEmailTemplateSQL(selectedTemplate, studentID, studentFirstname, studentSurname, studentEmail, employerFirstname, employerSurname, employerBusinessName, employerEmail)
        If selectedTemplate = "Exemplar Profiling Outstanding Alert" Then
            body &= BuildExemplarProfilingSummaryHtml()
        End If
        ' Create a new instance of Outlook application

        ' Get the image data from the database
        Dim imageData As Byte() = RetrieveImageDataFromDatabase()

        OutApp = CreateObject("Outlook.Application")
        ' Create a new email item
        OutMail = OutApp.CreateItem(0)
        ' Set email properties
        With OutMail
            If MainFrm.ComboBox12.Text = "Student Investigation" Then
                .To = ApptrainEmail
                .cc = AdminEmail
                .bcc = ""
                .Subject = MainFrm.ComboBox12.Text
                .HTMLbody = body & "<br><img src='data:image/jpeg;base64," & Convert.ToBase64String(imageData) & "' width='90%'> " & MainFrm.VersionLBL.Text
                .Display ' Display the email
            Else
                .To = employerEmail
                .cc = studentEmail
                .bcc = AdminEmail
                .Subject = MainFrm.ComboBox12.Text
                .HTMLbody = body & "<br><img src='data:image/jpeg;base64," & Convert.ToBase64String(imageData) & "' width='90%'> " & MainFrm.VersionLBL.Text
                .Display ' Display the email
            End If
            '-------------------------------------------------------------------
            ' Assuming you have a valid connection string defined earlier in your code
            Dim connectionString As String = SQLCon.connectionString

            ' Create a new SqlConnection object
            Using connection As New SqlConnection(connectionString)
                ' Open the connection
                connection.Open()

                If MainFrm.ComboBox12.Text = "Student Term Progress Report" Or MainFrm.ComboBox12.Text = "Yearly Student Report" Then
                    Try
                        ' Check if the student ID exists in the table
                        Dim queryCheck As String = "SELECT COUNT(*) FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"

                        Using commandCheck As New SqlCommand(queryCheck, connection)
                            ' Assuming 'studentID' is defined and initialized elsewhere in your code
                            commandCheck.Parameters.AddWithValue("@StudentID", studentID)
                            Dim count As Integer = CInt(commandCheck.ExecuteScalar())

                            If count > 0 Then
                                ' Student ID exists, so update the existing row
                                Dim queryUpdate As String = "UPDATE ElectrotechnologyReports.dbo.StudentLogs SET [Absent] = 0, [Late Arrival] = 0, [Early Departure] = 0 WHERE [Student ID] = @StudentID"
                                Using commandUpdate As New SqlCommand(queryUpdate, connection)
                                    commandUpdate.Parameters.AddWithValue("@StudentID", studentID)
                                    'commandUpdate.Parameters.AddWithValue("@ReportDate", MainFrm.DateTimePicker.Value.Date) ' Assuming DateTimePicker is your DateTimePicker control
                                    commandUpdate.ExecuteNonQuery()
                                    SetLastReportDateToToday(studentID)
                                End Using
                            Else
                                ' Student ID does not exist, so insert a new row
                                'Dim queryInsert As String = "INSERT INTO ElectrotechnologyReports.dbo.StudentLogs ([Student ID], [LastStudentReportDate]) VALUES (@StudentID, @TodayDate)"
                                Dim queryInsert As String = "INSERT INTO ElectrotechnologyReports.dbo.StudentLogs ([Student ID], [Absent], [Late Arrival], [Early Departure]) VALUES (@StudentID, 0, 0, 0)"
                                Using commandInsert As New SqlCommand(queryInsert, connection)
                                    commandInsert.Parameters.AddWithValue("@StudentID", studentID)
                                    'commandInsert.Parameters.AddWithValue("@ReportDate", MainFrm.DateTimePicker.Value.Date)
                                    commandInsert.ExecuteNonQuery()
                                End Using
                            End If
                        End Using

                        'MessageBox.Show("Responses written to SQL database successfully.")
                    Catch ex As Exception
                        ' Handle exceptions
                        MessageBox.Show("Error writing responses to SQL database: " & ex.Message)
                    End Try
                End If
                If MainFrm.ComboBox12.Text = "Absent Notice" Then
                    Try
                        ' Check if the student ID exists in the table
                        Dim queryCheck As String = "SELECT COUNT(*) FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"

                        Using commandCheck As New SqlCommand(queryCheck, connection)
                            ' Assuming 'studentID' is defined and initialized elsewhere in your code
                            commandCheck.Parameters.AddWithValue("@StudentID", studentID)
                            Dim count As Integer = CInt(commandCheck.ExecuteScalar())

                            If count > 0 Then
                                ' Student ID exists, so update the existing row
                                Dim queryUpdate As String = "UPDATE ElectrotechnologyReports.dbo.StudentLogs SET [Absent] = [Absent] + 1, [Yearly_Absent] = [Yearly_Absent] + 1 WHERE [Student ID] = @StudentID"
                                Using commandUpdate As New SqlCommand(queryUpdate, connection)
                                    commandUpdate.Parameters.AddWithValue("@StudentID", studentID)
                                    commandUpdate.ExecuteNonQuery()
                                End Using
                            Else
                                ' Student ID does not exist, so insert a new row
                                Dim queryInsert As String = "INSERT INTO ElectrotechnologyReports.dbo.StudentLogs ([Student ID], [Absent], [Yearly_Absent]) VALUES (@StudentID, @Absent, 1)"
                                Using commandInsert As New SqlCommand(queryInsert, connection)
                                    commandInsert.Parameters.AddWithValue("@StudentID", studentID)
                                    commandInsert.Parameters.AddWithValue("@Absent", +1) ' Assuming you want to insert today's date
                                    commandInsert.ExecuteNonQuery()
                                End Using
                            End If
                        End Using

                        ' MessageBox.Show("Responses written to SQL database successfully.")
                    Catch ex As Exception
                        ' Handle exceptions
                        MessageBox.Show("Error writing responses to SQL database: " & ex.Message)
                    End Try
                End If
                If MainFrm.ComboBox12.Text = "Early Departure Notice" Then
                    Try
                        ' Check if the student ID exists in the table
                        Dim queryCheck As String = "SELECT COUNT(*) FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"

                        Using commandCheck As New SqlCommand(queryCheck, connection)
                            ' Assuming 'studentID' is defined and initialized elsewhere in your code
                            commandCheck.Parameters.AddWithValue("@StudentID", studentID)
                            Dim count As Integer = CInt(commandCheck.ExecuteScalar())

                            If count > 0 Then
                                ' Student ID exists, so update the existing row
                                Dim queryUpdate As String = "UPDATE ElectrotechnologyReports.dbo.StudentLogs SET [Early Departure] = [Early Departure] + 1, [Yearly_Early_Departure] = [Yearly_Early_Departure] + 1 WHERE [Student ID] = @StudentID"

                                Using commandUpdate As New SqlCommand(queryUpdate, connection)
                                    commandUpdate.Parameters.AddWithValue("@StudentID", studentID)
                                    commandUpdate.ExecuteNonQuery()
                                End Using
                            Else
                                ' Student ID does not exist, so insert a new row
                                Dim queryInsert As String = "INSERT INTO ElectrotechnologyReports.dbo.StudentLogs ([Student ID], [Early Departure], [Yearly_Early_Departure]) VALUES (@StudentID, @EarlyDeparture)"
                                Using commandInsert As New SqlCommand(queryInsert, connection)
                                    commandInsert.Parameters.AddWithValue("@StudentID", studentID)
                                    commandInsert.Parameters.AddWithValue("@EarlyDeparture", +1) ' Assuming you want to insert today's date
                                    commandInsert.ExecuteNonQuery()
                                End Using
                            End If
                        End Using

                        ' MessageBox.Show("Responses written to SQL database successfully.")
                    Catch ex As Exception
                        ' Handle exceptions
                        MessageBox.Show("Error writing responses to SQL database: " & ex.Message)
                    End Try
                End If
                If MainFrm.ComboBox12.Text = "Late Arrival Notice" Then
                    Try
                        ' Check if the student ID exists in the table
                        Dim queryCheck As String = "SELECT COUNT(*) FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"

                        Using commandCheck As New SqlCommand(queryCheck, connection)
                            ' Assuming 'studentID' is defined and initialized elsewhere in your code
                            commandCheck.Parameters.AddWithValue("@StudentID", studentID)
                            Dim count As Integer = CInt(commandCheck.ExecuteScalar())

                            If count > 0 Then
                                ' Student ID exists, so update the existing row
                                Dim queryUpdate As String = "UPDATE ElectrotechnologyReports.dbo.StudentLogs SET [Late Arrival] = [Late Arrival] + 1, [Yearly_Late_Arrival] = [Yearly_Late_Arrival] + 1 WHERE [Student ID] = @StudentID"
                                Using commandUpdate As New SqlCommand(queryUpdate, connection)
                                    commandUpdate.Parameters.AddWithValue("@StudentID", studentID)
                                    commandUpdate.ExecuteNonQuery()
                                End Using
                            Else
                                ' Student ID does not exist, so insert a new row
                                Dim queryInsert As String = "INSERT INTO ElectrotechnologyReports.dbo.StudentLogs ([Student ID], [Late Arrival], [Yearly_Late_Arrival]) VALUES (@StudentID, @LateArrival)"
                                Using commandInsert As New SqlCommand(queryInsert, connection)
                                    commandInsert.Parameters.AddWithValue("@StudentID", studentID)
                                    commandInsert.Parameters.AddWithValue("@LateArrival", +1) ' Assuming you want to insert today's date
                                    commandInsert.ExecuteNonQuery()
                                End Using
                            End If
                        End Using

                        'MessageBox.Show("Responses written to SQL database successfully.")
                    Catch ex As Exception
                        ' Handle exceptions
                        MessageBox.Show("Error writing responses to SQL database: " & ex.Message)
                    End Try
                End If
            End Using

            '------------------------------------------------------------------
        End With

        ' Clean up
        OutMail = Nothing
        OutApp = Nothing
        MsgBox("Your Email has been Generated!")
    End Sub
    Public Function RetrieveImageDataFromDatabase() As Byte()
        Dim imageData As Byte() = Nothing

        ' Your SQL query to retrieve the image data
        Dim query As String = "SELECT TOP 1 [Email Signature Image] FROM ElectrotechnologyReports.dbo.EmailSettings"


        ' Define your SQL connection string
        Dim connectionString As String = SQLCon.connectionString

        ' Create a SqlConnection object
        Using connection As New SqlConnection(connectionString)
            ' Open the connection
            connection.Open()

            ' Create a SqlCommand object with your query and connection
            Using command As New SqlCommand(query, connection)
                ' Execute the query and retrieve the image data
                ' Use ExecuteScalar since you're retrieving a single value (the image data)
                imageData = DirectCast(command.ExecuteScalar(), Byte())
            End Using
        End Using

        ' Return the retrieved image data
        Return imageData
    End Function


    Public Sub GetEmailAddresses(ByRef adminEmail As String, ByRef apptrainEmail As String, ByRef FrankOffer As String, ByRef Trades As String)
        ' Define your connection string
        Dim connectionString As String = SQLCon.connectionString

        ' Define your SQL query to retrieve email addresses from the EmailSettings table
        Dim query As String = "SELECT EmailAddress FROM ElectrotechnologyReports.dbo.EmailSettings WHERE SendTo = @SendTo"

        ' Create a SqlConnection object
        Using connection As New SqlConnection(connectionString)
            ' Create a SqlCommand object with the SQL query and connection
            Using command As New SqlCommand(query, connection)
                ' Add parameters for SendTo values
                command.Parameters.AddWithValue("@SendTo", "Admin")
                connection.Open()
                '---------------------------------------------------------------
                ' Execute the SQL command and retrieve the Admin email address
                adminEmail = Convert.ToString(command.ExecuteScalar())
                '---------------------------------------------------------------
                ' Clear the parameter collection before reusing @SendTo parameter
                command.Parameters.Clear()

                ' Add parameter for Apptrain sendTo value
                command.Parameters.AddWithValue("@SendTo", "Apptrain")

                ' Execute the SQL command and retrieve the Apptrain email address
                apptrainEmail = Convert.ToString(command.ExecuteScalar())
                '---------------------------------------------------------------
                ' Clear the parameter collection before reusing @SendTo parameter
                command.Parameters.Clear()

                ' Add parameter for Frank Offer (App Creator) sendTo value
                command.Parameters.AddWithValue("@SendTo", "Frank Offer")

                ' Execute the SQL command and retrieve the Apptrain email address
                FrankOffer = Convert.ToString(command.ExecuteScalar())
                '---------------------------------------------------------------
                ' Clear the parameter collection before reusing @SendTo parameter
                command.Parameters.Clear()

                ' Add parameter for Frank Offer (App Creator) sendTo value
                command.Parameters.AddWithValue("@SendTo", "Trades")

                ' Execute the SQL command and retrieve the Apptrain email address
                Trades = Convert.ToString(command.ExecuteScalar())
                '---------------------------------------------------------------

            End Using
        End Using
    End Sub
    Private Function GetLastReportDate(studentID As String) As Date?
        Dim query As String = "SELECT LastStudentReportDate " &
                          "FROM ElectrotechnologyReports.dbo.StudentLogs " &
                          "WHERE [Student ID] = @StudentID"

        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@StudentID", studentID)

                Try
                    connection.Open()
                    Dim result As Object = command.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return DirectCast(result, Date)
                    Else
                        Return Nothing
                    End If
                Catch ex As Exception
                    MessageBox.Show("Error retrieving last report date: " & ex.Message)
                    Return Nothing
                End Try
            End Using
        End Using
    End Function




    Private Function GenerateEmailTemplateSQL(templateName As String, studentID As Double, studentFirstname As String, studentSurname As String, studentEmail As String, employerFirstname As String, employerSurname As String, employerBusinessName As String, employerEmail As String) As String
        Dim template As String = ""

        ' Database connection string
        Dim connectionString As String = SQLCon.connectionString

        ' SQL query to retrieve the email template content based on selected template name
        Dim query As String = "SELECT EmailBody FROM ElectrotechnologyReports.dbo.EmailTemplates WHERE EmailSubject = @TemplateSubject"

        ' Create a SqlConnection and SqlCommand objects
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                ' Add parameters to the SqlCommand
                command.Parameters.AddWithValue("@TemplateSubject", templateName)

                ' Open the connection
                connection.Open()

                ' Execute the query and read the template content
                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.Read() Then
                    template = reader("EmailBody").ToString()
                End If
            End Using
        End Using
        Dim lastReportDate As Date? = GetLastReportDate(studentID)
        If lastReportDate.HasValue Then
            If MainFrm.ComboBox12.Text = "Student Term Progress Report" Then
                lastReportDate = lastReportDate.Value.ToString("dddd, dd MMMM, yyyy")

            End If
        Else
            lastReportDate = Date.Today.ToString("dddd, dd MMMM, yyyy")
        End If

        ' Create a new SqlConnection object
        Using connection As New SqlConnection(connectionString)
            ' Open the connection
            connection.Open()

            If MainFrm.ComboBox12.Text = "Student Term Progress Report" Then
                Try
                    ' Check if the student ID exists in the table
                    Dim queryCheck As String = "SELECT COUNT(*) FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"

                    Using commandCheck As New SqlCommand(queryCheck, connection)
                        ' Assuming 'studentID' is defined and initialized elsewhere in your code
                        commandCheck.Parameters.AddWithValue("@StudentID", studentID)
                        Dim count As Integer = CInt(commandCheck.ExecuteScalar())

                        If count > 0 Then
                            ' Student ID exists, so update the existing row
                            Dim queryUpdate As String = "UPDATE ElectrotechnologyReports.dbo.StudentLogs SET [LastStudentReportDate] = @ReportDate WHERE [Student ID] = @StudentID"
                            Using commandUpdate As New SqlCommand(queryUpdate, connection)
                                commandUpdate.Parameters.AddWithValue("@StudentID", studentID)
                                commandUpdate.Parameters.AddWithValue("@ReportDate", lastReportDate) ' Assuming DateTimePicker1 is your DateTimePicker control
                                commandUpdate.ExecuteNonQuery()
                            End Using
                        Else
                            ' Student ID does not exist, so insert a new row
                            Dim queryInsert As String = "INSERT INTO ElectrotechnologyReports.dbo.StudentLogs ([Student ID], [LastStudentReportDate]) VALUES (@StudentID, @ReportDate)"
                            Using commandInsert As New SqlCommand(queryInsert, connection)
                                commandInsert.Parameters.AddWithValue("@StudentID", studentID)
                                commandInsert.Parameters.AddWithValue("@ReportDate", lastReportDate) ' Assuming DateTimePicker1 is your DateTimePicker control
                                commandInsert.ExecuteNonQuery()
                            End Using
                        End If
                    End Using

                    MessageBox.Show("Responses written to SQL database successfully.")
                Catch ex As Exception
                    ' Handle exceptions
                    MessageBox.Show("Error writing responses to SQL database: " & ex.Message)
                End Try
            End If

            '----
        End Using

        Call ObtainSQL()


        studentFirstname = MainFrm.StudentFirstnameLBL.Text
        studentSurname = MainFrm.StudentSurnameLBL.Text
        studentEmail = MainFrm.StudentEmailLBL.Text
        employerFirstname = MainFrm.EmployerFirstnameLBL.Text
        employerSurname = MainFrm.EmployerSurnameLBL.Text
        employerBusinessName = MainFrm.EmployerBusinessNameLBL.Text
        employerEmail = MainFrm.EmployerEmailLBL.Text

        Dim Todaydate As String = Date.Today.ToString("dddd, dd MMMM, yyyy")
        Dim TeacherName As String = MainFrm.TeacherNameLBL.Text
        Dim TeacherEmail As String = MainFrm.teacherEmailLabel.Text
        Dim Notes As String = MainFrm.NotesTB.Text

        Dim AdminEmail As String = "Electrotechnology.Admin@vu.edu.au"
        Dim ApptrainEmail As String = "Apptrain@vu.edu.au"
        Dim toEmail As String = ""
        Dim ccEmail As String = ""
        Dim bccEmail As String = ""
        Dim AttPun As String = MainFrm.ComboBox4.Text
        Dim ClassRoomEngagement As String = MainFrm.ComboBox5.Text
        Dim CourseProg As String = MainFrm.ComboBox6.Text
        Dim ReportingDate As String = MainFrm.DateTimePicker.Text
        Dim ReportingTime As String = MainFrm.TextBox1.Text
        Dim Result As String = MainFrm.ComboBox8.Text
        Dim UnitCode As String = MainFrm.Label30.Text
        Dim UnitTitle As String = MainFrm.Label31.Text
        Dim AbsentLog As String = MainFrm.Label23.Text
        Dim LateLog As String = MainFrm.Label22.Text
        Dim EarlyLog As String = MainFrm.Label19.Text
        Dim StudentID1 As Double = MainFrm.StudentIDLBL.Text


        'Dim lastReportDate As DateTime
        'Dim lastReportDate As Date? = GetLastReportDate(studentID)
        ' lastReportDate = lastReportDate.Value.ToString("dddd, dd MMMM, yyyy")
        '-----------------------------


        ' Replace placeholders with actual values
        template = template.Replace("[employerFirstname]", employerFirstname)
        template = template.Replace("[employerSurname]", employerSurname)
        template = template.Replace("[employerBusinessName]", employerBusinessName)
        template = template.Replace("[studentFirstname]", studentFirstname)
        template = template.Replace("[studentSurname]", studentSurname)

        Dim formattedLastReportDate As String = If(lastReportDate.HasValue, lastReportDate.Value.ToString("dddd, dd MMMM, yyyy"), "")
        template = template.Replace("[LastStudentReportDate]", formattedLastReportDate)

        'template = template.Replace("[LastStudentReportDate]", lastReportDate)

        template = template.Replace("[Todaydate]", Todaydate)

        template = template.Replace("[AttPun]", AttPun)
        template = template.Replace("[ClassRoomEngagement]", ClassRoomEngagement)
        template = template.Replace("[CourseProg]", CourseProg)
        template = template.Replace("[AbsentLog]", AbsentLog)
        template = template.Replace("[LateLog]", LateLog)
        template = template.Replace("[EarlyLog]", EarlyLog)
        template = template.Replace("[Notes]", Notes)
        template = template.Replace("[TeacherName]", TeacherName)
        template = template.Replace("[TeacherEmail]", TeacherEmail)
        template = template.Replace("[ReportingTime]", ReportingTime)
        template = template.Replace("[ReportingDate]", ReportingDate)
        template = template.Replace("[UnitCode]", UnitCode)
        template = template.Replace("[UnitTitle]", UnitTitle)
        template = template.Replace("[studentID]", studentID)
        template = template.Replace("[EmailedStudent]", emailedstudent)
        template = template.Replace("[RangStudent]", rangstudent)
        template = template.Replace("[OtherContact]", othercontact)
        template = template.Replace("[OtherContactText]", othertext)
        Dim missingFields As String = ""
        If MainFrm.ComboBox12.Text = "Yearly Student Report" Then
            ' Prompt the user to input the total days of school
            Dim totalSchoolDaysInput As String = InputBox("Please input the total days of school for the year:", "Total School Days")

            ' Validate the input (optional)
            If Not String.IsNullOrWhiteSpace(totalSchoolDaysInput) AndAlso IsNumeric(totalSchoolDaysInput) Then
                ' Assign the input value to the placeholder in your email template
                Dim totalSchoolDays As Integer = CInt(totalSchoolDaysInput)
                template = template.Replace("[TotalSchoolDays]", totalSchoolDays.ToString())

                ' Now you can proceed with further actions, such as sending the email with the updated template
                ' Or you can simply display a message confirming that the value has been assigned to the placeholder
                MessageBox.Show("Total school days updated successfully.")
            Else
                ' Handle invalid input or cancellation
                MessageBox.Show("Invalid input or cancellation. Please input a valid numeric value for total school days.")
            End If
        End If
        If missingFields <> "" Then
            MessageBox.Show("Please fill the following fields:" & vbCrLf & missingFields)
            Exit Function
        End If

        Dim Newquery As String = "SELECT Yearly_Early_Departure, Yearly_Absent, Yearly_Late_Arrival FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"
        ' Create a SqlConnection and SqlCommand objects
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(Newquery, connection)
                ' Add parameter for the student ID
                command.Parameters.AddWithValue("@StudentID", studentID)

                ' Open the connection
                connection.Open()

                ' Execute the query and read the result
                Using reader As SqlDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        ' Retrieve data from the reader
                        Dim studentYearlyEarlyDeparture As Integer = If(Not IsDBNull(reader("Yearly_Early_Departure")), CInt(reader("Yearly_Early_Departure")), 0)
                        Dim studentYearlyAbsent As Integer = If(Not IsDBNull(reader("Yearly_Absent")), CInt(reader("Yearly_Absent")), 0)
                        Dim studentYearlyLateArrival As Integer = If(Not IsDBNull(reader("Yearly_Late_Arrival")), CInt(reader("Yearly_Late_Arrival")), 0)

                        ' Populate the email template placeholders with retrieved data
                        template = template.Replace("[YearlyEarlyLog]", studentYearlyEarlyDeparture.ToString())
                        template = template.Replace("[YearlyAbsentLog]", studentYearlyAbsent.ToString())
                        template = template.Replace("[YearlyLateLog]", studentYearlyLateArrival.ToString())
                    Else
                        ' Handle case where no matching student log data is found
                        ' You might want to provide default values or handle this case differently
                    End If
                End Using
            End Using
        End Using

        'template = template.Replace("[studentID]", StudentID1)

        Return template
        'SetLastReportDateToToday(studentID)

    End Function
    Private Sub SetLastReportDateToToday(studentID As String)
        Dim query As String = "UPDATE ElectrotechnologyReports.dbo.StudentLogs " &
                          "SET LastStudentReportDate = GETDATE() " &
                          "WHERE [Student ID] = @StudentID"

        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@StudentID", studentID)

                Try
                    connection.Open()
                    Dim rowsAffected As Integer = command.ExecuteNonQuery()
                    If rowsAffected > 0 Then
                        'MessageBox.Show("Last student report date set to today's date successfully.")
                    Else
                        MessageBox.Show("Student ID not found.")
                    End If
                Catch ex As Exception
                    MessageBox.Show("Error setting last report date to today's date: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub

    Private Sub ObtainSQL()
        Dim connectionString As String = SQLCon.connectionString
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Assuming studentIDlbl is your label containing the Student ID

        If MainFrm.ComboBox12.Text = "Student Investigation" Then

            Dim query As String = "SELECT [Apptrain-Studentbeenemailed], [AppTrain-Haveyourangstudents], [AppTrain-OtherFormofcontact], [AppTrain-OtherText], [LastStudentReportDate] FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"

            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@StudentID", studentID)


                    connection.Open()

                    Dim reader As SqlDataReader = command.ExecuteReader()

                    If reader.Read() Then
                        emailedstudent = reader("Apptrain-Studentbeenemailed")
                        rangstudent = reader("AppTrain-Haveyourangstudents")
                        othercontact = reader("AppTrain-OtherFormofcontact")
                        othertext = reader("AppTrain-OtherText")
                        LastStudentReportDate = reader("LastStudentReportDate")


                        ' Now you can use these variables as needed
                        ' For example:
                        ' DisplayStudentData() ' Call a method to display the data
                    Else
                        ' Handle the case where no data is found for the given Student ID
                        ' For example:
                        ' MessageBox.Show("No data found for the student ID.")
                    End If

                    reader.Close()
                End Using
            End Using
            End if

    End Sub

End Module
