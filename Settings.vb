Imports Microsoft.Data.SqlClient
Imports Microsoft.VisualBasic.FileIO
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.VisualBasic.core

Public Class Settings
    Dim totalSteps As Integer = 100 ' Total number of steps
    Dim currentStep As Integer = 5  ' Current step
    ' Define maximum dimensions for allowed images
    Private Const MaxWidth As Integer = 1600
    Private Const MaxHeight As Integer = 600
    Dim dataSet As New DataSet()
    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Load initial email addresses from the database and display them in TextBoxes
        LoadCurrentSettings()
        PopulateTeacherComboBox()
        Dim connectionString As String = SQLCon.connectionString

        ' SQL query to select all data from your table
        Dim query As String = "Select *From ElectrotechnologyReports.dbo.TeacherList"

        ' Create a SqlConnection object to connect to the database
        Using connection As New SqlConnection(connectionString)
            ' Create a SqlCommand object with the SQL query and the SqlConnection
            Using command As New SqlCommand(query, connection)
                ' Create a DataTable to store the results of the SQL query
                Dim dataTable As New DataTable()

                ' Open the connection to the database
                connection.Open()

                ' Create a SqlDataAdapter to fill the DataTable with the results of the SQL query
                Using adapter As New SqlDataAdapter(command)
                    ' Fill the DataTable with the results of the SQL query
                    adapter.Fill(dataTable)
                End Using

                ' Bind the DataTable to the DataGridView
                DataGridView1.DataSource = dataTable
            End Using
        End Using
    End Sub

    Private Sub LoadEmailAddresses()
        ' Retrieve email addresses from the database
        Dim adminEmail As String = GetEmailAddress("Admin")
        Dim apptrainEmail As String = GetEmailAddress("Apptrain")
        Dim Trades As String = GetEmailAddress("Trades")

        ' Display email addresses in TextBoxes
        txtAdminEmail.Text = adminEmail
        txtApptrainEmail.Text = apptrainEmail
        TradesAdminTB.Text = Trades
    End Sub

    Private Function GetEmailAddress(sendTo As String) As String

        Dim emailAddress As String = ""

        ' Define your connection string
        Dim connectionString As String = SQLCon.connectionString

        ' SQL query to retrieve the email address based on the sendTo value
        Dim query As String = "SELECT EmailAddress FROM ElectrotechnologyReports.dbo.EmailSettings WHERE SendTo = @SendTo"



        ' Create a SqlConnection and SqlCommand objects
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                ' Add parameters to the SqlCommand
                command.Parameters.AddWithValue("@SendTo", sendTo)

                Try
                    ' Open the connection
                    connection.Open()

                    ' Execute the SQL command and retrieve the email address
                    emailAddress = Convert.ToString(command.ExecuteScalar())
                Catch ex As Exception
                    ' Handle any exceptions
                    Console.WriteLine("Error retrieving email address: " & ex.Message)
                End Try
            End Using
        End Using

        Return emailAddress
    End Function

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' Update email addresses in the database with values from TextBoxes
        UpdateEmailAddresses()
    End Sub

    Private Sub UpdateEmailAddresses()
        Dim adminEmail As String = txtAdminEmail.Text
        Dim apptrainEmail As String = txtApptrainEmail.Text
        Dim Trades As String = TradesAdminTB.Text

        ' Update email addresses in the database
        UpdateEmailAddress("Admin", adminEmail)
        UpdateEmailAddress("Apptrain", apptrainEmail)
        UpdateEmailAddress("Trades", Trades)

        MessageBox.Show("Email addresses updated successfully.")
    End Sub

    Private Sub UpdateEmailAddress(sendTo As String, emailAddress As String)
        ' Define your connection string
        Dim connectionString As String = SQLCon.connectionString

        ' SQL query to update the email address based on the sendTo value
        Dim query As String = "UPDATE ElectrotechnologyReports.dbo.EmailSettings SET EmailAddress = @EmailAddress WHERE SendTo = @SendTo"

        ' Create a SqlConnection and SqlCommand objects
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                ' Add parameters to the SqlCommand
                command.Parameters.AddWithValue("@EmailAddress", emailAddress)
                command.Parameters.AddWithValue("@SendTo", sendTo)

                Try
                    ' Open the connection
                    connection.Open()

                    ' Execute the SQL command to update the email address
                    command.ExecuteNonQuery()
                Catch ex As Exception
                    ' Handle any exceptions
                    Console.WriteLine("Error updating email address: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub
    Private Sub LoadCurrentSettings()
        ' Retrieve and display the current email settings in the form fields
        Dim adminEmail As String = GetEmailAddress("Admin")
        Dim apptrainEmail As String = GetEmailAddress("Apptrain")
        Dim Trades As String = GetEmailAddress("Trades")

        ' Update the form fields with the retrieved email addresses
        txtAdminEmail.Text = adminEmail
        txtApptrainEmail.Text = apptrainEmail
        TradesAdminTB.Text = Trades
    End Sub
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close() ' Close the form without saving changes
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MainFrm.ResetInvestigation()

    End Sub

    Private Sub UploadtoSQLdb()
        Dim connectionString As String = SQLCon.connectionString ' Assuming SQLCon is your SqlConnection object

        currentStep = 70
        ' Update progress bar to reflect current progress
        LoadingForm.UpdateProgress(currentStep)

        currentStep = 90
        ' Update progress bar to reflect current progress
        LoadingForm.UpdateProgress(currentStep)

        ' Constructing the SQL Query with MERGE statement
        Dim sqlQuery As String = "
            MERGE INTO ElectrotechnologyReports.dbo.StudentLogs AS target
            USING (SELECT [Student ID],
                          MAX(CASE WHEN units = 'UEECO0023' THEN 1 ELSE 0 END) AS UEECO0023,
                          MAX(CASE WHEN units = 'UEECD0007' THEN 1 ELSE 0 END) AS UEECD0007,
                          MAX(CASE WHEN units = 'UEECD0019' THEN 1 ELSE 0 END) AS UEECD0019,
                          MAX(CASE WHEN units = 'UEECD0020' THEN 1 ELSE 0 END) AS UEECD0020,
                          MAX(CASE WHEN units = 'UEECD0051' THEN 1 ELSE 0 END) AS UEECD0051,
                          MAX(CASE WHEN units = 'UEECD0046' THEN 1 ELSE 0 END) AS UEECD0046,
                          MAX(CASE WHEN units = 'UEECD0044' THEN 1 ELSE 0 END) AS UEECD0044,
                          MAX(CASE WHEN units = 'UEEEL0021' THEN 1 ELSE 0 END) AS UEEEL0021,
                          MAX(CASE WHEN units = 'UEEEL0019' THEN 1 ELSE 0 END) AS UEEEL0019,
                          MAX(CASE WHEN units = 'UEERE0001' THEN 1 ELSE 0 END) AS UEERE0001,
                          MAX(CASE WHEN units = 'UEEEL0023' THEN 1 ELSE 0 END) AS UEEEL0023,
                          MAX(CASE WHEN units = 'UEEEL0020' THEN 1 ELSE 0 END) AS UEEEL0020,
                          MAX(CASE WHEN units = 'UEEEL0025' THEN 1 ELSE 0 END) AS UEEEL0025,
                          MAX(CASE WHEN units = 'UEEEL0024' THEN 1 ELSE 0 END) AS UEEEL0024,
                          MAX(CASE WHEN units = 'UEEEL0008' THEN 1 ELSE 0 END) AS UEEEL0008,
                          MAX(CASE WHEN units = 'UEEEL0009' THEN 1 ELSE 0 END) AS UEEEL0009,
                          MAX(CASE WHEN units = 'UEEEL0010' THEN 1 ELSE 0 END) AS UEEEL0010,
                          MAX(CASE WHEN units = 'UEEDV0005' THEN 1 ELSE 0 END) AS UEEDV0005,
                          MAX(CASE WHEN units = 'UEEDV0008' THEN 1 ELSE 0 END) AS UEEDV0008,
                          MAX(CASE WHEN units = 'UEEEL0003' THEN 1 ELSE 0 END) AS UEEEL0003,
                          MAX(CASE WHEN units = 'UEEEL0018' THEN 1 ELSE 0 END) AS UEEEL0018,
                          MAX(CASE WHEN units = 'UEEEL0005' THEN 1 ELSE 0 END) AS UEEEL0005,
                          MAX(CASE WHEN units = 'UEECD0016' THEN 1 ELSE 0 END) AS UEECD0016,
                          MAX(CASE WHEN units = 'UEEEL0047' THEN 1 ELSE 0 END) AS UEEEL0047,
                          MAX(CASE WHEN units = 'HTLTAID009' THEN 1 ELSE 0 END) AS HTLTAID009,
                          MAX(CASE WHEN units = 'UETDRRF004' THEN 1 ELSE 0 END) AS UETDRRF004,
                          MAX(CASE WHEN units = 'UEEEL0012' THEN 1 ELSE 0 END) AS UEEEL0012,
                          MAX(CASE WHEN units = 'UEEEL0014' THEN 1 ELSE 0 END) AS UEEEL0014,
                          MAX(CASE WHEN units = 'UEEEL0039' THEN 1 ELSE 0 END) AS UEEEL0039
                   FROM ElectrotechnologyReports.dbo.StudentLogs
                   GROUP BY [Student ID]) AS source
            ON target.[Student ID] = source.[Student ID]
            WHEN MATCHED THEN
                UPDATE SET
                    target.UEECO0023 = source.UEECO0023,
                    target.UEECD0007 = source.UEECD0007,
                    target.UEECD0019 = source.UEECD0019,
                    target.UEECD0020 = source.UEECD0020,
                    target.UEECD0051 = source.UEECD0051,
                    target.UEECD0046 = source.UEECD0046,
                    target.UEECD0044 = source.UEECD0044,
                    target.UEEEL0021 = source.UEEEL0021,
                    target.UEEEL0019 = source.UEEEL0019,
                    target.UEERE0001 = source.UEERE0001,
                    target.UEEEL0023 = source.UEEEL0023,
                    target.UEEEL0020 = source.UEEEL0020,
                    target.UEEEL0025 = source.UEEEL0025,
                    target.UEEEL0024 = source.UEEEL0024,
                    target.UEEEL0008 = source.UEEEL0008,
                    target.UEEEL0009 = source.UEEEL0009,
                    target.UEEEL0010 = source.UEEEL0010,
                    target.UEEDV0005 = source.UEEDV0005,
                    target.UEEDV0008 = source.UEEDV0008,
                    target.UEEEL0003 = source.UEEEL0003,
                    target.UEEEL0018 = source.UEEEL0018,
                    target.UEEEL0005 = source.UEEEL0005,
                    target.UEECD0016 = source.UEECD0016,
                    target.UEEEL0047 = source.UEEEL0047,
                    target.HTLTAID009 = source.HTLTAID009,
                    target.UETDRRF004 = source.UETDRRF004,
                    target.UEEEL0012 = source.UEEEL0012,
                    target.UEEEL0014 = source.UEEEL0014,
                    target.UEEEL0039 = source.UEEEL0039
            WHEN NOT MATCHED BY TARGET THEN
                INSERT ([Student ID], UEECO0023, UEECD0007, UEECD0019, UEECD0020, UEECD0051, UEECD0046, UEECD0044, UEEEL0021, UEEEL0019, UEERE0001, UEEEL0023, UEEEL0020, UEEEL0025, UEEEL0024, UEEEL0008, UEEEL0009, UEEEL0010, UEEDV0005, UEEDV0008, UEEEL0003, UEEEL0018, UEEEL0005, UEECD0016, UEEEL0047, HTLTAID009, UETDRRF004, UEEEL0012, UEEEL0014, UEEEL0039)
                VALUES (source.[Student ID], source.UEECO0023, source.UEECD0007, source.UEECD0019, source.UEECD0020, source.UEECD0051, source.UEECD0046, source.UEECD0044, source.UEEEL0021, source.UEEEL0019, source.UEERE0001, source.UEEEL0023, source.UEEEL0020, source.UEEEL0025, source.UEEEL0024, source.UEEEL0008, source.UEEEL0009, source.UEEEL0010, source.UEEDV0005, source.UEEDV0008, source.UEEEL0003, source.UEEEL0018, source.UEEEL0005, source.UEECD0016, source.UEEEL0047, source.HTLTAID009, source.UETDRRF004, source.UEEEL0012, source.UEEEL0014, source.UEEEL0039);"

        currentStep = 90
        ' Update progress bar to reflect current progress
        LoadingForm.UpdateProgress(currentStep)

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Using command As New SqlCommand(sqlQuery, connection)
                    command.ExecuteNonQuery() ' Execute the SQL query
                End Using
            End Using
            LoadingForm.UpdateProgress(totalSteps)
            MessageBox.Show("Student units updated successfully.")
        Catch ex As Exception
            MessageBox.Show("Error updating student units: " & ex.Message)
        End Try

        LoadingForm.Hide()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        ' Prompt user to select CSV file
        Dim openFileDialog As New OpenFileDialog
        openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        openFileDialog.Title = "Select CSV File"
        LoadingForm.Show()

        currentStep = 10
        ' Update progress bar to reflect current progress
        LoadingForm.UpdateProgress(currentStep)

        Try
            ' Truncate existing table
            Dim connectionString = SQLCon.connectionString
            Dim tableName = "ElectrotechnologyReports.dbo.StudentUnitsDatabase"
            Using connection As New SqlConnection(connectionString)
                connection.Open()

                Dim truncateTableQuery = $"TRUNCATE TABLE {tableName}"
                Using truncateTableCommand As New SqlCommand(truncateTableQuery, connection)
                    truncateTableCommand.ExecuteNonQuery()
                End Using
            End Using

            currentStep = 20
            ' Update progress bar to reflect current progress
            LoadingForm.UpdateProgress(currentStep)

            ' Update progress bar to reflect current progress
            LoadingForm.UpdateProgress(10)

            If openFileDialog.ShowDialog = DialogResult.OK Then
                Dim csvFilePath = openFileDialog.FileName

                ' Read all lines from the CSV file
                Dim lines = File.ReadAllLines(csvFilePath)

                currentStep = 30
                ' Update progress bar to reflect current progress
                LoadingForm.UpdateProgress(currentStep)

                ' Identify column indices based on header names
                Dim headers = lines(0).Split(","c)
                Dim studentIDIndex = Array.IndexOf(headers, "Student ID")
                Dim studyPackageCodeIndex = Array.IndexOf(headers, "Study Package Code")
                Dim gradeCodeIndex = Array.IndexOf(headers, "Grade Code")
                Dim studentStudyPackageStatusIndex = Array.IndexOf(headers, "Student Study Package Status")

                currentStep = 40
                ' Update progress bar to reflect current progress
                LoadingForm.UpdateProgress(currentStep)


                ' Connect to SQL Server database
                ' Dim connectionString As String = SQLCon.connectionString
                '  Dim tableName As String = "ElectrotechnologyReports.dbo.StudentUnitsDatabase"
                Using connection As New SqlConnection(connectionString)
                    connection.Open()


                    currentStep = 50
                    ' Update progress bar to reflect current progress
                    LoadingForm.UpdateProgress(currentStep)

                    ' Iterate through each line in the CSV
                    For i = 1 To lines.Length - 1 ' Start from index 1 to skip header row
                        Dim data = lines(i).Split(","c)
                        Dim meetsCriteria = False

                        If data.Length > studentIDIndex AndAlso data.Length > studyPackageCodeIndex Then
                            ' Check if Grade Code is CBC, PP, or GC
                            If data(gradeCodeIndex) = "CBC" OrElse data(gradeCodeIndex) = "PP" OrElse data(gradeCodeIndex) = "GC" Then
                                meetsCriteria = True
                            End If

                            ' Check if Student Study Package Status is Credited, Passed, or Exempt
                            If data(studentStudyPackageStatusIndex) = "Credited" OrElse data(studentStudyPackageStatusIndex) = "Passed" OrElse data(studentStudyPackageStatusIndex) = "Exempt" Then
                                meetsCriteria = True
                            End If

                            ' Check if Student Study Package Status is Enrolled and Grade Code is CBC
                            If data(studentStudyPackageStatusIndex) = "Enrolled" AndAlso data(gradeCodeIndex) = "CBC" Then
                                meetsCriteria = True
                            End If

                            If meetsCriteria Then
                                ' Insert data into the database or perform further processing
                                ' Example: insertCommand.ExecuteNonQuery()
                                Dim studentID = data(studentIDIndex)
                                Dim unit = data(studyPackageCodeIndex)

                                ' Insert the record into the database
                                Dim insertQuery = $"INSERT INTO {tableName} ([Student ID], units) VALUES (@StudentID, @Unit)"
                                Using insertCommand As New SqlCommand(insertQuery, connection)
                                    insertCommand.Parameters.AddWithValue("@StudentID", studentID)
                                    insertCommand.Parameters.AddWithValue("@Unit", unit)
                                    insertCommand.ExecuteNonQuery()
                                End Using
                            End If
                        End If
                    Next

                    ' MessageBox.Show("Data inserted into the table.")
                End Using
            Else
                MessageBox.Show("No file selected.")
            End If
        Catch ex As Exception
            MessageBox.Show($"An error occurred: {ex.Message}")
        Finally
            currentStep = 60
            ' Update progress bar to reflect current progress
            LoadingForm.UpdateProgress(currentStep)
            ' Close the loading form
            LoadingForm.Close()
        End Try
        UploadtoSQLdb()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Open a file dialog to allow the user to select an image file
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Image Files (*.jpg, *.jpeg, *.png, *.gif)|*.jpg;*.jpeg;*.png;*.gif"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Try
                ' Load the selected image file
                Dim originalImage As New Bitmap(openFileDialog.FileName)

                ' Check if the image dimensions exceed the allowed maximum
                If originalImage.Width > MaxWidth OrElse originalImage.Height > MaxHeight Then
                    MessageBox.Show($"The selected image exceeds the maximum allowed dimensions of {MaxWidth}x{MaxHeight} pixels.")
                    Return
                End If

                ' Read the selected image file into a byte array
                Dim imageBytes As Byte() = File.ReadAllBytes(openFileDialog.FileName)

                ' Update the image data in the SQL Server database
                UpdateImageInDatabase(imageBytes)

            Catch ex As Exception
                MessageBox.Show("Error updating image: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub UpdateImageInDatabase(imageBytes As Byte())
        Try
            Dim originalImage As New Bitmap(New MemoryStream(imageBytes))

            If originalImage.Width > MaxWidth OrElse originalImage.Height > MaxHeight Then
                MessageBox.Show($"The selected image exceeds the maximum allowed dimensions of {MaxWidth}x{MaxHeight} pixels.")
                Return
            End If

            ' Update the image data in the SQL Server database
            Dim commandText As String = "IF EXISTS (SELECT 1 FROM ElectrotechnologyReports.dbo.EmailSettings) " &
                                    "UPDATE TOP (1) ElectrotechnologyReports.dbo.EmailSettings SET [Email Signature Image] = @ImageData " &
                                    "ELSE " &
                                    "INSERT INTO ElectrotechnologyReports.dbo.EmailSettings ([Email Signature Image]) VALUES (@ImageData)"

            Using connection As New SqlConnection(SQLCon.connectionString)
                connection.Open()

                Dim command As New SqlCommand(commandText, connection)
                command.Parameters.AddWithValue("@ImageData", imageBytes)

                Dim rowsAffected As Integer = command.ExecuteNonQuery()

                If rowsAffected > 0 Then
                    MessageBox.Show("Image updated successfully.")
                Else
                    MessageBox.Show("No records updated.")
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error updating image: " & ex.Message)
        End Try
    End Sub
    Private Sub UpdateStudentLogs()
        ' SQL query to execute the update and insert operations
        Dim sql As String = "
            -- Update StudentLogs table based on StudentUnitsDatabase
            UPDATE StudentLogs
            SET 
                UEECO0023 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEECO0023'
                ) THEN 1 ELSE 0 END,
                UEECD0007 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEECD0007'
                ) THEN 1 ELSE 0 END,
                UEECD0019 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEECD0019'
                ) THEN 1 ELSE 0 END,
                UEECD0020 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEECD0020'
                ) THEN 1 ELSE 0 END,
                UEECD0051 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEECD0051'
                ) THEN 1 ELSE 0 END,
                UEECD0046 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEECD0046'
                ) THEN 1 ELSE 0 END,
                UEECD0044 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEECD0044'
                ) THEN 1 ELSE 0 END,
                UEEEL0021 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0021'
                ) THEN 1 ELSE 0 END,
                UEEEL0019 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0019'
                ) THEN 1 ELSE 0 END,
                UEERE0001 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEERE0001'
                ) THEN 1 ELSE 0 END,
                UEEEL0023 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0023'
                ) THEN 1 ELSE 0 END,
                UEEEL0020 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0020'
                ) THEN 1 ELSE 0 END,
                UEEEL0025 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0025'
                ) THEN 1 ELSE 0 END,
                UEEEL0024 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0024'
                ) THEN 1 ELSE 0 END,
                UEEEL0008 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0008'
                ) THEN 1 ELSE 0 END,
                UEEEL0009 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0009'
                ) THEN 1 ELSE 0 END,
                UEEEL0010 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0010'
                ) THEN 1 ELSE 0 END,
                UEEDV0005 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEDV0005'
                ) THEN 1 ELSE 0 END,
                UEEDV0008 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEDV0008'
                ) THEN 1 ELSE 0 END,
                UEEEL0003 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0003'
                ) THEN 1 ELSE 0 END,
                UEEEL0018 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0018'
                ) THEN 1 ELSE 0 END,
                UEEEL0005 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0005'
                ) THEN 1 ELSE 0 END,
                UEECD0016 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEECD0016'
                ) THEN 1 ELSE 0 END,
                UEEEL0047 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0047'
                ) THEN 1 ELSE 0 END,
                HTLTAID009 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'HTLTAID009'
                ) THEN 1 ELSE 0 END,
                UETDRRF004 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UETDRRF004'
                ) THEN 1 ELSE 0 END,
                UEEEL0012 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0012'
                ) THEN 1 ELSE 0 END,
                UEEEL0014 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0014'
                ) THEN 1 ELSE 0 END,
                UEEEL0039 = CASE WHEN EXISTS (
                    SELECT 1 FROM StudentUnitsDatabase WHERE StudentUnitsDatabase.[Student ID] = StudentLogs.[Student ID] AND StudentUnitsDatabase.Units = 'UEEEL0039'
                ) THEN 1 ELSE 0 END
                -- Repeat for other units
            WHERE StudentLogs.[Student ID] IN (
                SELECT [Student ID] FROM StudentUnitsDatabase
            );

            -- Insert new rows into StudentLogs table for student IDs not already present
            INSERT INTO StudentLogs ([Student ID], UEECO0023, UEECD0007, UEECD0019, UEECD0020, UEECD0051, UEECD0046, UEECD0044, UEEEL0021, UEEEL0019, UEERE0001, UEEEL0023, UEEEL0020, UEEEL0025, UEEEL0024, UEEEL0008, UEEEL0009, UEEEL0010, UEEDV0005, UEEDV0008, UEEEL0003, UEEEL0018, UEEEL0005, UEECD0016, UEEEL0047, HTLTAID009, UETDRRF004, UEEEL0012, UEEEL0014, UEEEL0039)
            SELECT DISTINCT [Student ID],
                CASE WHEN Units = 'UEECO0023' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEECD0007' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEECD0019' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEECD0020' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEECD0051' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEECD0046' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEECD0044' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0021' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0019' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEERE0001' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0023' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0020' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0025' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0024' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0008' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0009' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0010' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEDV0005' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEDV0008' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0003' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0018' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0005' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEECD0016' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0047' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'HTLTAID009' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UETDRRF004' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0012' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0014' THEN 1 ELSE 0 END,
                CASE WHEN Units = 'UEEEL0039' THEN 1 ELSE 0 END
                -- Repeat for other units
            FROM StudentUnitsDatabase
            WHERE [Student ID] NOT IN (SELECT [Student ID] FROM StudentLogs);"

        ' Create a connection to the database
        Using connection As New SqlConnection(connectionString)
            ' Open the connection
            connection.Open()

            ' Create a command to execute the SQL query
            Using command As New SqlCommand(sql, connection)
                ' Execute the command
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Open a file dialog to select the CSV file
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "CSV files (*.csv)|*.csv"
        openFileDialog.Title = "Select a CSV File"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Read CSV data into the DataSet
            ReadCsvIntoDataSet(openFileDialog.FileName, dataSet)

            ' Filter the DataSet
            FilterDataSet(dataSet)

            ' Display the filtered data
            DataGridView1.DataSource = dataSet.Tables(0)

            ' Show only specific columns in the DataGridView
            Dim visibleColumns As String() = {"Student ID", "Study Package Code", "Grade Code", "Student Study Package Status"}
            For Each column As DataGridViewColumn In DataGridView1.Columns
                column.Visible = visibleColumns.Contains(column.HeaderText)
            Next

            ' Upload filtered data to SQL database
            UploadToDatabase(dataSet)
        End If
        UpdateDatabaseUpdateDate()
        UpdateStudentLogs()

    End Sub
    Private Sub UpdateDatabaseUpdateDate()
        ' Construct the SQL query to update or insert the date
        Dim query As String = "IF EXISTS (SELECT 1 FROM ElectrotechnologyReports.dbo.Updates WHERE ID = 1) " &
                          "BEGIN " &
                          "    UPDATE ElectrotechnologyReports.dbo.Updates SET DatabaseUpdateDate = @Date WHERE ID = 1 " &
                          "END " &
                          "ELSE " &
                          "BEGIN " &
                          "    INSERT INTO ElectrotechnologyReports.dbo.Updates (ID, DatabaseUpdateDate) VALUES (1, @Date) " &
                          "END"

        Try
            ' Create a new SqlConnection object using your connection string
            Using connection As New SqlConnection(SQLCon.connectionString)
                ' Create a new SqlCommand object with the query and connection
                Using command As New SqlCommand(query, connection)
                    ' Add a parameter for the date
                    command.Parameters.AddWithValue("@Date", DateTime.Today)

                    ' Open the connection
                    connection.Open()

                    ' Execute the SQL query
                    command.ExecuteNonQuery()
                End Using
            End Using

            MessageBox.Show("Database update date successfully updated.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error updating database update date: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub ReadCsvIntoDataSet(filePath As String, dataSet As DataSet)
        Using parser As New TextFieldParser(filePath)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")

            ' Read the header line to get column names
            Dim headers As String() = parser.ReadFields()

            ' Create a DataTable with the same structure as the CSV
            Dim dataTable As New DataTable()
            For Each header In headers
                dataTable.Columns.Add(header)
            Next
            dataSet.Tables.Add(dataTable)

            ' Read data into the DataTable
            While Not parser.EndOfData
                dataTable.Rows.Add(parser.ReadFields())
            End While
        End Using
    End Sub

    Sub FilterDataSet(dataSet As DataSet)
        ' Apply filtering conditions using LINQ
        Dim filteredRows = From row In dataSet.Tables(0).AsEnumerable()
                           Where row.Field(Of String)("Grade Code") = "CBC" Or
                             row.Field(Of String)("Grade Code") = "PP" Or
                             row.Field(Of String)("Grade Code") = "GC" Or
                             row.Field(Of String)("Student Study Package Status") = "Credited" Or
                             row.Field(Of String)("Student Study Package Status") = "Passed" Or
                             row.Field(Of String)("Student Study Package Status") = "Exempt" Or
                            (row.Field(Of String)("Student Study Package Status") = "Enrolled" AndAlso
                             row.Field(Of String)("Grade Code") = "CBC")
                           Select row

        ' Create a new DataTable with filtered rows
        Dim filteredDataTable As DataTable = dataSet.Tables(0).Clone()
        For Each filteredRow In filteredRows
            filteredDataTable.ImportRow(filteredRow)
        Next

        ' Replace the original DataTable with the filtered one
        dataSet.Tables.Clear()
        dataSet.Tables.Add(filteredDataTable)
    End Sub

    Sub UploadToDatabase(dataSet As DataSet)
        Dim connectionString As String = SQLCon.connectionString

        ' Establish connection to SQL Server
        Using connection As New SqlConnection(connectionString)
            ' Open the connection
            connection.Open()

            ' Truncate existing table
            Dim truncateCommand As New SqlCommand("TRUNCATE TABLE ElectrotechnologyReports.dbo.StudentUnitsDatabase", connection)
            truncateCommand.ExecuteNonQuery()

            ' Insert filtered data into the database
            Dim insertCommand As New SqlCommand("INSERT INTO ElectrotechnologyReports.dbo.StudentUnitsDatabase ([Student ID], [units]) VALUES (@StudentID, @Units)", connection)
            For Each row As DataRow In dataSet.Tables(0).Rows
                insertCommand.Parameters.Clear()
                insertCommand.Parameters.AddWithValue("@StudentID", row("Student ID"))
                insertCommand.Parameters.AddWithValue("@Units", row("Study Package Code"))
                insertCommand.ExecuteNonQuery()
            Next

            ' Close the connection
            connection.Close()
        End Using
    End Sub

    Private Sub InsertData()
        ' Retrieve data from TextBoxes
        Dim teacherFullName As String = TextBox1.Text
        Dim eNumber As String = TextBox2.Text
        Dim email As String = TextBox3.Text
        Dim contactNumber As String = TextBox4.Text
        Dim department As String = ComboBox1.Text
        Dim highestCertificate As String = ComboBox2.Text
        Dim position As String = ComboBox3.Text

        ' Construct SQL INSERT command
        Dim insertCommand As String = "INSERT INTO ElectrotechnologyReports.dbo.TeacherList (Teacher_Full_Name, E_Number, Email, Contact_Number, Department, Highest_Certificate_Taught, Position) " &
                                  "VALUES (@Teacher_Full_Name, @E_Number, @Email, @Contact_Number, @Department, @Highest_Certificate_Taught, @Position)"

        ' Create SqlConnection and SqlCommand objects
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(insertCommand, connection)
                ' Add parameters for column values
                command.Parameters.AddWithValue("@Teacher_Full_Name", teacherFullName)
                command.Parameters.AddWithValue("@E_Number", eNumber)
                command.Parameters.AddWithValue("@Email", email)
                command.Parameters.AddWithValue("@Contact_Number", contactNumber)
                command.Parameters.AddWithValue("@Department", department)
                command.Parameters.AddWithValue("@Highest_Certificate_Taught", highestCertificate)
                command.Parameters.AddWithValue("@Position", position)

                ' Open connection and execute the command
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End Using
        ' Clear the existing data source of the DataGridView
        DataGridView1.DataSource = Nothing
        '-------
        ' SQL query to select all data from your table
        Dim query As String = "Select * From ElectrotechnologyReports.dbo.TeacherList"

        ' Create a SqlConnection object to connect to the database
        Using connection As New SqlConnection(connectionString)
            ' Create a SqlCommand object with the SQL query and the SqlConnection
            Using command As New SqlCommand(query, connection)
                ' Create a DataTable to store the results of the SQL query
                Dim dataTable As New DataTable()

                ' Open the connection to the database
                connection.Open()

                ' Create a SqlDataAdapter to fill the DataTable with the results of the SQL query
                Using adapter As New SqlDataAdapter(command)
                    ' Fill the DataTable with the results of the SQL query
                    adapter.Fill(dataTable)
                End Using

                ' Bind the DataTable to the DataGridView
                DataGridView1.DataSource = dataTable
            End Using
        End Using
        PopulateTeacherComboBox()
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Dim missingFields As String = ""
        If TextBox1.Text = "" Then
            missingFields &= "- Full Name" & vbCrLf
        End If
        If TextBox2.Text = "" Then
            missingFields &= "- E-Number" & vbCrLf
        End If
        If TextBox3.Text = "" Then
            missingFields &= "- Email" & vbCrLf
        End If

        If ComboBox1.Text = "" Then
            missingFields &= "- Department" & vbCrLf
        End If
        If ComboBox2.Text = "" Then
            missingFields &= "- Highest Certificate Taught" & vbCrLf
        End If
        If ComboBox3.Text = "" Then
            missingFields &= "- Position" & vbCrLf
        End If

        If missingFields <> "" Then
            MessageBox.Show("Please fill the following fields:" & vbCrLf & missingFields)
            Exit Sub
        End If

        InsertData()

    End Sub
    Private Sub PopulateTeacherComboBox()
        ' Clear existing items in the ComboBox
        ComboBox4.Items.Clear()

        ' Construct SQL SELECT command to retrieve teachers' full names and IDs
        Dim selectCommand As String = "SELECT E_Number, Teacher_Full_Name FROM ElectrotechnologyReports.dbo.TeacherList ORDER BY Teacher_Full_Name"

        ' Create SqlConnection and SqlCommand objects
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(selectCommand, connection)
                ' Open connection
                connection.Open()

                ' Execute SQL command and read data
                Using reader As SqlDataReader = command.ExecuteReader()
                    ' Loop through the result set
                    While reader.Read()
                        ' Get teacher full name and ID
                        Dim teacherName As String = reader("Teacher_Full_Name").ToString()
                        Dim teacherId As String = reader("E_Number").ToString()

                        ' Create a KeyValuePair to store ID and name
                        Dim teacherPair As New KeyValuePair(Of String, String)(teacherId, teacherName)

                        ' Add the KeyValuePair to the ComboBox
                        ComboBox4.Items.Add(teacherPair)
                    End While
                End Using
            End Using
        End Using

        ' Set DisplayMember and ValueMember properties
        ComboBox4.DisplayMember = "Value"
        ComboBox4.ValueMember = "Key"
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Check if a teacher is selected in the ComboBox
        If ComboBox4.SelectedIndex <> -1 Then
            ' Retrieve the selected teacher's ID from the ComboBox
            Dim selectedTeacherPair As KeyValuePair(Of String, String) = DirectCast(ComboBox4.SelectedItem, KeyValuePair(Of String, String))
            Dim selectedTeacherId As String = selectedTeacherPair.Key

            ' Construct SQL DELETE command
            Dim deleteCommand As String = "DELETE FROM ElectrotechnologyReports.dbo.TeacherList WHERE E_Number = @TeacherId"

            ' Create SqlConnection and SqlCommand objects
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(deleteCommand, connection)
                    ' Add parameter for teacher ID
                    command.Parameters.AddWithValue("@TeacherId", selectedTeacherId)

                    ' Open connection and execute the command
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End Using

            ' Remove the selected teacher from the ComboBox
            ComboBox4.Items.Remove(ComboBox4.SelectedItem)

            ' Clear the selection in the ComboBox
            ComboBox4.SelectedIndex = -1

            ' Optionally, update any other UI components or data structures as needed
        Else
            MessageBox.Show("Please select a teacher to remove.")
        End If
        ' Clear the existing data source of the DataGridView
        DataGridView1.DataSource = Nothing
        '-------
        ' SQL query to select all data from your table
        Dim query As String = "Select * From ElectrotechnologyReports.dbo.TeacherList"

        ' Create a SqlConnection object to connect to the database
        Using connection As New SqlConnection(connectionString)
            ' Create a SqlCommand object with the SQL query and the SqlConnection
            Using command As New SqlCommand(query, connection)
                ' Create a DataTable to store the results of the SQL query
                Dim dataTable As New DataTable()

                ' Open the connection to the database
                connection.Open()

                ' Create a SqlDataAdapter to fill the DataTable with the results of the SQL query
                Using adapter As New SqlDataAdapter(command)
                    ' Fill the DataTable with the results of the SQL query
                    adapter.Fill(dataTable)
                End Using

                ' Bind the DataTable to the DataGridView
                DataGridView1.DataSource = dataTable
            End Using
        End Using
        PopulateTeacherComboBox()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        EmailTemplates.Show()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If ComboBox4.SelectedItem IsNot Nothing Then
            Button3.Visible = False
        Else
            Button3.Visible = True
        End If
        Button4.Visible = True
        Button7.Visible = True

        ' Retrieve the selected item from ComboBox4
        Dim selectedPair As KeyValuePair(Of String, String) = DirectCast(ComboBox4.SelectedItem, KeyValuePair(Of String, String))
        Dim selectedTeacher As String = selectedPair.Value.Trim()

        ' Debugging statement
        MessageBox.Show("Selected Teacher: " & selectedTeacher)

        ' Construct the SQL query with the selected teacher's full name
        Dim query As String = "SELECT * FROM ElectrotechnologyReports.dbo.TeacherList WHERE Teacher_Full_Name = @Teacher_Full_Name"

        Try
            Using connection As New SqlConnection(SQLCon.connectionString)
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@Teacher_Full_Name", selectedTeacher)

                    connection.Open()

                    Using reader As SqlDataReader = command.ExecuteReader()
                        If reader.Read() Then
                            TextBox1.Text = reader("Teacher_Full_Name").ToString()
                            TextBox2.Text = reader("E_Number").ToString()
                            TextBox3.Text = reader("Email").ToString()
                            TextBox4.Text = reader("Contact_Number").ToString()
                            ComboBox1.SelectedItem = reader("Department").ToString()
                            ComboBox2.SelectedItem = reader("Highest_Certificate_Taught").ToString()
                            ComboBox3.SelectedItem = reader("Position").ToString()

                            ' Populate other controls as needed
                        Else
                            ' Clear text boxes and combo boxes if no data found
                            TextBox1.Clear()
                            TextBox2.Clear()
                            TextBox3.Clear()
                            TextBox4.Clear()
                            ComboBox1.SelectedIndex = -1
                            ComboBox2.SelectedIndex = -1
                            ComboBox3.SelectedIndex = -1

                            MessageBox.Show("No data found for selected teacher.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub



    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Button4.Visible = False
        If ComboBox4.SelectedItem IsNot Nothing Then
            Button7.Visible = True
        Else
            Button7.Visible = False
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Button4.Visible = False
        If ComboBox4.SelectedItem IsNot Nothing Then
            Button7.Visible = True
        Else
            Button7.Visible = False
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Button4.Visible = False
        If ComboBox4.SelectedItem IsNot Nothing Then
            Button7.Visible = True
        Else
            Button7.Visible = False
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Button4.Visible = False
        If ComboBox4.SelectedItem IsNot Nothing Then
            Button7.Visible = True
        Else
            Button7.Visible = False
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Button4.Visible = False
        If ComboBox4.SelectedItem IsNot Nothing Then
            Button7.Visible = True
        Else
            Button7.Visible = False
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Button4.Visible = False
        If ComboBox4.SelectedItem IsNot Nothing Then
            Button7.Visible = True
        Else
            Button7.Visible = False
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Button4.Visible = False
        If ComboBox4.SelectedItem IsNot Nothing Then
            Button7.Visible = True
        Else
            Button7.Visible = False
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim teacherFullName As String = TextBox1.Text
        Dim eNumber As String = TextBox2.Text
        Dim email As String = TextBox3.Text
        Dim contactNumber As String = TextBox4.Text
        Dim department As String = ComboBox1.Text
        Dim highestCertificate As String = ComboBox2.Text
        Dim position As String = ComboBox3.Text

        ' Construct the SQL query for updating the database
        Dim query As String = "UPDATE ElectrotechnologyReports.dbo.TeacherList SET Teacher_Full_Name = @Teacher_Full_Name, Email = @Email, Contact_Number = @Contact_Number, Department = @Department, Highest_Certificate_Taught = @Highest_Certificate_Taught, Position = @Position WHERE E_Number = @E_Number"

        Try
            Using connection As New SqlConnection(SQLCon.connectionString)
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@Teacher_Full_Name", teacherFullName)
                    command.Parameters.AddWithValue("@E_Number", eNumber)
                    command.Parameters.AddWithValue("@Email", email)
                    command.Parameters.AddWithValue("@Contact_Number", contactNumber)
                    command.Parameters.AddWithValue("@Department", department)
                    command.Parameters.AddWithValue("@Highest_Certificate_Taught", highestCertificate)
                    command.Parameters.AddWithValue("@Position", position)

                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End Using

            MessageBox.Show("Details saved successfully.", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        '-------
        ' Clear the existing data source of the DataGridView
        DataGridView1.DataSource = Nothing
        '-------
        ' SQL query to select all data from your table
        Dim Newquery As String = "Select * From ElectrotechnologyReports.dbo.TeacherList"

        ' Create a SqlConnection object to connect to the database
        Using connection As New SqlConnection(connectionString)
            ' Create a SqlCommand object with the SQL query and the SqlConnection
            Using command As New SqlCommand(Newquery, connection)
                ' Create a DataTable to store the results of the SQL query
                Dim dataTable As New DataTable()

                ' Open the connection to the database
                connection.Open()

                ' Create a SqlDataAdapter to fill the DataTable with the results of the SQL query
                Using adapter As New SqlDataAdapter(command)
                    ' Fill the DataTable with the results of the SQL query
                    adapter.Fill(dataTable)
                End Using

                ' Bind the DataTable to the DataGridView
                DataGridView1.DataSource = dataTable
            End Using
        End Using
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        PopulateTeacherComboBox()

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        Button3.Visible = True
        Button7.Visible = True
        Button4.Visible = True
    End Sub
End Class