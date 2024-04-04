Imports Microsoft.Data.SqlClient
Imports Microsoft.Office.Server.Search.Administration
Imports Microsoft.SharePoint.Portal.WebControls
Imports Microsoft.VisualBasic.FileIO

Public Class Admin
    Dim dataSet As New DataSet()
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim ColumnsConfirmed As MsgBoxResult

        ColumnsConfirmed = MsgBox("Have you made sure that the following columns exist?" & vbCrLf & " Student ID " & vbCrLf & " Study Package Code " & vbCrLf & " Grade Code " & vbCrLf & " Student Study Package Status", vbYesNo)

        If ColumnsConfirmed = vbNo Then
            Exit Sub ' If columns are not confirmed, exit the subroutine
        End If

        Dim conversionConfirmed As MsgBoxResult

        conversionConfirmed = MsgBox("Have you converted this to a 'CSV comma delimited' file?", vbYesNo)

        If conversionConfirmed = vbNo Then
            Exit Sub ' If conversion is not confirmed, exit the subroutine
        End If

        ' Open a file dialog to select the CSV file
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "CSV files (*.csv)|*.csv"
        openFileDialog.Title = "Select a CSV File"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Read CSV data into the DataSet
            ReadCsvIntoDataSet(openFileDialog.FileName, DataSet)

            ' Filter the DataSet
            FilterDataSet(DataSet)

            ' Display the filtered data
            DataGridView1.DataSource = DataSet.Tables(0)

            ' Show only specific columns in the DataGridView
            Dim visibleColumns As String() = {"Student ID", "Study Package Code", "Grade Code", "Student Study Package Status"}
            For Each column As DataGridViewColumn In DataGridView1.Columns
                column.Visible = visibleColumns.Contains(column.HeaderText)
            Next

            ' Upload filtered data to SQL database
            UploadToDatabase(DataSet)
        End If
        UpdateDatabaseUpdateDate()
        UpdateStudentLogs()
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

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub Admin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadCurrentSettings()
    End Sub
End Class