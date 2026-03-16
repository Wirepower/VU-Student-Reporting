Imports Microsoft.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
'Imports System.Data.OleDb
Imports OfficeOpenXml


Module ExcelValidationModule
    Public Function ValidateExcelData(excelFilePath As String, sqlColumnNames As List(Of String)) As Boolean
        ' Define SQL column data types
        Dim sqlColumnDataTypes As New Dictionary(Of String, Type)() From {
            {"Agreement ID", GetType(String)},
            {"Student ID", GetType(String)},
            {"Student Given Name", GetType(String)},
            {"Student Family Name", GetType(String)},
            {"Epsilon Start Date", GetType(Date)},
            {"Epsilon End Date", GetType(Date)},
            {"Student Personal Email", GetType(String)},
            {"Student Personal Mobile", GetType(String)},
            {"Employer Surname", GetType(String)},
            {"Employer Given Name", GetType(String)},
            {"Employer Contact Phone", GetType(String)},
            {"Employer Email", GetType(String)},
            {"Agreement Category", GetType(String)},
            {"Course", GetType(String)},
            {"Course Title", GetType(String)},
            {"Course Status", GetType(String)},
            {"Course Location", GetType(String)},
            {"Block Group Code", GetType(String)},
            {"Agreement Status", GetType(String)},
            {"Agreement Task", GetType(String)},
            {"Employer Name", GetType(String)},
            {"Employer ABN", GetType(String)},
            {"School Name", GetType(String)},
            {"School Email", GetType(String)},
            {"Org Unit", GetType(String)},
            {"Agreement Name", GetType(String)},
            {"Agreement Type Code", GetType(String)},
            {"Training Plan Generated", GetType(String)},
            {"Signed Training Plan Uploaded", GetType(String)},
            {"Units for Employer Sign off", GetType(String)},
            {"Progress Report Generated", GetType(String)},
            {"Signed Progress Report Uploaded", GetType(String)},
            {"All Units Resulted?", GetType(String)},
            {"All Units Verified?", GetType(String)},
            {"Number Units Completed", GetType(Double)},
            {"Course Hours Completed", GetType(Double)},
            {"Any Sanctions", GetType(String)},
            {"Completion Training Plan Generated", GetType(String)},
            {"Signed Completion Training Plan Uploaded", GetType(String)},
            {"Department Email", GetType(String)},
            {"Student VU Email", GetType(String)},
            {"Actual Start Date", GetType(Date)},
            {"Actual End Date", GetType(Date)},
            {"Age of Agreement", GetType(Double)},
            {"Apprenticeship Client ID", GetType(String)}
        }

        ' Open Excel file and validate data
        Dim excelApp As New Application()
        Dim excelWorkbook As Workbook = excelApp.Workbooks.Open(excelFilePath)
        Dim excelWorksheet As Worksheet = excelWorkbook.Sheets(1)

        Dim isValid As Boolean = True

        ' Validate data types of each column and perform additional checks
        For col As Integer = 1 To excelWorksheet.UsedRange.Columns.Count
            Dim columnName As String = excelWorksheet.Cells(1, col).Value

            ' Ensure columnName is not null or empty
            If String.IsNullOrEmpty(columnName) Then
                MessageBox.Show($"Column name in column {col} is null or empty.", "Column Name Error")
                isValid = False
                Exit For
            End If

            ' Ensure columnName exists in sqlColumnDataTypes
            If Not sqlColumnDataTypes.ContainsKey(columnName) Then
                MessageBox.Show($"Column '{columnName}' is not recognized.", "Unknown Column")
                isValid = False
                Exit For
            End If

            Dim expectedType As Type = sqlColumnDataTypes(columnName)

            ' Track Student IDs for duplicate check
            Dim studentIDs As New HashSet(Of String)()

            For row As Integer = 2 To excelWorksheet.UsedRange.Rows.Count
                Dim cellValue As Object = excelWorksheet.Cells(row, col).Value

                ' Check if cell value is null or DBNull
                If cellValue Is Nothing OrElse TypeOf cellValue Is DBNull Then
                    Continue For ' Skip null values
                End If

                ' Check data type of cell value
                If Not expectedType.Equals(cellValue.GetType()) Then
                    ' Data type mismatch
                    MessageBox.Show($"Data type mismatch for column '{columnName}' in row {row}. Expected data type: {expectedType.Name}.", "Data Type Mismatch")
                    isValid = False
                    Exit For
                End If

                ' Check for duplicate Student IDs
                ' If columnName = "Student ID" Then
                'Dim studentID As String = DirectCast(cellValue, String)
                'If studentIDs.Contains(studentID) Then
                ' Duplicate Student ID found
                'MessageBox.Show($"Duplicate Student ID '{studentID}' found in row {row}.", "Duplicate Student ID")
                'isValid = False
                'Exit For
                'Else
                'studentIDs.Add(studentID)
                'End If
                'End If

                ' Additional checks and corrections for Student Personal Mobile column
                If columnName = "Student Personal Mobile" AndAlso TypeOf cellValue Is String Then
                    ' Perform additional checks and corrections
                    Dim mobileNumber As String = DirectCast(cellValue, String)

                    ' Remove non-numeric characters
                    mobileNumber = New String(mobileNumber.Where(Function(c) Char.IsDigit(c)).ToArray())

                    ' Convert prefixes (e.g., replace "+61" with "0")
                    If mobileNumber.StartsWith("+61") Then
                        mobileNumber = "0" & mobileNumber.Substring(3)
                    End If

                    ' Update cell value if necessary
                    If mobileNumber <> DirectCast(cellValue, String) Then
                        ' Automatically correct the cell value
                        excelWorksheet.Cells(row, col).Value = mobileNumber
                    End If
                End If

            Next

            If Not isValid Then
                Exit For
            End If
        Next

        ' Close Excel objects
        excelWorkbook.Close()
        excelApp.Quit()

        ' If validation is successful, prompt user to upload data to SQL
        If isValid Then
            Dim result As DialogResult = MessageBox.Show("Validation complete. Would you like to upload the validated data to SQL?", "Upload to SQL", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                isValid = UploadToSQL(excelFilePath)
            End If
        End If

        Return isValid
    End Function
    Private Function UploadToSQL(excelFilePath As String) As Boolean
        Dim excelApp As New Application()
        Dim workbook As Workbook = excelApp.Workbooks.Open(excelFilePath)
        Dim worksheet As Worksheet = workbook.Sheets(1) ' Assuming data is in the first worksheet

        Try
            ' Get the number of columns in the Excel sheet
            Dim numColumns As Integer = worksheet.UsedRange.Columns.Count

            ' Insert row with the provided data
            Dim newRow As Integer = worksheet.Cells(worksheet.Rows.Count, 1).End(XlDirection.xlUp).Row + 1

            ' Define the data array with the correct dimensions
            Dim data(0, numColumns - 1) As Object
            ' Add data to the cells in the new row
            data(0, 0) = "1234"
            data(0, 1) = "3635697"
            data(0, 2) = "Frank"
            data(0, 3) = "Offer"
            data(0, 4) = #12/7/2021#
            data(0, 5) = #12/5/2025#
            data(0, 6) = "Frank@email.com.au"
            data(0, 7) = "412345678"
            data(0, 8) = "Smith"
            data(0, 9) = "Peter"
            data(0, 10) = "039123456"
            data(0, 11) = "PeterS@EmployerEmail.com.au"
            data(0, 12) = "Apprentice"
            data(0, 13) = "UEE30820"
            data(0, 14) = "Certificate III in Electrotechnology Electrician"
            data(0, 15) = "Admitted"
            data(0, 16) = "Sunshine"
            data(0, 17) = "UEE30820_TEST"
            data(0, 18) = "Active"
            data(0, 19) = "Manage Obligations"
            data(0, 20) = "Electrical Business Name Pty Ltd (1234)"
            data(0, 21) = "12345678899"
            data(0, 22) = "N/A"
            data(0, 23) = "N/A"
            data(0, 24) = "Engineering and Electrotechnology"
            data(0, 25) = "Peter Smith - 1234567-12345345"
            data(0, 26) = "APP-FT"
            data(0, 27) = #4/8/2024#
            data(0, 28) = #4/16/2024#
            data(0, 29) = 1
            data(0, 30) = #3/6/2024#
            data(0, 31) = #4/10/2024#
            data(0, 32) = "Y"
            data(0, 33) = "Y"
            data(0, 34) = 29
            data(0, 35) = 1150
            data(0, 36) = "N"
            data(0, 37) = #1/18/2023#
            data(0, 38) = #2/28/2023#
            data(0, 39) = "elecandeng@vu.edu.au"
            data(0, 40) = "Frank@students.vu.edu.au"
            data(0, 41) = #7/19/2019#
            data(0, 42) = #2/28/2023#
            data(0, 43) = 710
            data(0, 44) = "1234567890"

            ' Set the range where data will be inserted
            Dim range As Range = worksheet.Range(worksheet.Cells(newRow, 1), worksheet.Cells(newRow, numColumns))
            range.Value = data

            ' Connection string to your SQL Server
            Dim sqlConnectionString As String = SQLCon.connectionString

            ' Create SQL connection
            Using sqlConnection As New SqlConnection(sqlConnectionString)
                ' Open the connection
                sqlConnection.Open()

                ' Truncate the SQL table
                Dim truncateCommand As New SqlCommand("TRUNCATE TABLE ElectrotechnologyReports.dbo.AgreementsDetails", sqlConnection)
                truncateCommand.ExecuteNonQuery()

                ' SQL bulk copy operation
                Using bulkCopy As New SqlBulkCopy(sqlConnection)
                    ' Set the destination table name
                    bulkCopy.DestinationTableName = "ElectrotechnologyReports.dbo.AgreementsDetails"

                    ' Map Excel columns to SQL table columns
                    bulkCopy.ColumnMappings.Add("Agreement ID", "Agreement ID")
                    bulkCopy.ColumnMappings.Add("Student ID", "Student ID")
                    bulkCopy.ColumnMappings.Add("Student Given Name", "Student Given Name")
                    bulkCopy.ColumnMappings.Add("Student Family Name", "Student Family Name")
                    bulkCopy.ColumnMappings.Add("Epsilon Start Date", "Epsilon Start Date")
                    bulkCopy.ColumnMappings.Add("Epsilon End Date", "Epsilon End Date")
                    bulkCopy.ColumnMappings.Add("Student Personal Email", "Student Personal Email")
                    bulkCopy.ColumnMappings.Add("Student Personal Mobile", "Student Personal Mobile")
                    bulkCopy.ColumnMappings.Add("Employer Surname", "Employer Surname")
                    bulkCopy.ColumnMappings.Add("Employer Given Name", "Employer Given Name")
                    bulkCopy.ColumnMappings.Add("Employer Contact Phone", "Employer Contact Phone")
                    bulkCopy.ColumnMappings.Add("Employer Email", "Employer Email")
                    bulkCopy.ColumnMappings.Add("Agreement Category", "Agreement Category")
                    bulkCopy.ColumnMappings.Add("Course", "Course")
                    bulkCopy.ColumnMappings.Add("Course Title", "Course Title")
                    bulkCopy.ColumnMappings.Add("Course Status", "Course Status")
                    bulkCopy.ColumnMappings.Add("Course Location", "Course Location")
                    bulkCopy.ColumnMappings.Add("Block Group Code", "Block Group Code")
                    bulkCopy.ColumnMappings.Add("Agreement Status", "Agreement Status")
                    bulkCopy.ColumnMappings.Add("Agreement Task", "Agreement Task")
                    bulkCopy.ColumnMappings.Add("Employer Name", "Employer Name")
                    bulkCopy.ColumnMappings.Add("Employer ABN", "Employer ABN")
                    bulkCopy.ColumnMappings.Add("School Name", "School Name")
                    bulkCopy.ColumnMappings.Add("School Email", "School Email")
                    bulkCopy.ColumnMappings.Add("Org Unit", "Org Unit")
                    bulkCopy.ColumnMappings.Add("Agreement Name", "Agreement Name")
                    bulkCopy.ColumnMappings.Add("Agreement Type Code", "Agreement Type Code")
                    bulkCopy.ColumnMappings.Add("Training Plan Generated", "Training Plan Generated")
                    bulkCopy.ColumnMappings.Add("Signed Training Plan Uploaded", "Signed Training Plan Uploaded")
                    bulkCopy.ColumnMappings.Add("Units for Employer Sign off", "Units for Employer Sign off")
                    bulkCopy.ColumnMappings.Add("Progress Report Generated", "Progress Report Generated")
                    bulkCopy.ColumnMappings.Add("Signed Progress Report Uploaded", "Signed Progress Report Uploaded")
                    bulkCopy.ColumnMappings.Add("All Units Resulted?", "All Units Resulted?")
                    bulkCopy.ColumnMappings.Add("All Units Verified?", "All Units Verified?")
                    bulkCopy.ColumnMappings.Add("Number Units Completed", "Number Units Completed")
                    bulkCopy.ColumnMappings.Add("Course Hours Completed", "Course Hours Completed")
                    bulkCopy.ColumnMappings.Add("Any Sanctions", "Any Sanctions")
                    bulkCopy.ColumnMappings.Add("Completion Training Plan Generated", "Completion Training Plan Generated")
                    bulkCopy.ColumnMappings.Add("Signed Completion Training Plan Uploaded", "Signed Completion Training Plan Uploaded")
                    bulkCopy.ColumnMappings.Add("Department Email", "Department Email")
                    bulkCopy.ColumnMappings.Add("Student VU Email", "Student VU Email")
                    bulkCopy.ColumnMappings.Add("Actual Start Date", "Actual Start Date")
                    bulkCopy.ColumnMappings.Add("Actual End Date", "Actual End Date")
                    bulkCopy.ColumnMappings.Add("Age of Agreement", "Age of Agreement")
                    bulkCopy.ColumnMappings.Add("Apprenticeship Client ID", "Apprenticeship Client ID")

                    ' Load data from Excel to SQL table
                    bulkCopy.WriteToServer(worksheet.ToDataTable())
                    ' Update StudentDatabaseDate in ElectrotechnologyReports.dbo.Updates table with today's date
                    Dim updateCommand As New SqlCommand("UPDATE ElectrotechnologyReports.dbo.Updates SET StudentDatabaseDate = GETDATE()", sqlConnection)
                    updateCommand.ExecuteNonQuery()
                End Using
            End Using

            ' Display success message
            MessageBox.Show("Data uploaded successfully to SQL table.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)



            Return True ' Upload successful
        Catch ex As Exception
            ' Handle any errors that occur during the upload process
            MessageBox.Show("Failed to upload data to SQL table. Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False ' Upload failed
        Finally
            ' Close Excel objects
            workbook.Close()
            excelApp.Quit()
        End Try
    End Function


End Module
