Imports System.Data
Imports System.IO
Imports System.Globalization
Imports System.Linq
Imports Microsoft.Data.SqlClient
Imports OfficeOpenXml

''' <summary>Excel .xlsx validation/upload via EPPlus (no Excel COM — avoids missing <c>office</c> assembly on .NET 8).</summary>
Module ExcelValidationModule

    Private Sub EnsureEpPlusLicense()
        ' EPPlus 5+ requires a license context (NonCommercial is appropriate for internal VU tooling).
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
    End Sub

    ''' <summary>EPPlus 7+ uses 0-based worksheet index (first sheet = 0). Excel COM used 1-based Worksheets(1).</summary>
    Private Function GetFirstWorksheet(package As ExcelPackage) As ExcelWorksheet
        Dim wb = package.Workbook
        If wb Is Nothing OrElse wb.Worksheets.Count = 0 Then
            Return Nothing
        End If
        Return wb.Worksheets(0)
    End Function

    ''' <summary>Matches expected SQL column types to values EPPlus/Excel may return (DateTime, Double, Int, etc.).</summary>
    Private Function CellValueMatchesExpectedType(expectedType As Type, cellValue As Object) As Boolean
        If cellValue Is Nothing Then
            Return True
        End If

        If expectedType Is GetType(String) Then
            ' SQL nvarchar; Excel often stores "text" as numeric (Agreement ID, codes) — EPPlus returns Double/Long.
            Return TypeOf cellValue Is String OrElse TypeOf cellValue Is Double OrElse TypeOf cellValue Is Integer OrElse
                TypeOf cellValue Is Long OrElse TypeOf cellValue Is Decimal OrElse TypeOf cellValue Is Single
        End If

        If expectedType Is GetType(Date) OrElse expectedType Is GetType(DateTime) Then
            ' Excel stores dates as serial numbers; EPPlus can surface them as Double.
            If TypeOf cellValue Is DateTime OrElse TypeOf cellValue Is Date Then
                Return True
            End If
            If TypeOf cellValue Is Double Then
                Dim serial = CDbl(cellValue)
                Return serial > 0
            End If
            ' Text dates (e.g. "15/03/2024") after export.
            If TypeOf cellValue Is String Then
                Dim d As DateTime
                Return DateTime.TryParse(DirectCast(cellValue, String), CultureInfo.CurrentCulture, DateTimeStyles.None, d) OrElse
                    DateTime.TryParse(DirectCast(cellValue, String), CultureInfo.InvariantCulture, DateTimeStyles.None, d)
            End If
            Return False
        End If

        If expectedType Is GetType(Double) Then
            If TypeOf cellValue Is Double OrElse TypeOf cellValue Is Integer OrElse TypeOf cellValue Is Long OrElse
                TypeOf cellValue Is Decimal OrElse TypeOf cellValue Is Single Then
                Return True
            End If
            If TypeOf cellValue Is String Then
                Dim d As Double
                Return TryCoerceStringToSqlFloat(DirectCast(cellValue, String), d)
            End If
            Return False
        End If

        Return expectedType.IsInstanceOfType(cellValue)
    End Function

    ''' <summary>
    ''' SQL uses float for Student ID / mobile; Excel often has formatted text ("04xx xxx", "+61 ...").
    ''' Strip to digits, map +61 → 0, then parse as number for bulk copy.
    ''' </summary>
    Private Function TryCoerceStringToSqlFloat(raw As String, ByRef outDouble As Double) As Boolean
        If String.IsNullOrWhiteSpace(raw) Then
            Return False
        End If
        Dim t = raw.Trim()
        If Double.TryParse(t, NumberStyles.Any, CultureInfo.CurrentCulture, outDouble) Then Return True
        If Double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, outDouble) Then Return True

        Dim digitsOnly = New String(t.Where(Function(c) Char.IsDigit(c)).ToArray())
        If digitsOnly.Length = 0 Then
            Return False
        End If
        ' After digit strip, "+61 4xx" → "614..." ; convert AU mobiles to leading 0 for a stable numeric form.
        If digitsOnly.StartsWith("61", StringComparison.Ordinal) AndAlso digitsOnly.Length >= 11 Then
            digitsOnly = "0" & digitsOnly.Substring(2)
        End If

        Dim asLong As Long
        If Long.TryParse(digitsOnly, NumberStyles.Integer, CultureInfo.InvariantCulture, asLong) Then
            outDouble = CDbl(asLong)
            Return True
        End If
        Return Double.TryParse(digitsOnly, NumberStyles.Any, CultureInfo.InvariantCulture, outDouble)
    End Function

    Private Function IsNullOrWhite(value As Object) As Boolean
        If value Is Nothing OrElse value Is DBNull.Value Then
            Return True
        End If
        Return String.IsNullOrWhiteSpace(value.ToString())
    End Function

    Private Function IsEntireRowEmpty(row As DataRow) As Boolean
        For Each obj As Object In row.ItemArray
            If Not IsNullOrWhite(obj) Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Sub EnsureFrankTestRow(table As DataTable)
        If table Is Nothing OrElse table.Rows.Count = 0 Then
            Return
        End If

        For Each r As DataRow In table.Rows
            Dim sid = If(r("Student ID"), Nothing)
            Dim blockGroup = If(r("Block Group Code"), Nothing)
            If (Not IsNullOrWhite(sid) AndAlso Convert.ToDouble(sid, CultureInfo.InvariantCulture) = 3635697.0R) OrElse
               (Not IsNullOrWhite(blockGroup) AndAlso String.Equals(Convert.ToString(blockGroup, CultureInfo.InvariantCulture), "UEE30820_TEST", StringComparison.OrdinalIgnoreCase)) Then
                Return
            End If
        Next

        Dim testRow As DataRow = table.NewRow()
        testRow("Agreement ID") = "1234"
        testRow("Student ID") = 3635697.0R
        testRow("Student Given Name") = "Frank"
        testRow("Student Family Name") = "Offer"
        testRow("Epsilon Start Date") = New DateTime(2021, 12, 7)
        testRow("Epsilon End Date") = New DateTime(2025, 12, 5)
        testRow("Student Personal Email") = "Frank@email.com.au"
        testRow("Student Personal Mobile") = 412345678.0R
        testRow("Employer Surname") = "Smith"
        testRow("Employer Given Name") = "Peter"
        testRow("Employer Contact Phone") = "039123456"
        testRow("Employer Email") = "PeterS@EmployerEmail.com.au"
        testRow("Agreement Category") = "Apprentice"
        testRow("Course") = "UEE30820"
        testRow("Course Title") = "Certificate III in Electrotechnology Electrician"
        testRow("Course Status") = "Admitted"
        testRow("Course Location") = "Sunshine"
        testRow("Block Group Code") = "UEE30820_TEST"
        testRow("Agreement Status") = "Active"
        testRow("Agreement Task") = "Manage Obligations"
        testRow("Employer Name") = "Electrical Business Name Pty Ltd (1234)"
        testRow("Employer ABN") = "12345678899"
        testRow("School Name") = "N/A"
        testRow("School Email") = "N/A"
        testRow("Org Unit") = "Engineering and Electrotechnology"
        testRow("Agreement Name") = "Peter Smith - 1234567-12345345"
        testRow("Agreement Type Code") = "APP-FT"
        testRow("Training Plan Generated") = New DateTime(2024, 4, 8)
        testRow("Signed Training Plan Uploaded") = New DateTime(2024, 4, 16)
        testRow("Units for Employer Sign off") = 1.0R
        testRow("Progress Report Generated") = New DateTime(2024, 3, 6)
        testRow("Signed Progress Report Uploaded") = New DateTime(2024, 4, 10)
        testRow("All Units Resulted?") = "Y"
        testRow("All Units Verified?") = "Y"
        testRow("Number Units Completed") = 29.0R
        testRow("Course Hours Completed") = 1150.0R
        testRow("Any Sanctions") = "N"
        testRow("Completion Training Plan Generated") = New DateTime(2023, 1, 18)
        testRow("Signed Completion Training Plan Uploaded") = New DateTime(2023, 2, 28)
        testRow("Department Email") = "elecandeng@vu.edu.au"
        testRow("Student VU Email") = "Frank@students.vu.edu.au"
        testRow("Actual Start Date") = New DateTime(2019, 7, 19)
        testRow("Actual End Date") = New DateTime(2023, 2, 28)
        testRow("Age of Agreement") = 710.0R
        testRow("Apprenticeship Client ID") = "1234567890"
        table.Rows.Add(testRow)
    End Sub

    Private Function TryNormalizeForSql(expectedType As Type, rawValue As Object, ByRef normalized As Object) As Boolean
        normalized = DBNull.Value
        If rawValue Is Nothing OrElse rawValue Is DBNull.Value Then
            Return True
        End If

        If expectedType Is GetType(String) Then
            normalized = Convert.ToString(rawValue, CultureInfo.InvariantCulture)
            Return True
        End If

        If expectedType Is GetType(Double) Then
            If TypeOf rawValue Is Double OrElse TypeOf rawValue Is Integer OrElse TypeOf rawValue Is Long OrElse
                TypeOf rawValue Is Decimal OrElse TypeOf rawValue Is Single Then
                normalized = Convert.ToDouble(rawValue, CultureInfo.InvariantCulture)
                Return True
            End If
            If TypeOf rawValue Is String Then
                Dim d As Double
                If TryCoerceStringToSqlFloat(DirectCast(rawValue, String), d) Then
                    normalized = d
                    Return True
                End If
            End If
            Return False
        End If

        If expectedType Is GetType(Date) OrElse expectedType Is GetType(DateTime) Then
            If TypeOf rawValue Is DateTime OrElse TypeOf rawValue Is Date Then
                normalized = CType(rawValue, DateTime)
                Return True
            End If
            If TypeOf rawValue Is Double Then
                Dim serial = CDbl(rawValue)
                If serial > 0 Then
                    normalized = DateTime.FromOADate(serial)
                    Return True
                End If
                Return False
            End If
            If TypeOf rawValue Is String Then
                Dim dt As DateTime
                If DateTime.TryParse(DirectCast(rawValue, String), CultureInfo.CurrentCulture, DateTimeStyles.None, dt) OrElse
                   DateTime.TryParse(DirectCast(rawValue, String), CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then
                    normalized = dt
                    Return True
                End If
            End If
            Return False
        End If

        normalized = rawValue
        Return True
    End Function

    Public Function ValidateExcelData(excelFilePath As String, sqlColumnNames As List(Of String)) As Boolean
        Dim ext = Path.GetExtension(excelFilePath).ToLowerInvariant()
        If ext = ".xls" Then
            MessageBox.Show(
                "The old Excel 97–2003 format (.xls) is not supported here. Please open the file in Excel and Save As ""Excel Workbook (.xlsx)"", then try again.",
                "Unsupported format",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)
            Return False
        End If

        Dim sqlColumnDataTypes As New Dictionary(Of String, Type)() From {
            {"Agreement ID", GetType(String)},
            {"Student ID", GetType(Double)},
            {"Student Given Name", GetType(String)},
            {"Student Family Name", GetType(String)},
            {"Epsilon Start Date", GetType(Date)},
            {"Epsilon End Date", GetType(Date)},
            {"Student Personal Email", GetType(String)},
            {"Student Personal Mobile", GetType(Double)},
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

        EnsureEpPlusLicense()
        Dim validatedDataTable As DataTable = Nothing

        Using package As New ExcelPackage(New FileInfo(excelFilePath))
            Dim excelWorksheet = GetFirstWorksheet(package)
            If excelWorksheet Is Nothing Then
                MessageBox.Show("The workbook has no worksheets.", "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If
            Dim dimension = excelWorksheet.Dimension
            If dimension Is Nothing Then
                MessageBox.Show("The first worksheet is empty.", "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            ' Detect the row that actually contains headers in raw exports.
            Dim headerRow As Integer = -1
            Dim headerMap As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Dim scanTo As Integer = Math.Min(dimension.End.Row, dimension.Start.Row + 30)
            For row As Integer = dimension.Start.Row To scanTo
                Dim tempMap As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                For col As Integer = dimension.Start.Column To dimension.End.Column
                    Dim header = If(excelWorksheet.Cells(row, col).Value?.ToString(), "").Trim()
                    If Not String.IsNullOrEmpty(header) AndAlso Not tempMap.ContainsKey(header) Then
                        tempMap(header) = col
                    End If
                Next

                If tempMap.ContainsKey("Student ID") AndAlso tempMap.ContainsKey("Agreement ID") AndAlso tempMap.ContainsKey("Epsilon Start Date") Then
                    ' Prefer the lowest matching header row in scan window.
                    ' This handles files where row 3 and row 4 both look like headers.
                    headerRow = row
                    headerMap = tempMap
                End If
            Next

            If headerRow = -1 Then
                MessageBox.Show("Could not detect the header row in the selected workbook.", "Header Detection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            For Each requiredColumn In sqlColumnNames
                If Not headerMap.ContainsKey(requiredColumn) Then
                    MessageBox.Show($"Required column '{requiredColumn}' is missing.", "Missing Column", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return False
                End If
            Next

            validatedDataTable = New DataTable()
            For Each colName In sqlColumnNames
                validatedDataTable.Columns.Add(colName)
            Next

            Dim seenStudentIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For row As Integer = headerRow + 1 To dimension.End.Row
                Dim dr As DataRow = validatedDataTable.NewRow()
                Dim hasAnyValue As Boolean = False

                For Each colName In sqlColumnNames
                    Dim colIndex As Integer = headerMap(colName)
                    Dim rawValue As Object = excelWorksheet.Cells(row, colIndex).Value
                    If Not IsNullOrWhite(rawValue) Then
                        hasAnyValue = True
                    End If

                    Dim normalized As Object = Nothing
                    Dim expectedType As Type = sqlColumnDataTypes(colName)
                    If Not TryNormalizeForSql(expectedType, rawValue, normalized) Then
                        ' Mobile occasionally contains non-numeric text in raw exports.
                        ' SQL column is float; treat unparseable mobile text as NULL rather than hard-failing import.
                        If String.Equals(colName, "Student Personal Mobile", StringComparison.OrdinalIgnoreCase) Then
                            normalized = DBNull.Value
                            dr(colName) = normalized
                            Continue For
                        End If

                        Dim actualType As String = If(rawValue Is Nothing, "null", rawValue.GetType().Name)
                        MessageBox.Show(
                            $"Data type mismatch for column '{colName}' in row {row}. Expected: {expectedType.Name}; actual: {actualType}.",
                            "Data Type Mismatch",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
                        Return False
                    End If
                    dr(colName) = normalized
                Next

                If Not hasAnyValue Then
                    Continue For
                End If

                Dim sidText As String = ""
                If Not IsNullOrWhite(dr("Student ID")) Then
                    sidText = Convert.ToString(dr("Student ID"), CultureInfo.InvariantCulture)
                End If
                If String.IsNullOrWhiteSpace(sidText) Then
                    Continue For
                End If
                If seenStudentIds.Contains(sidText) Then
                    Continue For
                End If
                seenStudentIds.Add(sidText)
                validatedDataTable.Rows.Add(dr)
            Next
        End Using

        Dim result As DialogResult = MessageBox.Show(
            "Validation complete. Would you like to upload the validated data to SQL?",
            "Upload to SQL",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Return UploadToSQL(validatedDataTable)
        End If

        Return True
    End Function

    Private Function UploadToSQL(dataTable As DataTable) As Boolean
        If dataTable Is Nothing Then
            MessageBox.Show("No validated worksheet data is available for upload.", "Upload Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        Try
            Dim cleanTable As DataTable = dataTable.Clone()
            Dim removedEmptyRows As Integer = 0
            Dim removedMissingStudentIdRows As Integer = 0

            For Each row As DataRow In dataTable.Rows
                If IsEntireRowEmpty(row) Then
                    removedEmptyRows += 1
                    Continue For
                End If

                If Not dataTable.Columns.Contains("Student ID") OrElse IsNullOrWhite(row("Student ID")) Then
                    removedMissingStudentIdRows += 1
                    Continue For
                End If

                cleanTable.ImportRow(row)
            Next

            If cleanTable.Rows.Count = 0 Then
                MessageBox.Show("No valid rows found to upload after removing blank rows / missing Student ID.", "Upload Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            EnsureFrankTestRow(cleanTable)

            Dim sqlConnectionString As String = SQLCon.connectionString

            Using sqlConnection As New SqlConnection(sqlConnectionString)
                sqlConnection.Open()

                Dim truncateCommand As New SqlCommand("TRUNCATE TABLE ElectrotechnologyReports.dbo.AgreementsDetails", sqlConnection)
                truncateCommand.ExecuteNonQuery()

                Using bulkCopy As New SqlBulkCopy(sqlConnection)
                    bulkCopy.DestinationTableName = "ElectrotechnologyReports.dbo.AgreementsDetails"

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

                    bulkCopy.WriteToServer(cleanTable)

                    Dim updateCommand As New SqlCommand("UPDATE ElectrotechnologyReports.dbo.Updates SET StudentDatabaseDate = GETDATE()", sqlConnection)
                    updateCommand.ExecuteNonQuery()
                End Using
            End Using

            If removedEmptyRows > 0 OrElse removedMissingStudentIdRows > 0 Then
                MessageBox.Show(
                    $"Upload completed with row cleanup." & vbCrLf &
                    $"- Removed empty rows: {removedEmptyRows}" & vbCrLf &
                    $"- Removed rows missing Student ID: {removedMissingStudentIdRows}",
                    "Upload Cleanup",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)
            End If

            MessageBox.Show("Data uploaded successfully to SQL table.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return True
        Catch ex As Exception
            MessageBox.Show("Failed to upload data to SQL table. Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Function UploadToSQL(excelFilePath As String) As Boolean
        EnsureEpPlusLicense()

        Try
            Using package As New ExcelPackage(New FileInfo(excelFilePath))
                Dim worksheet = GetFirstWorksheet(package)
                If worksheet Is Nothing Then
                    MessageBox.Show("The workbook has no worksheets.", "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return False
                End If
                Dim dataTable As DataTable = worksheet.ToDataTable()

                Dim sqlConnectionString As String = SQLCon.connectionString

                Using sqlConnection As New SqlConnection(sqlConnectionString)
                    sqlConnection.Open()

                    Dim truncateCommand As New SqlCommand("TRUNCATE TABLE ElectrotechnologyReports.dbo.AgreementsDetails", sqlConnection)
                    truncateCommand.ExecuteNonQuery()

                    Using bulkCopy As New SqlBulkCopy(sqlConnection)
                        bulkCopy.DestinationTableName = "ElectrotechnologyReports.dbo.AgreementsDetails"

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

                        bulkCopy.WriteToServer(dataTable)

                        Dim updateCommand As New SqlCommand("UPDATE ElectrotechnologyReports.dbo.Updates SET StudentDatabaseDate = GETDATE()", sqlConnection)
                        updateCommand.ExecuteNonQuery()
                    End Using
                End Using
            End Using

            MessageBox.Show("Data uploaded successfully to SQL table.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return True
        Catch ex As Exception
            MessageBox.Show("Failed to upload data to SQL table. Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

End Module
