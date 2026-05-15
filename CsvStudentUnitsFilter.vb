Imports System.Collections.Generic
Imports System.Data

''' <summary>CSV → StudentUnitsDatabase filtering: trim/BOM, case-insensitive grade/status, dedupe student+unit (Exempt beats Planned+CBC, etc.).</summary>
Friend Module CsvStudentUnitsFilter

    Friend Function CsvCellTrim(row As DataRow, columnName As String) As String
        If row Is Nothing OrElse row.Table Is Nothing OrElse Not row.Table.Columns.Contains(columnName) Then Return String.Empty
        If row.IsNull(columnName) Then Return String.Empty
        Dim t As String = Convert.ToString(row(columnName)).Trim()
        If t.Length > 0 AndAlso t(0) = ChrW(&HFEFF) Then t = t.Substring(1).Trim()
        Return t
    End Function

    ''' <summary>Same rules as legacy LINQ filter, with case-insensitive values and no Field(Of String) / DBNull issues.</summary>
    Friend Function RowQualifiesForStudentUnitsUpload(row As DataRow) As Boolean
        Dim g = CsvCellTrim(row, "Grade Code")
        Dim s = CsvCellTrim(row, "Student Study Package Status")
        If g.Equals("CBC", StringComparison.OrdinalIgnoreCase) OrElse
           g.Equals("PP", StringComparison.OrdinalIgnoreCase) OrElse
           g.Equals("GC", StringComparison.OrdinalIgnoreCase) Then Return True
        If s.Equals("Credited", StringComparison.OrdinalIgnoreCase) OrElse
           s.Equals("Passed", StringComparison.OrdinalIgnoreCase) OrElse
           s.Equals("Exempt", StringComparison.OrdinalIgnoreCase) Then Return True
        If s.Equals("Enrolled", StringComparison.OrdinalIgnoreCase) AndAlso g.Equals("CBC", StringComparison.OrdinalIgnoreCase) Then Return True
        Return False
    End Function

    ''' <summary>When the same student+unit appears more than once (e.g. Planned vs Exempt), keep the strongest row for staging.</summary>
    Friend Function UploadRowPriority(row As DataRow) As Integer
        Dim g = CsvCellTrim(row, "Grade Code")
        Dim s = CsvCellTrim(row, "Student Study Package Status")
        If s.Equals("Exempt", StringComparison.OrdinalIgnoreCase) Then Return 300
        If s.Equals("Passed", StringComparison.OrdinalIgnoreCase) Then Return 250
        If s.Equals("Credited", StringComparison.OrdinalIgnoreCase) Then Return 200
        If s.Equals("Enrolled", StringComparison.OrdinalIgnoreCase) AndAlso g.Equals("CBC", StringComparison.OrdinalIgnoreCase) Then Return 150
        If s.Equals("Planned", StringComparison.OrdinalIgnoreCase) AndAlso
           (g.Equals("CBC", StringComparison.OrdinalIgnoreCase) OrElse g.Equals("PP", StringComparison.OrdinalIgnoreCase) OrElse g.Equals("GC", StringComparison.OrdinalIgnoreCase)) Then Return 45
        If g.Equals("CBC", StringComparison.OrdinalIgnoreCase) OrElse g.Equals("PP", StringComparison.OrdinalIgnoreCase) OrElse g.Equals("GC", StringComparison.OrdinalIgnoreCase) Then Return 100
        Return 0
    End Function

    Friend Sub FilterDataSetForStudentUnitsUpload(dataSet As DataSet)
        If dataSet Is Nothing OrElse dataSet.Tables.Count = 0 Then Return
        Dim source As DataTable = dataSet.Tables(0)

        Dim bestRowByKey As New Dictionary(Of String, DataRow)(StringComparer.OrdinalIgnoreCase)
        Dim bestPriByKey As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

        For Each row As DataRow In source.Rows
            If Not RowQualifiesForStudentUnitsUpload(row) Then Continue For
            Dim sid = SQLCon.NormalizeStudentIdForLogs(row("Student ID"))
            Dim unit = CsvCellTrim(row, "Study Package Code")
            If String.IsNullOrWhiteSpace(sid) OrElse String.IsNullOrWhiteSpace(unit) Then Continue For
            Dim p = UploadRowPriority(row)
            Dim key As String = sid & "|" & unit
            If Not bestPriByKey.ContainsKey(key) OrElse p > bestPriByKey(key) Then
                bestPriByKey(key) = p
                bestRowByKey(key) = row
            End If
        Next

        Dim filteredDataTable As DataTable = source.Clone()
        For Each k As String In bestRowByKey.Keys
            filteredDataTable.ImportRow(bestRowByKey(k))
        Next

        dataSet.Tables.Clear()
        dataSet.Tables.Add(filteredDataTable)
    End Sub

End Module
