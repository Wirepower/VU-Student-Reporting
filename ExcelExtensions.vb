Imports System.Data
Imports System.Runtime.CompilerServices
Imports OfficeOpenXml

''' <summary>EPPlus helpers — no Excel COM / no <c>office</c> interop assembly required.</summary>
Module ExcelExtensions
    <Extension()>
    Public Function ToDataTable(worksheet As ExcelWorksheet) As DataTable
        Dim dt As New DataTable()
        Dim dimension = worksheet.Dimension
        If dimension Is Nothing Then
            Return dt
        End If

        For col As Integer = dimension.Start.Column To dimension.End.Column
            Dim headerObj = worksheet.Cells(1, col).Value
            Dim headerText = If(headerObj?.ToString(), "").Trim()
            If String.IsNullOrEmpty(headerText) Then
                headerText = "Column" & col
            End If
            dt.Columns.Add(headerText)
        Next

        For row As Integer = 2 To dimension.End.Row
            Dim newRow As DataRow = dt.Rows.Add()
            Dim colIndex As Integer = 0
            For col As Integer = dimension.Start.Column To dimension.End.Column
                newRow(colIndex) = worksheet.Cells(row, col).Value
                colIndex += 1
            Next
        Next

        Return dt
    End Function
End Module
