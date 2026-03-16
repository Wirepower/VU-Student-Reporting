Imports Microsoft.Office.Interop.Excel
Imports System.Data

Module ExcelExtensions
    <System.Runtime.CompilerServices.Extension()>
    Public Function ToDataTable(worksheet As Worksheet) As System.Data.DataTable
        Dim dt As New System.Data.DataTable()

        For col As Integer = 1 To worksheet.UsedRange.Columns.Count
            dt.Columns.Add(DirectCast(worksheet.Cells(1, col).Value, String))
        Next

        For row As Integer = 2 To worksheet.UsedRange.Rows.Count
            Dim newRow As DataRow = dt.Rows.Add()
            For col As Integer = 1 To worksheet.UsedRange.Columns.Count
                newRow(col - 1) = worksheet.Cells(row, col).Value
            Next
        Next

        Return dt
    End Function
End Module




