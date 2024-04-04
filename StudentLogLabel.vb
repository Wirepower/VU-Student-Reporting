Imports Excel = Microsoft.Office.Interop.Excel

Module ExcelLookupModule

    Public Sub PerformLookup(StudentIDLBL As Label, Label19 As Label, Label22 As Label, Label23 As Label)
        If String.IsNullOrEmpty(StudentIDLBL.Text) Then
            ' If StudentIDLBL.Text is empty, exit
            Exit Sub
        End If

        Try
            Dim xlApp As New Excel.Application()
            xlApp.DisplayAlerts = False

            ' Open the workbook
            Dim xlWorkbook As Excel.Workbook = xlApp.Workbooks.Open("P:\VUPoly\MT&T\IT, Electrical and Engineering\Student Reporting Database\StudentLogs.xlsx")
            Dim StudentLogWorksheet As Excel.Worksheet = xlWorkbook.ActiveSheet
            Dim lastRow As Integer = StudentLogWorksheet.Cells(StudentLogWorksheet.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row

            Dim lookupValue As String = StudentIDLBL.Text
            Dim foundRange As Excel.Range = StudentLogWorksheet.Columns("A").Find(lookupValue, LookIn:=Excel.XlFindLookIn.xlValues)

            If Not foundRange Is Nothing Then
                ' Get the row number where the match is found
                Dim rowNumber As Integer = foundRange.Row

                ' Retrieve values from columns E, F, and G of the same row
                Dim valueE As String = StudentLogWorksheet.Cells(rowNumber, "E").Value
                Dim valueF As String = StudentLogWorksheet.Cells(rowNumber, "F").Value
                Dim valueG As String = StudentLogWorksheet.Cells(rowNumber, "G").Value

                ' Set the values to corresponding labels on the form
                Label19.Text = If(String.IsNullOrEmpty(valueE), "0", valueE.ToString())
                Label22.Text = If(String.IsNullOrEmpty(valueF), "0", valueF.ToString())
                Label23.Text = If(String.IsNullOrEmpty(valueG), "0", valueG.ToString())
            Else
                MessageBox.Show("Student ID not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            ' Close the workbook and quit Excel
            xlWorkbook.Close(SaveChanges:=False)
            xlApp.Quit()

            ' Release COM objects
            System.Runtime.InteropServices.Marshal.ReleaseComObject(foundRange)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(StudentLogWorksheet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
        Catch ex As Exception
            ' Handle the error (e.g., display an error message)
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


End Module