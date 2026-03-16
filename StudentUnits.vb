Imports System.ComponentModel
Imports Microsoft.Data.SqlClient
'Imports System.Data.SqlClient
Imports Student_Attendance_Reporting
Imports System.Windows.Forms
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports iText.Forms
Imports System.Reflection.PortableExecutable
Imports System.IO
Imports System.Globalization

Public Class StudentUnits
    Private connection As SqlConnection
    Friend Shared studentUnitsForm As New StudentUnits()


    Private Sub CloseBTN_Click(sender As Object, e As EventArgs) Handles CloseBTN.Click
        MainFrm.UnitAlertLbl.Text = UnitAlertLbl1.Text
        Me.Close()
        MainFrm.Refresh()
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub LoadDataIntoCheckedListBox()
        ' Get the student ID from the MainForm
        Dim studentID As String = MainFrm.StudentIDLBL.Text

        ' Check if the student ID exists in the StudentLogs table
        Dim query As String = "SELECT COUNT(*) FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"

        Dim count As Integer = 0

        Using command As New SqlCommand(query, connection)
            command.Parameters.AddWithValue("@StudentID", studentID)
            count = Convert.ToInt32(command.ExecuteScalar())
        End Using

        Dim row As Integer = -1

        ' If the student ID doesn't exist, add it to the StudentLogs table
        If count = 0 Then
            Dim insertQuery As String = "INSERT INTO ElectrotechnologyReports.dbo.StudentLogs ([Student ID]) VALUES (@StudentID); SELECT SCOPE_IDENTITY();"

            Try
                Using connection As New SqlConnection(connectionString)
                    connection.Open()
                    Using insertCommand As New SqlCommand(insertQuery, connection)
                        insertCommand.Parameters.AddWithValue("@StudentID", studentID)
                        row = Convert.ToInt32(insertCommand.ExecuteScalar())
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Error inserting student ID: " & ex.Message)
            End Try
        Else
            ' If the student ID already exists, do whatever you need to do
            ' For example, you can display a message or perform another action
            MessageBox.Show("Student ID already exists.")
        End If
    End Sub

    Private Sub SelectedStudentLBL_Click(sender As Object, e As EventArgs) Handles SelectedStudentLBL.Click
        SelectedStudentLBL.Text = MainFrm.SelectedStudentLBL.Text
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEECO0023" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox1.Checked
        UpdateDatabase(studentID, columnName, newValue)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEECD0007" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox2.Checked
        UpdateDatabase(studentID, columnName, newValue)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEECD0019" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox3.Checked
        UpdateDatabase(studentID, columnName, newValue)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub
    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEECD0020" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox4.Checked
        UpdateDatabase(studentID, columnName, newValue)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEECD0051" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox5.Checked
        UpdateDatabase(studentID, columnName, newValue)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEECD0046" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox6.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub
    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEECD0044" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox7.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0021" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox8.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0019" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox9.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub
    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEERE0001" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox10.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0023" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox11.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0020" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox12.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub
    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0025" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox13.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox14_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox14.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0024" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox14.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0008" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox15.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
    End Sub
    Private Sub CheckBox16_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox16.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0009" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox16.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox17_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox17.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0010" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox17.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox18_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox18.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEDV0005" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox18.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub
    Private Sub CheckBox19_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox19.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEDV0008" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox19.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox20_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox20.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0003" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox20.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub
    Private Sub CheckBox21_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox21.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0018" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox21.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox22_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox22.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0005" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox22.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub
    Private Sub CheckBox23_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox23.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEECD0016" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox23.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox24_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox24.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0047" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox24.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox25_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox25.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "HLTAID009" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox25.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub
    Private Sub CheckBox26_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox26.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UETDRRF004" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox26.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
    End Sub

    Private Sub CheckBox27_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox27.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0014" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox27.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
        UpdateButtonVisibility()
    End Sub

    Private Sub CheckBox28_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox28.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0012" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox28.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
        UpdateButtonVisibility()
    End Sub

    Private Sub CheckBox29_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox29.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0039" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox29.Checked
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
        UpdateButtonVisibility()
    End Sub
    Private Sub PopulateTeacherCombo()
        ' Replace "Your_Connection_String_Here" with your actual connection string
        Dim connectionString As String = SQLCon.connectionString

        ' Create and open a SqlConnection
        Using connection As New SqlConnection(connectionString)
            connection.Open()

            ' Clear existing items in the ComboBox
            ComboBox1.Items.Clear()

            ' SQL query to retrieve all teachers' names
            Dim query As String = "SELECT Teacher_Full_Name FROM ElectrotechnologyReports.dbo.TeacherList WHERE Highest_Certificate_Taught = 'Certificate III' ORDER BY Teacher_Full_Name ASC"

            ' Create a SqlCommand object with the query and connection
            Using command As New SqlCommand(query, connection)
                ' Execute the query and retrieve the data
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        ' Add each teacher's name to the ComboBox
                        ComboBox1.Items.Add(reader.GetString(0))
                    End While
                End Using
            End Using
        End Using
    End Sub
    Private Sub StudentUnits_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        VersionLBL.Text = MainFrm.VersionLBL.Text
        SelectedStudentLBL.Text = MainFrm.SelectedStudentLBL.Text
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
        UpdateButtonVisibility()
        PopulateTeacherCombo()
        UpdateLabelWithDatabaseDate()
        '-------------------------------------------------------------
        'Enable below once Address column is working in SQL database
        TextBox1.Text = MainFrm.Label28.Text
        '-------------------------------------------------------------
        If CheckBox27.Checked And CheckBox29.Checked Then
            CheckBox41.Visible = True

        Else
            CheckBox41.Visible = False

        End If

    End Sub
    Private Sub UpdateLabelWithDatabaseDate()
        Try
            ' Construct the SQL query to retrieve the database update date
            Dim query As String = "SELECT DatabaseUpdateDate FROM ElectrotechnologyReports.dbo.Updates WHERE ID = 1"

            ' Create a new SqlConnection object using your connection string
            Using connection As New SqlConnection(SQLCon.connectionString)
                ' Create a new SqlCommand object with the query and connection
                Using command As New SqlCommand(query, connection)
                    ' Open the connection
                    connection.Open()

                    ' Execute the SQL query and get the result
                    Dim result As Object = command.ExecuteScalar()

                    ' Check if the result is not null
                    If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                        ' Convert the result to DateTime
                        Dim databaseUpdateDate As DateTime = Convert.ToDateTime(result)


                        ' Set the label's text property with the database update date formatted as "dd/MM/yyyy"
                        DateLBL.Text = "Database Current as of: " & databaseUpdateDate.ToString("dd/MM/yyyy")

                    Else
                        ' If the result is null or DBNull, display a message indicating no date is available
                        DateLBL.Text = "Database Update Date Not Available"
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' Handle any errors
            MessageBox.Show("Error retrieving database update date: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub UnitAlertLbl1_Click(sender As Object, e As EventArgs) Handles UnitAlertLbl1.Click

    End Sub

    Private Sub UpdateButtonVisibility()
        ' Check if all checkboxes are checked
        Dim allChecked As Boolean = True
        For Each cb As CheckBox In {CheckBox1, CheckBox2, CheckBox3, CheckBox4, CheckBox5, CheckBox6, CheckBox7, CheckBox8, CheckBox9, CheckBox10, CheckBox11, CheckBox12, CheckBox13, CheckBox14, CheckBox15, CheckBox16, CheckBox17, CheckBox18, CheckBox19, CheckBox20, CheckBox21, CheckBox22, CheckBox23, CheckBox24, CheckBox25, CheckBox26, CheckBox28, CheckBox27}
            If Not cb.Checked Then
                allChecked = False
                Exit For
            End If
        Next
        If allChecked Then
            TextBox1.Visible = True
            Label2.Visible = True
            Label5.Visible = True
            CheckBox40.Visible = True
            Label6.Visible = True
            ComboBox1.Visible = True
        Else
            TextBox1.Visible = False
            Label2.Visible = False
            Label5.Visible = False
            CheckBox40.Visible = False
            Label6.Visible = False
            ComboBox1.Visible = False
            'CheckBox41.Visible = False
            ' Set button visibility based on checkbox state
        End If

        If CheckBox27.Checked And CheckBox29.Checked Then
            CheckBox41.Visible = True
        Else
            CheckBox41.Visible = False
        End If

        If Not String.IsNullOrEmpty(TextBox1.Text) AndAlso (CheckBox40.Checked OrElse CheckBox41.Checked) Then
            Button1.Visible = True

        Else
            Button1.Visible = False

        End If
    End Sub
    Public Sub PopulatePdfWithParameters()
        Dim templatePath As String = "LEATemplate.pdf"
        Dim outputDirectory As String = "P:\VUPoly\MT&T\IT, Electrical And Engineering\Submitted LEA Authorisation Forms"
        Dim Todaydate As String = DateTime.Today.ToString("ddMMyyyy")
        Dim fileName As String = MainFrm.StudentIDLBL.Text & "_" & MainFrm.StudentFirstnameLBL.Text & "_" & MainFrm.StudentSurnameLBL.Text & "_" & Todaydate & ".pdf"
        Dim outputPath As String = System.IO.Path.Combine(outputDirectory, fileName)

        Try
            If Not Directory.Exists(outputDirectory) Then
                Directory.CreateDirectory(outputDirectory)
            End If

            Using reader As New PdfReader(templatePath)
                Using stamper As New PdfStamper(reader, New FileStream(outputPath, FileMode.Create))
                    Dim form As AcroFields = stamper.AcroFields

                    Dim fullName As String = MainFrm.StudentFirstnameLBL.Text & " " & MainFrm.StudentSurnameLBL.Text
                    form.SetField("FullName", fullName)
                    form.SetField("Address", TextBox1.Text)
                    form.SetField("Email", MainFrm.StudentEmailLBL.Text)
                    form.SetField("Telephone", MainFrm.Label29.Text)
                    form.SetField("Epsilon", MainFrm.Label34.Text)
                    form.SetField("Address", TextBox1.Text)

                    'Section 1
                    If CheckBox40.Checked Then
                        form.SetField("TeacherName", ComboBox1.Text)
                        form.SetField("Date", Today.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture))
                        form.SetField("TeacherName2", "")
                        form.SetField("Date2", "")
                    End If

                    'Section 2
                    If CheckBox41.Checked Then
                        form.SetField("TeacherName2", ComboBox1.Text)
                        form.SetField("Date2", Today.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture))
                    End If

                    stamper.FormFlattening = True
                End Using
            End Using
            PdfHelper.OpenPdfWithDefaultViewer(outputPath)
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        ' Specify the path to the folder you want to open
        Dim folderPath As String = "P:\VUPoly\MT&T\IT, Electrical and Engineering\Submitted LEA Authorisation Forms\ARCHIVE\"

        Try
            ' Open the folder using the default file explorer with the specified folder path

            Process.Start("explorer.exe", $"/select,""{folderPath}""")
        Catch ex As Exception
            ' Handle any errors that may occur
            MessageBox.Show("Error opening folder: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If Not String.IsNullOrEmpty(TextBox1.Text) AndAlso (CheckBox40.Checked OrElse CheckBox41.Checked) Then
            Button1.Visible = True
        Else
            Button1.Visible = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim response As String

        ' Prompt the user
        response = MessageBox.Show("Is the student's profiling up to date?", "Profile Check", MessageBoxButtons.YesNo)

        ' Check the response
        If response = DialogResult.Yes Then
            ' Call the function to populate PDF with parameters
            PopulatePdfWithParameters()
            'MessageBox.Show("Student profile is up to date. Proceeding with code...", "Profile Status")
        ElseIf response = DialogResult.No Then
            ' Prompt the user
            MessageBox.Show("Student needs to be up to date with profiling before an LEA Authority Form can be generated", "Profile Status")
            ' Exit the subroutine
            Exit Sub
        End If



        'PopulatePdfWithParameters()
    End Sub

    Private Sub CheckBox40_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox40.CheckedChanged
        UpdateButtonVisibility()
    End Sub

    Private Sub CheckBox41_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox41.CheckedChanged


        UpdateButtonVisibility()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
End Class