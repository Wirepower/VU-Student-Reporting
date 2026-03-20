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
    Private ReadOnly unitCheckBoxes As New Dictionary(Of String, CheckBox)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly originalCheckBoxText As New Dictionary(Of CheckBox, String)
    Private isRevertingUnitOverride As Boolean = False
    


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

    Private Function PromptOverridePassword() As String
        Using prompt As New Form()
            prompt.Text = "Admin Override Required"
            prompt.StartPosition = FormStartPosition.CenterParent
            prompt.FormBorderStyle = FormBorderStyle.FixedDialog
            prompt.MaximizeBox = False
            prompt.MinimizeBox = False
            prompt.ClientSize = New Size(420, 140)

            Dim lbl As New Label() With {
                .AutoSize = False,
                .Location = New Point(12, 12),
                .Size = New Size(396, 35),
                .Text = "This requires ADMIN authority to override. Input Password"
            }

            Dim txt As New TextBox() With {
                .Location = New Point(12, 55),
                .Size = New Size(396, 24),
                .UseSystemPasswordChar = True
            }

            Dim okBtn As New Button() With {
                .Text = "OK",
                .Location = New Point(252, 95),
                .DialogResult = DialogResult.OK
            }

            Dim cancelBtn As New Button() With {
                .Text = "Cancel",
                .Location = New Point(333, 95),
                .DialogResult = DialogResult.Cancel
            }

            prompt.Controls.Add(lbl)
            prompt.Controls.Add(txt)
            prompt.Controls.Add(okBtn)
            prompt.Controls.Add(cancelBtn)
            prompt.AcceptButton = okBtn
            prompt.CancelButton = cancelBtn

            Dim result As DialogResult = prompt.ShowDialog(Me)
            If result = DialogResult.OK Then
                Return txt.Text
            End If
        End Using

        Return Nothing
    End Function

    Private Function EnsureAdminOverrideAuthorized(targetCheckBox As CheckBox) As Boolean
        If targetCheckBox Is Nothing Then
            Return False
        End If

        If isRevertingUnitOverride Then
            Return False
        End If

        ' Programmatic SQL/UI state loads should not be blocked by the admin prompt.
        If Not targetCheckBox.Focused Then
            Return True
        End If

        Dim enteredPassword As String = PromptOverridePassword()
        If enteredPassword Is Nothing Then
            isRevertingUnitOverride = True
            Try
                targetCheckBox.Checked = Not targetCheckBox.Checked
            Finally
                isRevertingUnitOverride = False
            End Try
            Return False
        End If

        Dim adminPasswords() As String = {"Wpower84", "Admin123"}
        Dim isAdmin As Boolean = adminPasswords.Any(Function(p) String.Equals(enteredPassword, p, StringComparison.Ordinal))
        If isAdmin Then
            Return True
        End If

        MessageBox.Show("Incorrect Password, logged in as normal user")
        isRevertingUnitOverride = True
        Try
            targetCheckBox.Checked = Not targetCheckBox.Checked
        Finally
            isRevertingUnitOverride = False
        End Try
        Return False
    End Function

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEECO0023" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox1.Checked
        If Not EnsureAdminOverrideAuthorized(CheckBox1) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox2) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox3) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox4) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox5) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox6) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox7) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox8) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox9) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox10) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox11) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox12) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox13) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox14) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox15) Then Return
        UpdateDatabase(studentID, columnName, newValue)
        UpdateLabelsFromDatabase(studentID)
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
    End Sub
    Private Sub CheckBox16_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox16.CheckedChanged
        Dim studentID As String = MainFrm.StudentIDLBL.Text ' Replace with the actual student ID
        Dim columnName As String = "UEEEL0009" ' Replace with the corresponding column name
        Dim newValue As Boolean = CheckBox16.Checked
        If Not EnsureAdminOverrideAuthorized(CheckBox16) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox17) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox18) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox19) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox20) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox21) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox22) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox23) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox24) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox25) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox26) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox27) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox28) Then Return
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
        If Not EnsureAdminOverrideAuthorized(CheckBox29) Then Return
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
    Private Async Sub StudentUnits_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        VersionLBL.Text = MainFrm.VersionLBL.Text
        SelectedStudentLBL.Text = MainFrm.SelectedStudentLBL.Text
        InitializeUnitCheckBoxMap()
        CompletionChecker.UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        CompletionChecker.LoadCheckBoxStates(MainFrm.StudentIDLBL.Text)
        UpdateLabelsFromDatabase(MainFrm.StudentIDLBL.Text)
        MainFrm.UnitAlertLbl.Refresh()
        UnitAlertLbl1.Refresh()
        UpdateButtonVisibility()
        PopulateTeacherCombo()
        UpdateLabelWithDatabaseDate()
        If ExemplarProfilingApi.IsConfigured() Then
            SetProfilingSummary("Ready to refresh per-unit profiling percentages.", Color.DarkGreen)
        Else
            SetProfilingSummary("Profiling API is not configured for this installation.", Color.DarkOrange)
        End If
        '-------------------------------------------------------------
        'Enable below once Address column is working in SQL database
        TextBox1.Text = MainFrm.Label28.Text
        '-------------------------------------------------------------
        If CheckBox27.Checked And CheckBox29.Checked Then
            CheckBox41.Visible = True

        Else
            CheckBox41.Visible = False

        End If

        ' Auto-load per-unit profiling percentages when the form opens.
        If ExemplarProfilingApi.IsConfigured() AndAlso unitCheckBoxes.Count > 0 Then
            Try
                Cursor = Cursors.WaitCursor
                Dim result As ExemplarUnitProgressResult = Await ExemplarProfilingApi.GetStudentUnitProgressAsync(
                    MainFrm.StudentFirstnameLBL.Text,
                    MainFrm.StudentSurnameLBL.Text,
                    MainFrm.StudentEmailLBL.Text,
                    "",
                    unitCheckBoxes.Keys,
                    New String() {
                        "acab955d-09f4-4d04-ae6e-a6dc463a1e48",
                        "f7e89709-7528-446b-a71f-3cecbbd911b2"
                    }
                )

                If result IsNot Nothing AndAlso result.IsSuccessful Then
                    ApplyUnitProfilingAnnotations(result)
                    SetProfilingSummary("Profiling percentages loaded.", Color.DarkGreen)
                Else
                    SetProfilingSummary("Student qualification request failed: " &
                                         If(result IsNot Nothing, result.ErrorMessage, "Unknown error"), Color.Maroon)
                End If
            Catch ex As Exception
                SetProfilingSummary("Student qualification request failed: " & ex.Message, Color.Maroon)
            Finally
                Cursor = Cursors.Default
            End Try
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

    Private Sub InitializeUnitCheckBoxMap()
        If unitCheckBoxes.Count > 0 Then
            Return
        End If

        For Each cb As CheckBox In {
            CheckBox1, CheckBox2, CheckBox3, CheckBox4, CheckBox5, CheckBox6, CheckBox7, CheckBox8, CheckBox9, CheckBox10,
            CheckBox11, CheckBox12, CheckBox13, CheckBox14, CheckBox15, CheckBox16, CheckBox17, CheckBox18, CheckBox19, CheckBox20,
            CheckBox21, CheckBox22, CheckBox23, CheckBox24, CheckBox25, CheckBox26, CheckBox27, CheckBox28, CheckBox29
        }
            originalCheckBoxText(cb) = cb.Text
            Dim unitCode As String = ExtractUnitCode(cb.Text)
            If Not String.IsNullOrWhiteSpace(unitCode) AndAlso Not unitCheckBoxes.ContainsKey(unitCode) Then
                unitCheckBoxes(unitCode) = cb
            End If
        Next
    End Sub

    Private Function ExtractUnitCode(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then
            Return ""
        End If

        Dim parts() As String = text.Split("-"c)
        If parts.Length = 0 Then
            Return ""
        End If

        Return parts(0).Trim().ToUpperInvariant()
    End Function

    Private Function TryGetUnitLabel(unitCode As String, suffix As String) As Label
        If String.IsNullOrWhiteSpace(unitCode) OrElse String.IsNullOrWhiteSpace(suffix) Then
            Return Nothing
        End If

        Dim labelName As String = unitCode.Trim().ToUpperInvariant() & suffix.Trim().ToUpperInvariant()
        Dim found As Control() = Me.Controls.Find(labelName, True)
        If found IsNot Nothing AndAlso found.Length > 0 Then
            Return TryCast(found(0), Label)
        End If

        Return Nothing
    End Function

    Private Iterator Function GetAllLabelsRecursively(parent As Control) As IEnumerable(Of Label)
        For Each child As Control In parent.Controls
            Dim lbl As Label = TryCast(child, Label)
            If lbl IsNot Nothing Then
                Yield lbl
            End If

            For Each nested As Label In GetAllLabelsRecursively(child)
                Yield nested
            Next
        Next
    End Function

    Private Function TryParsePercent(text As String) As Nullable(Of Double)
        If String.IsNullOrWhiteSpace(text) Then
            Return Nothing
        End If

        Dim t As String = text.Trim()
        If String.Equals(t, "N/A", StringComparison.OrdinalIgnoreCase) Then
            Return Nothing
        End If

        If t.EndsWith("%"c) Then
            t = t.Substring(0, t.Length - 1).Trim()
        End If

        Dim value As Double
        If Double.TryParse(t, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, value) Then
            Return value
        End If

        Return Nothing
    End Function

    Private Sub UpdateAveragePercentageLabel()
        ' Label13 is created in the Designer by you.
        Dim found As Control() = Me.Controls.Find("Label13", True)
        If found Is Nothing OrElse found.Length = 0 Then
            Return
        End If

        Dim avgLbl As Label = TryCast(found(0), Label)
        If avgLbl Is Nothing Then
            Return
        End If

        Dim expSum As Double = 0
        Dim expCount As Integer = 0
        Dim demoSum As Double = 0
        Dim demoCount As Integer = 0

        For Each lbl As Label In GetAllLabelsRecursively(Me)
            If lbl Is Nothing OrElse String.IsNullOrWhiteSpace(lbl.Name) Then
                Continue For
            End If

            Dim val As Nullable(Of Double) = TryParsePercent(lbl.Text)
            If Not val.HasValue Then
                Continue For
            End If

            If lbl.Name.EndsWith("EXP", StringComparison.OrdinalIgnoreCase) Then
                expSum += val.Value
                expCount += 1
            ElseIf lbl.Name.EndsWith("DEMO", StringComparison.OrdinalIgnoreCase) Then
                demoSum += val.Value
                demoCount += 1
            End If
        Next

        If expCount = 0 AndAlso demoCount = 0 Then
            avgLbl.Text = "N/A"
            Return
        End If

        Dim avgExp As Double = If(expCount > 0, expSum / expCount, Double.NaN)
        Dim avgDemo As Double = If(demoCount > 0, demoSum / demoCount, Double.NaN)

        Dim combinedAvg As Double
        If expCount > 0 AndAlso demoCount > 0 Then
            combinedAvg = (avgExp + avgDemo) / 2.0
        ElseIf expCount > 0 Then
            combinedAvg = avgExp
        Else
            combinedAvg = avgDemo
        End If

        avgLbl.Text = $"{CInt(Math.Round(combinedAvg, 0))}%"
    End Sub

    Private Sub SetProfilingSummary(message As String, color As Color)
        ProfilingSummaryLbl.Text = message
        ProfilingSummaryLbl.ForeColor = color
    End Sub

    Private Sub ResetProfilingAnnotations()
        For Each pair In originalCheckBoxText
            pair.Key.Text = pair.Value
        Next

        ' Reset the experience/demonstration labels for all known units.
        For Each code In unitCheckBoxes.Keys
            Dim expLbl As Label = TryGetUnitLabel(code, "EXP")
            If expLbl IsNot Nothing Then expLbl.Text = "N/A"

            Dim demoLbl As Label = TryGetUnitLabel(code, "DEMO")
            If demoLbl IsNot Nothing Then demoLbl.Text = "N/A"
        Next

        UpdateAveragePercentageLabel()
    End Sub

    Private Sub ApplyUnitProfilingAnnotations(progressResult As ExemplarUnitProgressResult)
        ResetProfilingAnnotations()

        ' Iterate the API's returned unit keys directly to avoid any key casing mismatch.
        For Each pair In progressResult.Units
            Dim code As String = pair.Key
            Dim unitInfo As ExemplarUnitProgressItem = pair.Value

            Dim percentText As String = "N/A"
            If unitInfo.Percentage.HasValue Then
                percentText = unitInfo.Percentage.Value.ToString("0.##") & "%"
            End If

            Dim cardsText As String = ""
            If unitInfo.CompletedCards.HasValue AndAlso unitInfo.TotalCards.HasValue AndAlso unitInfo.TotalCards.Value > 0 Then
                cardsText = $" ({unitInfo.CompletedCards.Value}/{unitInfo.TotalCards.Value} cards)"
            End If

            Dim expPct As String = If(unitInfo.ExperienceCardsPercentage.HasValue,
                                      CInt(Math.Round(unitInfo.ExperienceCardsPercentage.Value, 0)).ToString() & "%",
                                      TryComputeCardsPercent(unitInfo.ExperienceCardsRawJson))
            Dim demoPct As String = If(unitInfo.DemonstrationCardsPercentage.HasValue,
                                       CInt(Math.Round(unitInfo.DemonstrationCardsPercentage.Value, 0)).ToString() & "%",
                                       TryComputeCardsPercent(unitInfo.DemonstrationCardsRawJson))

            ' Update the designer labels.
            Dim expLbl As Label = TryGetUnitLabel(code, "EXP")
            If expLbl IsNot Nothing Then expLbl.Text = expPct

            Dim demoLbl As Label = TryGetUnitLabel(code, "DEMO")
            If demoLbl IsNot Nothing Then demoLbl.Text = demoPct
        Next

        UpdateAveragePercentageLabel()
    End Sub

    Private Sub WriteProfilingDebugDump(rawJson As String)
        Try
            Dim dumpPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExemplarUnitProfilingDebug.txt")
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("Student Units Profiling Debug Dump")
            sb.AppendLine("Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            sb.AppendLine()
            sb.AppendLine(rawJson)

            File.WriteAllText(dumpPath, sb.ToString(), System.Text.Encoding.UTF8)
            SetProfilingSummary("Profiling debug dump written to " & dumpPath, Color.DarkGreen)
            Try
                Process.Start(New ProcessStartInfo() With {
                    .FileName = dumpPath,
                    .UseShellExecute = True
                })
            Catch ex As Exception
                SetProfilingSummary("Profiling debug dump written, but could not open it automatically: " & ex.Message, Color.DarkOrange)
            End Try
        Catch ex As Exception
            SetProfilingSummary("Failed to write profiling debug dump: " & ex.Message, Color.Maroon)
        End Try
    End Sub

    Private Sub WriteProfilingDebugDump(result As ExemplarUnitProgressResult)
        Try
            Dim dumpPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExemplarUnitProfilingDebug.txt")
            Dim sb As New System.Text.StringBuilder()

            sb.AppendLine("Student Units Profiling Debug Dump")
            sb.AppendLine("Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            sb.AppendLine()

            For Each pair In result.Units
                Dim unitCode As String = pair.Key
                Dim unitItem As ExemplarUnitProgressItem = pair.Value

                sb.AppendLine("=== UNIT " & unitCode & " ===")

                Dim expPct As String = If(unitItem.ExperienceCardsPercentage.HasValue,
                                          CInt(Math.Round(unitItem.ExperienceCardsPercentage.Value, 0)).ToString() & "%",
                                          TryComputeCardsPercent(unitItem.ExperienceCardsRawJson))
                Dim demoPct As String = If(unitItem.DemonstrationCardsPercentage.HasValue,
                                           CInt(Math.Round(unitItem.DemonstrationCardsPercentage.Value, 0)).ToString() & "%",
                                           TryComputeCardsPercent(unitItem.DemonstrationCardsRawJson))

                sb.AppendLine("=== EXPERIENCE_CARDS ===" & expPct)
                If String.Equals(expPct, "N/A", StringComparison.OrdinalIgnoreCase) Then
                    Dim hasRaw As Boolean = Not String.IsNullOrWhiteSpace(unitItem.ExperienceCardsRawJson)
                    sb.AppendLine("  Debug(exp): " & If(hasRaw, "raw_present", "raw_empty"))
                End If
                sb.AppendLine("=== DEMONSTRATION_CARDS ===" & demoPct)
                If String.Equals(demoPct, "N/A", StringComparison.OrdinalIgnoreCase) Then
                    Dim hasRaw As Boolean = Not String.IsNullOrWhiteSpace(unitItem.DemonstrationCardsRawJson)
                    sb.AppendLine("  Debug(demo): " & If(hasRaw, "raw_present", "raw_empty"))
                End If

                ' Debug one known unit to inspect what the /progression/cards endpoint returns.
                If String.Equals(unitCode, "UEECD0007", StringComparison.OrdinalIgnoreCase) Then
                    Dim cardsRaw As String = unitItem.ProgressionCardsEndpointRawJson
                    Dim hasCardsEndpoint As Boolean = Not String.IsNullOrWhiteSpace(cardsRaw)
                    sb.AppendLine("  Debug(cards_endpoint): " & If(hasCardsEndpoint, "raw_present", "raw_empty"))
                    If hasCardsEndpoint Then
                        Dim snippetLen As Integer = Math.Min(600, cardsRaw.Length)
                        sb.AppendLine("  Debug(cards_endpoint_snippet): " & cardsRaw.Substring(0, snippetLen))
                    End If
                End If
                sb.AppendLine()
            Next

            File.WriteAllText(dumpPath, sb.ToString(), System.Text.Encoding.UTF8)
            SetProfilingSummary("Profiling debug dump written to " & dumpPath, Color.DarkGreen)

            Try
                Process.Start(New ProcessStartInfo() With {
                    .FileName = dumpPath,
                    .UseShellExecute = True
                })
            Catch ex As Exception
                SetProfilingSummary("Profiling debug dump written, but could not open it automatically: " & ex.Message, Color.DarkOrange)
            End Try
        Catch ex As Exception
            SetProfilingSummary("Failed to write profiling debug dump: " & ex.Message, Color.Maroon)
        End Try
    End Sub

    Private Function TryComputeCardsPercent(cardsRawJson As String) As String
        Try
            If String.IsNullOrWhiteSpace(cardsRawJson) Then Return "N/A"

            Using doc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(cardsRawJson)
                Dim root = doc.RootElement

                Dim total As Integer = 0
                Dim complete As Integer = 0
                Dim statusesFound As Integer = 0

                Dim cardsArrayFound As Boolean = False

                If root.ValueKind = System.Text.Json.JsonValueKind.Array Then
                    cardsArrayFound = True
                    For Each item In root.EnumerateArray()
                        total += 1
                        Dim status As String = ExtractCardStatus(item)
                        If Not String.IsNullOrWhiteSpace(status) Then statusesFound += 1
                        If IsCardCompleteStatus(status) Then complete += 1
                    Next
                Else
                    ' Some API responses wrap the array inside an object, e.g. { "cards": [ ... ] }
                    ' Walk the tree to find the first array and compute percentages from it.
                    cardsArrayFound = TryComputeCardsPercentFromFirstArray(root, total, complete, statusesFound, 0)
                End If

                If Not cardsArrayFound Then Return "N/A"

                If total = 0 Then Return "N/A"

                ' If the API only returns card_id/user_id (no status/completion field),
                ' treat all returned cards as complete.
                If statusesFound = 0 Then complete = total

                Dim pct As Integer = CInt(Math.Round((complete / total) * 100D))
                Return pct.ToString() & "%"
            End Using
        Catch
            Return "N/A"
        End Try
    End Function

    Private Function TryComputeCardsPercentFromFirstArray(root As System.Text.Json.JsonElement,
                                                           ByRef total As Integer,
                                                           ByRef complete As Integer,
                                                           ByRef statusesFound As Integer,
                                                           depth As Integer) As Boolean
        If depth > 10 Then Return False

        Select Case root.ValueKind
            Case System.Text.Json.JsonValueKind.Array
                total = 0
                complete = 0
                statusesFound = 0
                For Each item In root.EnumerateArray()
                    total += 1
                    Dim status As String = ExtractCardStatus(item)
                    If Not String.IsNullOrWhiteSpace(status) Then statusesFound += 1
                    If IsCardCompleteStatus(status) Then complete += 1
                Next
                Return True
            Case System.Text.Json.JsonValueKind.Object
                For Each prop In root.EnumerateObject()
                    If TryComputeCardsPercentFromFirstArray(prop.Value, total, complete, statusesFound, depth + 1) Then Return True
                Next
        End Select

        Return False
    End Function

    Private Function ExtractCardStatus(card As System.Text.Json.JsonElement) As String
        Try
            If card.ValueKind = System.Text.Json.JsonValueKind.String Then
                Return card.GetString()
            End If

            If card.ValueKind <> System.Text.Json.JsonValueKind.Object Then
                Return ""
            End If

            ' Common status field names.
            For Each key In New String() {"status", "card_status", "state", "completion_status", "completionStatus", "cardStatus"}
                Dim v As System.Text.Json.JsonElement
                If card.TryGetProperty(key, v) Then
                    If v.ValueKind = System.Text.Json.JsonValueKind.String Then
                        Return v.GetString()
                    End If
                End If
            Next

            ' Fallback for boolean approval/completion flags.
            For Each boolKey In New String() {"approved", "is_approved", "complete", "is_complete", "completed", "is_completed"}
                Dim v As System.Text.Json.JsonElement
                If card.TryGetProperty(boolKey, v) Then
                    If v.ValueKind = System.Text.Json.JsonValueKind.True Then
                        Return boolKey
                    End If
                End If
            Next

            Return ""
        Catch
            Return ""
        End Try
    End Function

    Private Function IsCardCompleteStatus(status As String) As Boolean
        If String.IsNullOrWhiteSpace(status) Then Return False
        Dim s As String = status.Trim().ToUpperInvariant()
        Return s.Contains("APPROVED") OrElse s.Contains("COMPLETE") OrElse s.Contains("COMPLETED") OrElse
               s.Contains("SUBMITTED") OrElse s.Contains("VERIFIED") OrElse s.Contains("SATISFIED") OrElse
               s.Contains("PASSED")
    End Function

    Private Sub SetControlVisibilityByName(controlName As String, visible As Boolean)
        If String.IsNullOrWhiteSpace(controlName) Then
            Return
        End If

        Dim found As Control() = Me.Controls.Find(controlName, True)
        If found Is Nothing OrElse found.Length = 0 Then
            Return
        End If

        found(0).Visible = visible
    End Sub

    Private Function TryGetLabel13PercentValue() As Nullable(Of Double)
        Dim found As Control() = Me.Controls.Find("Label13", True)
        If found Is Nothing OrElse found.Length = 0 Then
            Return Nothing
        End If

        Dim avgLbl As Label = TryCast(found(0), Label)
        If avgLbl Is Nothing Then
            Return Nothing
        End If

        Return TryParsePercent(avgLbl.Text)
    End Function

    Private Sub ApplyLowAverageOverrideForCheckbox28And29()
        Dim avgPercent As Nullable(Of Double) = TryGetLabel13PercentValue()
        Dim isBelow85 As Boolean = avgPercent.HasValue AndAlso avgPercent.Value < 85.0
        Dim triggeredBy28 As Boolean = CheckBox28.Checked AndAlso isBelow85
        Dim triggeredBy29 As Boolean = CheckBox29.Checked AndAlso isBelow85

        ' Toggle notice labels regardless so old state does not stick.
        SetControlVisibilityByName("Label14", triggeredBy28)
        SetControlVisibilityByName("Label15", triggeredBy29)

        If triggeredBy28 OrElse triggeredBy29 Then
            ' Your override: keep existing checks, then force-hide/show these controls.
            Button1.Visible = False
            Button2.Visible = False
            Label2.Visible = False
            TextBox1.Visible = False
            ComboBox1.Visible = False
            Label5.Visible = False
            CheckBox40.Visible = False
            CheckBox41.Visible = False
            Button3.Visible = True
        Else
            ' Revert only controls driven by this override.
            Button3.Visible = False
        End If
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
            Button2.Visible = True

        Else
            Button1.Visible = False
            Button2.Visible = False

        End If

        ' Keep existing behavior above; apply extra rule for checkbox 28/29 + low average.
        ApplyLowAverageOverrideForCheckbox28And29()
    End Sub
    ''' <summary>Collects the form field position (page + rectangle) for a stamp so we can draw the image on top after flattening.</summary>
    Private Sub CollectStampPosition(form As AcroFields, fieldName As String, outPositions As List(Of Tuple(Of Integer, Rectangle)))
        Try
            Dim positions = form.GetFieldPositions(fieldName)
            If positions Is Nothing OrElse positions.Count = 0 Then Return
            Dim fp As AcroFields.FieldPosition = CType(positions(0), AcroFields.FieldPosition)
            outPositions.Add(Tuple.Create(fp.page, fp.position))
        Catch
            ' Field may not exist or have different name (e.g. "Stamp 1"); try partial name match
            Try
                For Each key As String In form.Fields.Keys
                    If key.IndexOf("Stamp1", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso fieldName = "Stamp1" Then
                        Dim positions = form.GetFieldPositions(key)
                        If positions IsNot Nothing AndAlso positions.Count > 0 Then
                            Dim fp As AcroFields.FieldPosition = CType(positions(0), AcroFields.FieldPosition)
                            outPositions.Add(Tuple.Create(fp.page, fp.position))
                            Return
                        End If
                    End If
                    If key.IndexOf("Stamp2", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso fieldName = "Stamp2" Then
                        Dim positions = form.GetFieldPositions(key)
                        If positions IsNot Nothing AndAlso positions.Count > 0 Then
                            Dim fp As AcroFields.FieldPosition = CType(positions(0), AcroFields.FieldPosition)
                            outPositions.Add(Tuple.Create(fp.page, fp.position))
                            Return
                        End If
                    End If
                Next
            Catch
                ' Ignore
            End Try
        End Try
    End Sub

    Private Sub CollectFieldPosition(form As AcroFields, fieldName As String, outPositions As List(Of Tuple(Of Integer, Rectangle)))
        Try
            Dim positions = form.GetFieldPositions(fieldName)
            If positions IsNot Nothing AndAlso positions.Count > 0 Then
                Dim fp As AcroFields.FieldPosition = CType(positions(0), AcroFields.FieldPosition)
                outPositions.Add(Tuple.Create(fp.page, fp.position))
                Return
            End If
        Catch
            ' Fall through to partial-match scan.
        End Try

        Try
            For Each key As String In form.Fields.Keys
                If key.IndexOf(fieldName, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Dim positions = form.GetFieldPositions(key)
                    If positions IsNot Nothing AndAlso positions.Count > 0 Then
                        Dim fp As AcroFields.FieldPosition = CType(positions(0), AcroFields.FieldPosition)
                        outPositions.Add(Tuple.Create(fp.page, fp.position))
                        Return
                    End If
                End If
            Next
        Catch
            ' Ignore position lookup errors.
        End Try
    End Sub

    Private Function PromptForImageFile(dialogTitle As String) As String
        Using ofd As New OpenFileDialog()
            ofd.Title = dialogTitle
            ofd.Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp;*.gif|All Files|*.*"
            ofd.Multiselect = False
            If ofd.ShowDialog() = DialogResult.OK Then
                Return ofd.FileName
            End If
        End Using
        Return ""
    End Function

    Private Function GetVuStampPath(templatePath As String) As String
        ' Fixed asset: app should always use VU Stamp image from app/project roots.
        Dim candidates As New List(Of String)()
        Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
        Dim rootTemplateDir As String = If(String.IsNullOrWhiteSpace(templatePath), "", System.IO.Path.GetDirectoryName(templatePath))

        candidates.Add(System.IO.Path.Combine(baseDir, "VU Stamp.png"))
        candidates.Add(System.IO.Path.Combine(baseDir, "VU Stamp.jpg"))

        Dim parent As String = System.IO.Path.GetDirectoryName(baseDir.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar))
        For i As Integer = 0 To 5
            If String.IsNullOrWhiteSpace(parent) Then Exit For
            candidates.Add(System.IO.Path.Combine(parent, "VU Stamp.png"))
            candidates.Add(System.IO.Path.Combine(parent, "VU Stamp.jpg"))
            parent = System.IO.Path.GetDirectoryName(parent)
        Next

        If Not String.IsNullOrWhiteSpace(rootTemplateDir) Then
            candidates.Add(System.IO.Path.Combine(rootTemplateDir, "VU Stamp.png"))
            candidates.Add(System.IO.Path.Combine(rootTemplateDir, "VU Stamp.jpg"))
        End If

        For Each path As String In candidates
            If Not String.IsNullOrWhiteSpace(path) AndAlso File.Exists(path) Then
                Return path
            End If
        Next

        Return ""
    End Function

    Private Sub AddImageToPositions(stamper As PdfStamper, imagePath As String, positions As List(Of Tuple(Of Integer, Rectangle)))
        If String.IsNullOrWhiteSpace(imagePath) OrElse Not File.Exists(imagePath) Then
            Return
        End If
        If positions Is Nothing OrElse positions.Count = 0 Then
            Return
        End If

        For Each pos In positions
            Dim img As Image = Image.GetInstance(imagePath)
            Dim pageNum As Integer = pos.Item1
            Dim rect As Rectangle = pos.Item2
            img.ScaleToFit(rect.Width, rect.Height)
            img.SetAbsolutePosition(rect.Left + (rect.Width - img.ScaledWidth) / 2,
                                    rect.Bottom + (rect.Height - img.ScaledHeight) / 2)
            stamper.GetOverContent(pageNum).AddImage(img)
        Next
    End Sub

    Public Function PopulatePdfWithParameters(Optional openGeneratedPdf As Boolean = True,
                                              Optional openArchiveFolder As Boolean = True) As String
        Dim templatePath As String = "LEATemplate.pdf"
        Dim outputDirectory As String = "P:\VUPoly\MT&T\IT, Electrical And Engineering\Submitted LEA Authorisation Forms"
        Dim Todaydate As String = DateTime.Today.ToString("ddMMyyyy")
        Dim fileName As String = MainFrm.StudentIDLBL.Text & "_" & MainFrm.StudentFirstnameLBL.Text & "_" & MainFrm.StudentSurnameLBL.Text & "_" & Todaydate & ".pdf"
        Dim outputPath As String = System.IO.Path.Combine(outputDirectory, fileName)

        Try
            If Not Directory.Exists(outputDirectory) Then
                Directory.CreateDirectory(outputDirectory)
            End If

            ' Pass 1: fill form and flatten to memory so stamps can be drawn on top in pass 2
            Dim flattenedBytes As Byte()
            Using reader As New PdfReader(templatePath)
                Using memStream As New MemoryStream()
                    Using stamper As New PdfStamper(reader, memStream)
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
                    flattenedBytes = memStream.ToArray()
                End Using
            End Using

            ' Pass 2: get stamp positions from original, then add stamp images on top of flattened PDF
            Dim stampPath As String = GetVuStampPath(templatePath)

            Dim signaturePath As String = ""
            If CheckBox40.Checked OrElse CheckBox41.Checked Then
                signaturePath = PromptForImageFile("Select signature image (used for Signature and Signature_2)")
            End If

            Dim stampPositions As New List(Of Tuple(Of Integer, Rectangle))
            Dim signature1Positions As New List(Of Tuple(Of Integer, Rectangle))
            Dim signature2Positions As New List(Of Tuple(Of Integer, Rectangle))
            Using readerForPos As New PdfReader(templatePath)
                Using stamperPos As New PdfStamper(readerForPos, New MemoryStream())
                    Dim formPos As AcroFields = stamperPos.AcroFields
                    If CheckBox40.Checked Then CollectStampPosition(formPos, "Stamp1", stampPositions)
                    If CheckBox41.Checked Then CollectStampPosition(formPos, "Stamp2", stampPositions)
                    If CheckBox40.Checked Then CollectFieldPosition(formPos, "Signature", signature1Positions)
                    If CheckBox41.Checked Then CollectFieldPosition(formPos, "Signature_2", signature2Positions)
                End Using
            End Using

            Using readerFlattened As New PdfReader(flattenedBytes)
                Using stamper2 As New PdfStamper(readerFlattened, New FileStream(outputPath, FileMode.Create))
                    AddImageToPositions(stamper2, stampPath, stampPositions)
                    AddImageToPositions(stamper2, signaturePath, signature1Positions)
                    AddImageToPositions(stamper2, signaturePath, signature2Positions)
                End Using
            End Using

            If openGeneratedPdf Then
                PdfHelper.OpenPdfWithDefaultViewer(outputPath)
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ""
        End Try

        If openArchiveFolder Then
            ' Specify the path to the folder you want to open
            Dim folderPath As String = "P:\VUPoly\MT&T\IT, Electrical and Engineering\Submitted LEA Authorisation Forms\ARCHIVE\"

            Try
                ' Open the folder using the default file explorer with the specified folder path
                Process.Start("explorer.exe", $"/select,""{folderPath}""")
            Catch ex As Exception
                ' Handle any errors that may occur
                MessageBox.Show("Error opening folder: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If

        Return outputPath
    End Function

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If Not String.IsNullOrEmpty(TextBox1.Text) AndAlso (CheckBox40.Checked OrElse CheckBox41.Checked) Then
            Button1.Visible = True
            Button2.Visible = True
        Else
            Button1.Visible = False
            Button2.Visible = False
        End If
    End Sub

    Private Async Function ConfirmProfilingIfStudentNotFoundAsync() As Task(Of Boolean)
        Dim shouldPrompt As Boolean = False
        If ExemplarProfilingApi.IsConfigured() Then
            Try
                Dim lookup As ExemplarProfileLookupResult = Await ExemplarProfilingApi.LookupStudentProfileAsync(
                    MainFrm.StudentFirstnameLBL.Text,
                    MainFrm.StudentSurnameLBL.Text,
                    MainFrm.StudentEmailLBL.Text
                )
                If lookup IsNot Nothing AndAlso Not lookup.IsSuccessful Then
                    Dim statusText As String = If(lookup.StatusText, "")
                    Dim detailText As String = If(lookup.DetailText, "")
                    If statusText.IndexOf("Student not found", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       detailText.IndexOf("No matching Exemplar student was found", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        shouldPrompt = True
                    End If
                End If
            Catch
                ' If lookup fails unexpectedly, keep existing workflow.
            End Try
        End If

        If Not shouldPrompt Then
            Return True
        End If

        Dim response As DialogResult = MessageBox.Show(
            "Student not found on Profiling API. Is the student's profiling up to date?",
            "Profile Check",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )
        If response <> DialogResult.Yes Then
            MessageBox.Show("Student needs to be up to date with profiling before an LEA Authority Form can be generated", "Profile Status")
            Return False
        End If

        Return True
    End Function

    Private Sub CreateOutlookDraft(toAddress As String, subjectText As String, bodyText As String, Optional attachmentPath As String = "")
        Try
            ' Match the existing working approach used elsewhere in the app.
            Dim OutApp As Object = CreateObject("Outlook.Application")
            Dim OutMail As Object = OutApp.CreateItem(0) ' olMailItem

            With OutMail
                .To = toAddress
                .CC = "electrotechnology.admin@vu.edu.au"
                .Subject = subjectText
                .Body = bodyText
                If Not String.IsNullOrWhiteSpace(attachmentPath) AndAlso File.Exists(attachmentPath) Then
                    .Attachments.Add(attachmentPath)
                End If
                .Display()
            End With

            OutMail = Nothing
            OutApp = Nothing
        Catch ex As Exception
            MessageBox.Show("Unable to create Outlook email draft: " & ex.Message, "Email Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not Await ConfirmProfilingIfStudentNotFoundAsync() Then
            Exit Sub
        End If

        PopulatePdfWithParameters()
    End Sub

    Private Async Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Not Await ConfirmProfilingIfStudentNotFoundAsync() Then
            Exit Sub
        End If

        Dim pdfPath As String = PopulatePdfWithParameters()
        If String.IsNullOrWhiteSpace(pdfPath) Then
            Return
        End If

        Dim studentName As String = (MainFrm.StudentFirstnameLBL.Text & " " & MainFrm.StudentSurnameLBL.Text).Trim()
        Dim subjectText As String = "Congratulations - LEA Authorisation Form"
        Dim bodyText As String =
            "Hi " & studentName & "," & Environment.NewLine & Environment.NewLine &
            "Congratulations. Please find attached your completed LEA Authorisation Form." & Environment.NewLine & Environment.NewLine &
            "Kind regards," & Environment.NewLine &
            "Electrotechnology Administration"

        CreateOutlookDraft(MainFrm.StudentEmailLBL.Text, subjectText, bodyText, pdfPath)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim studentName As String = (MainFrm.StudentFirstnameLBL.Text & " " & MainFrm.StudentSurnameLBL.Text).Trim()
        Dim subjectText As String = "Profiling Outcome - Insufficient Cards"
        Dim bodyText As String =
            "Hi " & studentName & "," & Environment.NewLine & Environment.NewLine &
            "Your profiling cards currently do not meet the required thresholds for progression." & Environment.NewLine &
            "Requirement: 85% / 100% (system allows 99% where applicable)." & Environment.NewLine & Environment.NewLine &
            "Please continue submitting profiling cards and contact us if you need assistance." & Environment.NewLine & Environment.NewLine &
            "Kind regards," & Environment.NewLine &
            "Electrotechnology Administration"

        CreateOutlookDraft(MainFrm.StudentEmailLBL.Text, subjectText, bodyText)
    End Sub

    Private Sub CheckBox40_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox40.CheckedChanged
        UpdateButtonVisibility()
    End Sub

    Private Sub CheckBox41_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox41.CheckedChanged
        UpdateButtonVisibility()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
    End Sub

    Private Async Sub RefreshProfilingBtn_Click(sender As Object, e As EventArgs) Handles RefreshProfilingBtn.Click
        InitializeUnitCheckBoxMap()
        Dim studentId As String = MainFrm.StudentIDLBL.Text?.Trim()
        Dim firstName As String = MainFrm.StudentFirstnameLBL.Text?.Trim()
        Dim lastName As String = MainFrm.StudentSurnameLBL.Text?.Trim()
        Dim agreementEmail As String = MainFrm.StudentEmailLBL.Text?.Trim()
        ExemplarEmailOverrides.InvalidateCacheForStudent(studentId)
        Dim overrideEmail As String = ExemplarEmailOverrides.GetOverride(studentId)
        Dim lookupEmail As String = If(String.IsNullOrWhiteSpace(overrideEmail), agreementEmail, overrideEmail)

        If unitCheckBoxes.Count = 0 Then
            SetProfilingSummary("No unit checkboxes were found to update.", Color.Maroon)
            Return
        End If

        RefreshProfilingBtn.Enabled = False
        Cursor = Cursors.WaitCursor
        SetProfilingSummary("Loading student qualification data...", Color.SteelBlue)

        Try
            Dim result As ExemplarUnitProgressResult = Await ExemplarProfilingApi.GetStudentUnitProgressAsync(
                firstName,
                lastName,
                lookupEmail,
                "",
                unitCheckBoxes.Keys,
                New String() {
                    "acab955d-09f4-4d04-ae6e-a6dc463a1e48",
                    "f7e89709-7528-446b-a71f-3cecbbd911b2"
                }
            )

            If result.IsSuccessful Then
                ApplyUnitProfilingAnnotations(result)
                WriteProfilingDebugDump(result)
                SetProfilingSummary("Student qualification data written to debug file.", Color.DarkGreen)
            Else
                WriteProfilingDebugDump("Student qualification request failed: " & result.ErrorMessage)
                SetProfilingSummary("Student qualification request failed: " & result.ErrorMessage, Color.Maroon)
            End If
        Catch ex As Exception
            WriteProfilingDebugDump("Student qualification request failed: " & ex.Message)
            SetProfilingSummary("Student qualification request failed: " & ex.Message, Color.Maroon)
        Finally
            Cursor = Cursors.Default
            RefreshProfilingBtn.Enabled = True
        End Try
    End Sub
End Class