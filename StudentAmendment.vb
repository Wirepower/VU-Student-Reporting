Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports Microsoft.Office.Interop
Imports System.Net.Mail
Imports Microsoft.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports Microsoft.Office.Server.UserProfiles
Imports Microsoft.Office.Server.Search.WebControls

Public Class StudentAmendment
    Private connection As SqlConnection
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Private Sub PopulateTeacherCombo()
        ' Replace "Your_Connection_String_Here" with your actual connection string
        Dim connectionString As String = SQLCon.connectionString

        ' Create and open a SqlConnection
        Using connection As New SqlConnection(connectionString)
            connection.Open()

            ' Clear existing items in the ComboBox
            SubmittedbyCB.Items.Clear()
            ComboBox3.Items.Clear()
            ComboBox2.Items.Clear()

            ' SQL query to retrieve all teachers' names
            Dim query As String = "SELECT Teacher_Full_Name FROM ElectrotechnologyReports.dbo.TeacherList WHERE Highest_Certificate_Taught = 'Certificate III' ORDER BY Teacher_Full_Name ASC"

            ' Create a SqlCommand object with the query and connection
            Using command As New SqlCommand(query, connection)
                ' Execute the query and retrieve the data
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        ' Add each teacher's name to the ComboBox
                        SubmittedbyCB.Items.Add(reader.GetString(0))
                        ComboBox2.Items.Add(reader.GetString(0))
                        ComboBox3.Items.Add(reader.GetString(0))
                    End While
                End Using
            End Using
        End Using
    End Sub
    ' Function to perform the lookup
    Private Sub LookupTeacherEmail()
        ' Initialize the result variable
        Dim teacherID As String = ""

        ' Database connection string
        Dim connectionString As String = SQLCon.connectionString

        ' SQL query to retrieve the email based on the selected teacher full name
        Dim query As String = "SELECT Email FROM ElectrotechnologyReports.dbo.TeacherList WHERE Teacher_Full_Name = @TeacherFullName"

        Try
            ' Create a SqlConnection and SqlCommand objects
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    ' Add parameter for the selected teacher full name
                    command.Parameters.AddWithValue("@TeacherFullName", ComboBox2.Text)

                    ' Open the connection
                    connection.Open()

                    ' Execute the query and retrieve the email
                    Dim result As Object = command.ExecuteScalar()
                    If result IsNot Nothing Then
                        ' Set the retrieved email to Label29.Text
                        Label29.Text = result.ToString()
                    Else
                        ' If no matching teacher found, clear Label29.Text
                        Label29.Text = ""
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' Handle exceptions
            MessageBox.Show("Error retrieving teacher email: " & ex.Message)
        End Try
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'If CheckBox1.Checked Then
        'RemoveStudent = MsgBox("Has this been approved by Management?", vbYesNo)
        'If vbYes Then
        'End If
        'Else
        '   MsgBox("Please seek approval from Management before proceeding")
        'End If
        If CheckBox2.Checked Then
            ' Ask for management approval
            Dim moveApproval As MsgBoxResult = MsgBox("Has this been approved by Management?", vbYesNo, "Moving a Student")
            If moveApproval = vbNo Then
                MsgBox("Please seek approval from Management before proceeding.", vbOK)
                Exit Sub
            End If

            ' Validate the proposed blockgroup student numbers
            Dim studentNumbersValidation As MsgBoxResult = MsgBox("Are the proposed blockgroup " & ComboBox1.Text & " student numbers under 20 students?", vbYesNo, "Moving a Student")
            If studentNumbersValidation = vbNo Then
                MsgBox("Please validate proposed student numbers before proceeding.", vbOK)
                Exit Sub
            End If
        End If

        If CheckBox3.Checked Then
            ' Ask for management approval
            Dim addApproval As MsgBoxResult = MsgBox("Has this been approved by Management?", vbYesNo, "Adding a Student")
            If addApproval = vbNo Then
                MsgBox("Please seek approval from Management before proceeding.", vbOK)
                Exit Sub
            End If

            ' Validate the proposed blockgroup student numbers
            Dim studentNumbersValidation As MsgBoxResult = MsgBox("Are the proposed blockgroup " & ComboBox1.Text & " student numbers under 20 students?", vbYesNo, "Adding a Student")
            If studentNumbersValidation = vbNo Then
                MsgBox("Please validate proposed student numbers before proceeding.", vbOK)
                Exit Sub
            End If
        End If



        Dim missingFields As String = ""
        Dim action As String = ""

        ' Check if both ComboBox and CheckBox are not selected or both are selected
        'If (ComboBox1.SelectedIndex = -1 AndAlso Not CheckBox1.Checked) OrElse
        '(ComboBox1.SelectedIndex <> -1 AndAlso CheckBox1.Checked) Then
        'missingFields &= "- Proposed Blockgroup Change or Delete the student Checkbox." & vbCrLf & "It cant be both!" & vbCrLf
        ' End If


        ' Check which checkbox is checked
        If CheckBox1.Checked Then
            action = "deleted"
        ElseIf CheckBox2.Checked Then
            action = "changed"
        ElseIf CheckBox3.Checked Then
            action = "added"
        End If

        ' Check for missing fields
        If DateTimePicker1.Text = "" Then
            missingFields &= "- Date" & vbCrLf
        End If

        If SubmittedbyCB.Text = "" Then
            missingFields &= "- Submitted By" & vbCrLf
        End If

        ' If there are missing fields, display message and exit sub
        If missingFields <> "" Then
            MessageBox.Show("Please fill the following fields:" & vbCrLf & missingFields)
            Exit Sub
        End If

        Dim OutApp As Object
        Dim OutMail As Object
        Dim body As String
        Dim body1 As String
        Dim AdminEmail As String
        Dim ApptrainEmail As String
        Dim FrankOffer As String
        Dim Trades As String
        Dim imageData As Byte() = RetrieveImageDataFromDatabase()

        ' Retrieve email addresses from the EmailSettings table
        GetEmailAddresses(AdminEmail, ApptrainEmail, FrankOffer, Trades)

        ' Construct the email body for the first email
        body = "Hello, <BR>
    I would like to get the following student blockgroup " & action & ". <BR><BR><BR>
    Student ID: " & Label3.Text & "<BR><BR>
    Student Name: " & Label4.Text & " " & Label5.Text & " <BR><BR>
    Student Email: " & Label6.Text & " <BR><BR>
    From BlockGroup: " & Label20.Text & " to Blockgroup: " & ComboBox1.Text & "<BR><BR>
    Starting from: " & DateTimePicker1.Text & "<BR><BR>
    Submitted by: " & SubmittedbyCB.Text & "<BR><BR>"

        ' Create a new instance of Outlook Application
        OutApp = CreateObject("Outlook.Application")

        ' Create a new email item
        OutMail = OutApp.CreateItem(0)

        ' Set email properties for the first email
        With OutMail
            .To = Trades
            .cc = AdminEmail
            .bcc = ""
            .Subject = "Student Amendment Request - ADD/MOVE/DELETE Student from Blockgroup"
            .HTMLbody = body & "<br><img src='data:image/jpeg;base64," & Convert.ToBase64String(imageData) & "' width='90%'> " & MainFrm.VersionLBL.Text
            .Display ' Display the email
        End With

        ' Clean up
        ReleaseComObject(OutMail)
        ReleaseComObject(OutApp)

        ' Display confirmation message
        MsgBox("Your Email has been Generated!")

        ' Check if the checkbox is checked to generate the second email
        If CheckBox2.Checked Then
            ' Construct the email body for the second email
            body1 = "Dear, " & Label7.Text & " / " & Label4.Text & "<BR> 
        I am writing to inform you that your apprentice, " & Label4.Text & " " & Label5.Text & ", will be transitioning to class " & ComboBox1.Text & "<BR>
        The reason for this change: " & ComboBox4.Text & "<BR>
        Classes Commence for this class on: " & DateTimePicker1.Text & " at 8:00AM<BR><BR>
        Should you have any questions or require further clarification, please feel free to contact me by replying to this email.<BR><BR>" &
        "Best regards,<BR>" & SubmittedbyCB.Text


            ' Create a new instance of Outlook Application
            OutApp = CreateObject("Outlook.Application")

            ' Create a new email item
            OutMail = OutApp.CreateItem(0)

            ' Set email properties for the second email
            With OutMail
                .To = Label6.Text
                .cc = Label10.Text
                .bcc = ""
                .Subject = "Notice: Change of Apprentice Class and/or Day"
                .HTMLbody = body1 & "<br><img src='data:image/jpeg;base64," & Convert.ToBase64String(imageData) & "' width='90%'> " & MainFrm.VersionLBL.Text
                .Display ' Display the email
            End With

            ' Clean up
            ReleaseComObject(OutMail)
            ReleaseComObject(OutApp)

            ' Display confirmation message
            MsgBox("Employer/Student Email has been Generated!")
        End If
    End Sub


    ' Method to release COM objects to prevent memory leaks
    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            End If
        Catch ex As Exception
            Console.WriteLine("Error releasing COM object: " & ex.Message)
        Finally
            obj = Nothing
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub StudentAmendment_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label3.Text = MainFrm.StudentIDLBL.Text
        Label4.Text = MainFrm.StudentFirstnameLBL.Text
        Label5.Text = MainFrm.StudentSurnameLBL.Text
        Label6.Text = MainFrm.StudentEmailLBL.Text
        Label7.Text = MainFrm.EmployerFirstnameLBL.Text
        Label8.Text = MainFrm.EmployerSurnameLBL.Text
        Label9.Text = MainFrm.EmployerBusinessNameLBL.Text
        Label10.Text = MainFrm.EmployerEmailLBL.Text
        Label20.Text = MainFrm.BlockGroupLBL.Text
        PopulateTeacherCombo()
        PopulateBlockGroupCB()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        MsgBox("Please note that alterations, can take up to 14 days. As these are made on Student ONE database which can take time for it to update.")

        Dim OutApp As Object
        Dim OutMail As Object
        Dim body As String
        Dim AdminEmail As String
        Dim ApptrainEmail As String
        Dim FrankOffer As String
        Dim Trades As String
        Dim imageData As Byte() = RetrieveImageDataFromDatabase()

        ' Retrieve email addresses from the EmailSettings table
        GetEmailAddresses(AdminEmail, ApptrainEmail, FrankOffer, Trades)

        ' Create a new instance of Outlook Application
        Dim outlookApp As New Outlook.Application()

        body = "Dear Administrator,<br><br>
I hope this message finds you well. I wanted to bring to your attention an issue regarding the employer details of a student.<br><br>
Student ID: " & Label3.Text & "<br>
Student Name: " & Label4.Text & " " & Label5.Text & "<br>
Student Email: " & Label6.Text & "<br>
Student Contact Phone Number: " & MainFrm.Label29.Text & "<br><br>
It has come to our notice that there may be discrepancies in the employer information of the above-mentioned student.<br>
As part of our commitment to accuracy and compliance, it's crucial for us to ensure that all student records are up-to-date.<br><br>
If the student is currently employed, we kindly request your assistance in verifying and updating their employer details in the Student ONE database.<br>
If, however, the student is not currently employed, we kindly request the removal of any outdated employer information from their Student ONE profile.<br><br>
Your attention to this matter would be greatly appreciated.<br><br>
Thank you for your cooperation.<br><br>
Submitted by: " & SubmittedbyCB.Text & "<br><br>
This message is auto-generated for your convenience."

        OutApp = CreateObject("Outlook.Application")
        ' Create a new email item
        OutMail = OutApp.CreateItem(0)

        ' Set email properties
        With OutMail

            .To = ApptrainEmail
            .cc = Trades
            .bcc = ""
            .Subject = "Student Amendment Request - Incorrect Employer Information"
            .HTMLbody = body & "<br><img src='data:image/jpeg;base64," & Convert.ToBase64String(imageData) & "' width='90%'> " & MainFrm.VersionLBL.Text
            .Display ' Display the email

            ' Display the email
            .Display(True)

            ' Release COM objects
            ReleaseComObject(outlookApp)
        End With

        ' Clean up
        OutMail = Nothing
        OutApp = Nothing
        Me.Close()
        MsgBox("Your Email has been Generated!")
    End Sub
    Private Sub PopulateBlockGroupCB()
        ' Replace "Your_Connection_String_Here" with your actual connection string
        Dim connectionString As String = SQLCon.connectionString

        ' Create and open a SqlConnection
        Using connection As New SqlConnection(connectionString)
            connection.Open()

            Dim commandText As String = "SELECT DISTINCT [Block Group Code] FROM ElectrotechnologyReports.dbo.AgreementsDetails"
            Using command As New SqlCommand(commandText, connection)
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        If Not reader.IsDBNull(0) Then
                            ' Get the block group code from the reader
                            Dim blockGroupCode As String = reader.GetString(0)

                            ' Remove non-numeric characters from the blockGroupCode
                            'blockGroupCode = New String(blockGroupCode.Where(Function(c) Char.IsDigit(c)).ToArray())

                            ' Add the modified blockGroupCode to the ComboBox
                            ComboBox1.Items.Add(blockGroupCode)
                        End If
                    End While
                End Using
            End Using
        End Using
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Label24.Visible = True
            ComboBox3.Visible = True
            Label19.Visible = True
            Label20.Visible = True
            ComboBox4.Visible = True
            Label26.Visible = True
            ComboBox2.Visible = True
            Label25.Visible = True
            Label21.Visible = True
            ComboBox1.Visible = True
            Button3.Visible = True
            Label28.Visible = False
            Label22.Visible = True
            DateTimePicker1.Visible = True
            CheckBox1.Checked = False
            CheckBox3.Checked = False
        Else
            Label24.Visible = False
            ComboBox3.Visible = False
            Label19.Visible = False
            Label20.Visible = False
            ComboBox4.Visible = False
            Label26.Visible = False
            ComboBox1.Visible = False
            Label21.Visible = False
            ComboBox2.Visible = False
            Label25.Visible = False
            Label21.Visible = False
            ComboBox1.Visible = False
            Button3.Visible = False
            Label28.Visible = False
            Label22.Visible = False
            DateTimePicker1.Visible = False
            CheckBox1.Checked = False
            CheckBox2.Checked = False
            CheckBox3.Checked = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            Label24.Visible = False
            ComboBox3.Visible = False
            Label19.Visible = False
            Label20.Visible = False
            ComboBox4.Visible = False
            Label26.Visible = False
            ComboBox1.Visible = True
            Label21.Visible = True
            Button3.Visible = True
            Label28.Visible = True
            Label22.Visible = False
            DateTimePicker1.Visible = True
            CheckBox1.Checked = False
            CheckBox2.Checked = False
        Else
            Label24.Visible = False
            ComboBox3.Visible = False
            Label19.Visible = False
            Label20.Visible = False
            ComboBox4.Visible = False
            Label26.Visible = False
            ComboBox1.Visible = False
            Label21.Visible = False
            ComboBox2.Visible = False
            Label25.Visible = False
            Label21.Visible = False
            ComboBox1.Visible = False
            Button3.Visible = False
            Label28.Visible = False
            Label22.Visible = False
            DateTimePicker1.Visible = False
            CheckBox1.Checked = False
            CheckBox2.Checked = False
            CheckBox3.Checked = False
        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Label24.Visible = False
            ComboBox3.Visible = False
            Label19.Visible = True
            Label20.Visible = True
            ComboBox4.Visible = False
            Label26.Visible = False
            Button3.Visible = True
            Label28.Visible = True
            Label22.Visible = False
            DateTimePicker1.Visible = True
            CheckBox2.Checked = False
            CheckBox3.Checked = False
        Else
            Label24.Visible = False
            ComboBox3.Visible = False
            Label19.Visible = False
            Label20.Visible = False
            ComboBox4.Visible = False
            Label26.Visible = False
            ComboBox1.Visible = False
            Label21.Visible = False
            ComboBox2.Visible = False
            Label25.Visible = False
            Label21.Visible = False
            ComboBox1.Visible = False
            Button3.Visible = False
            Label28.Visible = False
            Label22.Visible = False
            DateTimePicker1.Visible = False
            CheckBox1.Checked = False
            CheckBox2.Checked = False
            CheckBox3.Checked = False
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        LookupTeacherEmail()
    End Sub

    Private Sub SubmittedbyCB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SubmittedbyCB.SelectedIndexChanged

    End Sub
End Class