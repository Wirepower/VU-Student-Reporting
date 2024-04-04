Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports Microsoft.Office.Interop
Imports System.Net.Mail
Imports Microsoft.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports Microsoft.Office.Server.UserProfiles

Public Class StudentAmendment
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

            ' SQL query to retrieve all teachers' names
            Dim query As String = "SELECT Teacher_Full_Name FROM ElectrotechnologyReports.dbo.TeacherList WHERE Highest_Certificate_Taught = 'Certificate III' ORDER BY Teacher_Full_Name ASC"

            ' Create a SqlCommand object with the query and connection
            Using command As New SqlCommand(query, connection)
                ' Execute the query and retrieve the data
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        ' Add each teacher's name to the ComboBox
                        SubmittedbyCB.Items.Add(reader.GetString(0))
                    End While
                End Using
            End Using
        End Using
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim missingFields As String = ""
        Dim action As String = ""
        If (ComboBox1.SelectedIndex = -1 AndAlso Not CheckBox1.Checked) OrElse
    (ComboBox1.SelectedIndex <> -1 AndAlso CheckBox1.Checked) Then
            ' Either a ComboBox selection is made or a Checkbox is checked, but not both
            ' Add the missing field message to the string
            missingFields &= "- Proposed Blockgroup Change or Delete the student Checkbox." & vbCrLf & "It cant be both!" & vbCrLf
        End If

        If CheckBox1.Checked Then
            ' Set the action variable to "DELETE"
            action = "deleted"
        Else
            action = "changed"
        End If

        If DateTimePicker1.Text = "" Then
            missingFields &= "- Date" & vbCrLf
        End If

        If SubmittedbyCB.Text = "" Then
            missingFields &= "- Submitted By" & vbCrLf
        End If
        If missingFields <> "" Then
            MessageBox.Show("Please fill the following fields:" & vbCrLf & missingFields)
            Exit Sub
        End If

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

        body = "Hello, <BR>
I would like to get the following student blockgroup " & action & ". <BR><BR><BR>
Student ID: " & Label3.Text & "<BR><BR>
Student Name: " & Label4.Text & " " & Label5.Text & " <BR><BR>
Student Email: " & Label6.Text & " <BR><BR>
From BlockGroup: " & Label20.Text & " to Blockgroup: " & ComboBox1.Text & "<BR><BR>
Starting from: " & DateTimePicker1.Text & "<BR><BR>
Submitted by: " & SubmittedbyCB.Text & "<BR><BR>"



        OutApp = CreateObject("Outlook.Application")
        ' Create a new email item
        OutMail = OutApp.CreateItem(0)

        ' Set email properties
        With OutMail

            .To = FrankOffer
            .cc = AdminEmail
            .bcc = ""
            .Subject = "Student Amendment Request - ADD/MOVE/DELETE Student from Blockgroup"
            .HTMLbody = Body & "<br><img src='data:image/jpeg;base64," & Convert.ToBase64String(imageData) & "' width='90%'> " & MainFrm.VersionLBL.Text
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

            .To = Trades
            .cc = ""
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
                            blockGroupCode = New String(blockGroupCode.Where(Function(c) Char.IsDigit(c)).ToArray())

                            ' Add the modified blockGroupCode to the ComboBox
                            ComboBox1.Items.Add(blockGroupCode)
                        End If
                    End While
                End Using
            End Using
        End Using
    End Sub

End Class