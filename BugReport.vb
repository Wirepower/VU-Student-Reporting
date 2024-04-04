Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports Microsoft.Office.Interop
Imports System.Net.Mail
Imports System.Runtime.InteropServices
Imports Microsoft.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement


Public Class BugReport
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim missingFields As String = ""
        If IssueTB.Text = "" Then
            missingFields &= "- What is the Issue" & vbCrLf
        End If

        If HowTB.Text = "" Then
            missingFields &= "- How can this issue be replicated" & vbCrLf
        End If

        If ReportByTB.Text = "" Then
            missingFields &= "- Who is making this report" & vbCrLf
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

        body = "Dear Team,<BR><BR>
I am writing to bring to your attention a potential issue encountered in the application 'Student Attendance Reporting', version " & MainFrm.VersionLBL.Text & ".<BR><BR>
Issue:<BR>
" & IssueTB.Text & "<BR><BR>
Steps to Replicate:<BR>
" & HowTB.Text & "<BR><BR>
Additional Information:<BR>
" & OtherTB.Text & "<BR><BR>
This issue was reported by:<BR>
" & ReportByTB.Text & "<BR>"

        OutApp = CreateObject("Outlook.Application")
        ' Create a new email item
        OutMail = OutApp.CreateItem(0)

        ' Set email properties
        With OutMail
            .To = FrankOffer
            .cc = ""
            .bcc = ""
            .Subject = "BUG REPORT: STUDENT ATTENDANCE REPORTING " & MainFrm.VersionLBL.Text & "."
            .HTMLbody = body & "<br><img src='data:image/jpeg;base64," & Convert.ToBase64String(imageData) & "' width='90%'> " & MainFrm.VersionLBL.Text
            .Display ' Display the email
        End With

        ' Release COM objects
        Marshal.ReleaseComObject(OutMail)
        Marshal.ReleaseComObject(OutApp)

        ' Clean up
        OutMail = Nothing
        OutApp = Nothing

        Me.Close()
        MsgBox("Your Email has been Generated!")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
    End Sub

    Private Sub BugReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PopulateTeacherCombo()
    End Sub
    Private Sub PopulateTeacherCombo()
        ' Replace "Your_Connection_String_Here" with your actual connection string
        Dim connectionString As String = SQLCon.connectionString

        ' Create and open a SqlConnection
        Using connection As New SqlConnection(connectionString)
            connection.Open()

            ' Clear existing items in the ComboBox
            ReportByTB.Items.Clear()

            ' SQL query to retrieve all teachers' names
            Dim query As String = "SELECT Teacher_Full_Name FROM ElectrotechnologyReports.dbo.TeacherList WHERE Highest_Certificate_Taught = 'Certificate III' ORDER BY Teacher_Full_Name ASC"

            ' Create a SqlCommand object with the query and connection
            Using command As New SqlCommand(query, connection)
                ' Execute the query and retrieve the data
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        ' Add each teacher's name to the ComboBox
                        ReportByTB.Items.Add(reader.GetString(0))
                    End While
                End Using
            End Using
        End Using
    End Sub
End Class