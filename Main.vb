Imports System.ComponentModel
Imports System.IO
Imports Microsoft.Data.SqlClient
Imports Microsoft.Win32
Imports System.Net
'Imports System.Data.SqlClient
Imports Student_Attendance_Reporting
Imports System.Timers
Imports System.Diagnostics.Eventing
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Outlook
Imports System.Text
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab


Public Class MainFrm
    ''' <summary>Display version derived from assembly version (set in .vbproj). Keeps UI in sync with update check.</summary>
    Private ReadOnly Property Version As String
        Get
            Dim v As System.Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            If v Is Nothing Then Return "V0.0"
            Return "V" & v.Major.ToString() & "." & v.Minor.ToString()
        End Get
    End Property

    Private connection As SqlConnection
    Private WithEvents connectionCheckTimer As New Timer()
    Public studentID As String
    Public studentFirstname As String
    Public studentSurname As String
    Public studentEmail As String
    Public employerFirstname As String
    Public employerSurname As String
    Public employerBusinessName As String
    Public employerEmail As String
    Private template As String = "" ' Declaration at the class level
    Private _profilingApiToolTip As ToolTip

    ' Declare teacher list workbook and worksheet
    Private Sub UpdateReconnectButtonVisibility()
        ' Show reconnect button if either side is "down":
        ' - SQL: connection is closed
        ' - API: profiling API is not configured (no usable bearer token)
        Dim sqlClosed As Boolean = (connection Is Nothing) OrElse (connection.State = ConnectionState.Closed)
        Dim apiClosed As Boolean = Not ExemplarProfilingApi.IsConfigured()
        btnReconnect.Visible = sqlClosed OrElse apiClosed
    End Sub

    ''' <summary>
    ''' Reconnect SQL and also force the Exemplar API to re-initialize (refresh token).
    ''' </summary>
    Private Async Sub btnReconnect_Click(sender As Object, e As EventArgs) Handles btnReconnect.Click
        Try
            Dim sqlClosed As Boolean = (connection Is Nothing) OrElse (connection.State = ConnectionState.Closed)
            Dim apiClosed As Boolean = Not ExemplarProfilingApi.IsConfigured()

            ' Re-open only the closed connection(s).
            If sqlClosed Then
                SQLCon.OpenConnection(connection)
            End If
            UpdateReconnectButtonVisibility()

            If apiClosed Then
                ' Force a fresh bearer token next time we call the API.
                ExemplarProfilingApi.ClearCachedToken()
            End If

            If ExemplarProfilingApi.IsConfigured() Then
                SetProfilingApiStatus("Ready", "", Color.DarkGreen)
            Else
                SetProfilingApiStatus("Not configured", ExemplarProfilingApi.GetNotConfiguredReason(), Color.DarkOrange)
            End If

            ' If a student is already selected, refresh profiling so the UI reflects the reconnect.
            If StudentCB.SelectedIndex >= 0 AndAlso StudentCB.SelectedItem IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(StudentIDLBL.Text) Then
                Await RefreshSelectedStudentProfilingAsync()
                UpdateExemplarProfilingEmailButtonVisibility()
            End If
        Catch ex As System.Exception
            SetProfilingApiStatus("Error reconnecting", ex.Message, Color.Maroon)
        End Try
    End Sub
    Private Function IsOutlookInstalled() As Boolean
        Dim registryCheckResult As Boolean = RegistryCheck()
        Dim fileCheckResult As Boolean = FileCheck()
        Return registryCheckResult Or fileCheckResult
    End Function

    Private Function RegistryCheck() As Boolean
        Dim outlookKey As RegistryKey = Registry.ClassesRoot.OpenSubKey("Outlook.Application")
        If outlookKey IsNot Nothing Then
            outlookKey.Close()
            Return True
        Else
            Return False
        End If
    End Function

    Private Function FileCheck() As Boolean
        Dim outlookExePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft Office", "root", "OfficeXX", "OUTLOOK.EXE")
        Dim outlookExePathAlt As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft Office", "root", "OfficeXX", "OUTLOOK.EXE")

        Return File.Exists(outlookExePath) OrElse File.Exists(outlookExePathAlt)
    End Function
    Private Function IsAnyConnectRunning() As Boolean
        ' Check if the Cisco AnyConnect process is running
        Dim processes() As Process = Process.GetProcessesByName("vpnui")

        If processes.Length > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Async Sub MainFrm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim loadingForm As New LoadingForm()
        loadingForm.Show()
        ' Define custom increments
        Dim totalSteps As Integer = 100 ' Total number of steps
        Dim currentStep As Integer = 5  ' Current step
        ' Initialize the status label
        SQLCon.InitializeStatusLabel(StatusLbl)
        ' Get the SQL connection
        connection = SQLCon.GetConnection()
        ' Open the connection
        SQLCon.OpenConnection(connection)
        If connection.State = ConnectionState.Open Then
            If Not IsOutlookInstalled() Then
                MessageBox.Show("Microsoft Outlook is not detected on this system. This application requires Outlook to be installed in order to function properly. Please install Microsoft Outlook and try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
            End If

            ' Check if the Cisco AnyConnect process is running
            If IsAnyConnectRunning() Then
                ' AnyConnect is running, prompt the user to connect
                Dim result As DialogResult = MessageBox.Show("For security reasons, this application requires a VPN connection. Please ensure you are connected to Cisco AnyConnect VPN before using the application. Do you want to proceed anyway?", "VPN Connection Required", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

                If result = DialogResult.No Then
                    'MsgBox("User Selected No")
                    ' User chose "No", close the application
                    Me.Close()
                End If
            Else
                ' AnyConnect is not running, prompt the user to connect
                Dim result As DialogResult = MessageBox.Show("For security reasons, this application requires a VPN connection. Please ensure you are connected to Cisco AnyConnect VPN before using the application. Do you want to proceed anyway?", "VPN Connection Required", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

                If result = DialogResult.No Then
                    'MsgBox("User Selected No")
                    ' User chose "No", close the application
                    Me.Close()
                End If
            End If

            Dim directoryPath As String = "P:\VUPoly\MT&T\IT, Electrical and Engineering\Student Reporting Database\"

            If Not Directory.Exists(directoryPath) Then
                MessageBox.Show("Notification: Access to P-Drive Unavailable" & vbCrLf & vbCrLf & "It appears that access to the P-Drive directory is currently unavailable. We recommend reaching out to your IT department for further assistance." & vbCrLf & vbCrLf & "Please note that while this program will continue to operate without access to the P-Drive directory, it is essential to recognize that you will lose the capability to receive software updates.", "P-Drive Access Unavailable", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

            ' Set visibility of reconnect button (SQL closed OR API not configured)
            ConfigureResponsiveLayout()
            ' Configure the timer control
            connectionCheckTimer.Interval = 1000 ' 1 second interval
            connectionCheckTimer.Start()
            UpdateReconnectButtonVisibility()
            ' Show the loading form

            ' Populate the ComboBox with unique values from the "Block Group Code" column
            PopulateWeekdays()
            PopulateBlockGroupCB()
            PopulateTeacherComboBox()
            Populateunit()
            PopulateEmailSubjectComboBox()
            ResetApptrainData()
            UpdateStudentDatabaseLabel()
            If ExemplarProfilingApi.IsConfigured() Then
                SetProfilingApiStatus("Ready", "", Color.DarkGreen)
            Else
                SetProfilingApiStatus("Not configured", ExemplarProfilingApi.GetNotConfiguredReason(), Color.DarkOrange)
            End If
            'Put Code here - Load Form/application

            currentStep = 25
            ' Update progress bar to reflect current progress
            loadingForm.UpdateProgress(currentStep)

            'Put Code here - Load Agreements database

            currentStep = 50
            ' Update progress bar to reflect current progress
            loadingForm.UpdateProgress(currentStep)

            'Put Code here - Load studentlog Database

            currentStep = 75
            ' Update progress bar to reflect current progress
            loadingForm.UpdateProgress(currentStep)
            System.Windows.Forms.Application.DoEvents()
            loadingForm.Label1.Text = "Intializing Databases.. Please Wait"
            ' Initialize Excel objects



            'Put Code here - Load Unit Database


            Button8.Visible = False
            Button7.Visible = False
            Button3.Visible = False
            System.Windows.Forms.Application.DoEvents()
            currentStep = 90
            ' Update progress bar to reflect current progress
            loadingForm.UpdateProgress(currentStep)

            'Put Code here - Load Teacher Database

            ComboBox4.Items.Add("Satisfactory")
            ComboBox4.Items.Add("Not Satisfactory")

            ComboBox5.Items.Add("Satisfactory")
            ComboBox5.Items.Add("Not Satisfactory")

            ComboBox6.Items.Add("Satisfactory")
            ComboBox6.Items.Add("Not Satisfactory")

            ComboBox4.Text = ""
            ComboBox5.Text = ""
            ComboBox6.Text = ""
            VersionLBL.Text = Version
            '-------
            SettingsForm.MassEmailChkBx.Checked = My.Settings.MassEmail
            ' Retrieve the stored state of the checkbox from application settings
            SettingsForm.MassEmailChkBx.Checked = My.Settings.MassEmail

            ' Get the initial visibility of the MassEmailBtn button based on the checkbox state
            If SettingsForm.MassEmailChkBx.Checked Then
                MassEmailBtn.Visible = True
            Else
                MassEmailBtn.Visible = False
            End If
            '-------

            System.Windows.Forms.Application.DoEvents()
            loadingForm.Label1.Text = "Loading Complete!"
            loadingForm.UpdateProgress(totalSteps)
            ' Simulate a delay
            System.Threading.Thread.Sleep(1000)
            'CheckVersionAndDisplayInfo()
            ' Close the loading form once loading is finished
            loadingForm.Close()

            ' Stage 2 OTA: check for mandatory updates on startup.
            ' Await so the UI thread can continue painting/showing the main window.
            Await CheckForUpdatesAsync(showNoUpdateMessage:=False)
        Else
            Me.Hide()
            SQLError.Show()
        End If
    End Sub

    'Private Sub CheckConnections()
    '   CheckNetworkDriveConnection()
    '   CheckSqlServerConnection()
    'End Sub
    Private Sub DownloadAndUpdate()
        Try
            ' Download the update file
            Dim updateFileUrl As String = "P:\VUPoly\MT&T\IT, Electrical and Engineering\Student Reporting Database\StudentAttendanceReporting.exe"
            Dim client As New WebClient()
            client.DownloadFile(updateFileUrl, "StudentAttendanceReporting.exe")

            ' Display a confirmation message when the update is successfully downloaded and installed
            MessageBox.Show("Update downloaded and installed successfully.", "Update Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' Execute the downloaded update file
            Process.Start("StudentAttendanceReporting.exe")
        Catch ex As System.Exception
            ' Handle any errors that occur during the download or execution
            MessageBox.Show("An error occurred while downloading or executing the update.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>Runs the startup update check without blocking the form load (avoids deadlock).</summary>
    Private Async Sub RunStartupUpdateCheckAsync()
        Await CheckForUpdatesAsync(showNoUpdateMessage:=False)
    End Sub

    Private Async Function CheckForUpdatesAsync(showNoUpdateMessage As Boolean) As Task
        Try
            Dim gitHubResult As GitHubUpdateCheckResult = Await UpdateModule.CheckForGitHubUpdateAsync()
            If gitHubResult.IsSuccessful Then
                If gitHubResult.IsMandatory Then
                    MessageBox.Show(
                        "A mandatory application update is required." & vbCrLf &
                        "Please update now to continue using the application." & vbCrLf & vbCrLf &
                        "The updater will start now.",
                        "Mandatory Update Required",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    )

                    Dim mandatoryInstall As GitHubUpdateInstallResult = Await UpdateModule.DownloadAndLaunchGitHubUpdateAsync(gitHubResult)
                    If mandatoryInstall.IsSuccessful Then
                        MessageBox.Show("Update launched successfully. The application will now close so the update can complete.", "Update In Progress", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show(
                            "Unable to launch the mandatory updater." & vbCrLf &
                            mandatoryInstall.ErrorMessage & vbCrLf & vbCrLf &
                            "The application will now close to prevent running an out-of-date version.",
                            "Mandatory Update Failed",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        )
                    End If

                    System.Windows.Forms.Application.Exit()
                    Return
                End If

                If gitHubResult.IsUpdateAvailable Then
                    Dim promptResult As DialogResult = MessageBox.Show(
                        "A newer version is available." & vbCrLf &
                        "Do you want to download and launch this update now?",
                        "Update Available",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information
                    )

                    If promptResult = DialogResult.Yes Then
                        Dim installResult As GitHubUpdateInstallResult = Await UpdateModule.DownloadAndLaunchGitHubUpdateAsync(gitHubResult)
                        If installResult.IsSuccessful Then
                            MessageBox.Show("Update launched successfully. The application will close so the update can complete.", "Update In Progress", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            System.Windows.Forms.Application.Exit()
                        Else
                            MessageBox.Show("The update could not be launched." & vbCrLf & installResult.ErrorMessage, "Update Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    End If

                    Return
                End If

                If showNoUpdateMessage Then
                    MessageBox.Show(
                        "Your application is up to date.",
                        "No Update Available",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    )
                End If

                Return
            End If
        Catch
            ' Continue to legacy updater fallback below.
        End Try

        ' Legacy SQL/P-drive update path fallback remains available.
        If UpdateModule.IsUpdateAvailable() Then
            Dim result As DialogResult = MessageBox.Show("An update is available. Do you want to download and install it?", "Update Available", MessageBoxButtons.YesNo)
            If result = DialogResult.Yes Then
                DownloadAndUpdate()
            End If
        ElseIf showNoUpdateMessage Then
            MessageBox.Show("Your application is up to date.", "No Update Available", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Function

    ''' <summary>DPI-aware responsive layout: anchors so the form adapts to all screen resolutions and scales correctly on high-DPI.</summary>
    Private Sub ConfigureResponsiveLayout()
        Me.MinimumSize = New Size(1280, 900)
        ' Ensure form can be resized by user; AutoScroll allows scrolling when content is taller than window
        Me.MaximumSize = New Size(0, 0)

        ' ---- Header: title and version centered when form widens ----
        Label1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        Label1.TextAlign = ContentAlignment.TopCenter
        VersionLBL.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        VersionLBL.TextAlign = ContentAlignment.TopCenter

        ' ---- Top-right action buttons (stay at right edge) ----
        Button1.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Button2.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Button9.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        MassEmailBtn.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Button3.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Button10.Anchor = AnchorStyles.Top Or AnchorStyles.Right

        ' ---- Top-left: logo and DB date ----
        Label36.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        Label37.Anchor = AnchorStyles.Top Or AnchorStyles.Right

        ' ---- Block/class row: stretch across ----
        Label2.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        BlockGroupCB.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        Label27.Anchor = AnchorStyles.Top Or AnchorStyles.Left

        ' ---- Search row: label left, textbox and button stay at right ----
        Label26.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        StudentIDTextBox.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Button6.Anchor = AnchorStyles.Top Or AnchorStyles.Right

        ' ---- Student selection and selected student area ----
        Label4.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        StudentCB.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        GroupBox1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        SelectedStudentLBL.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        ' ---- Email/subject and teacher row: stretch ----
        ComboBox12.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        teacherNameComboBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        DateTimePicker.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        ' ---- Alerts and messages: full width ----
        UnitAlertLbl.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        InvestigationLBL.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        resitLabel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        ' ---- Notes area ----
        Label32.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        Button11.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        NotesTB.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        ' ---- Full-width action buttons (Submit / Student Investigation) ----
        Button7.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        Button8.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        ' ---- Separator and status ----
        Label14.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        StatusLbl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        btnReconnect.Anchor = AnchorStyles.Top Or AnchorStyles.Left

        ' ---- Profiling API section (left column, labels can wrap) ----
        Label38.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        Label39.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        ProfilingMissingLbl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        ProfilingMissingValLbl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        ProfilingNotVerifiedLbl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        ProfilingNotVerifiedValLbl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        ProfilingEmployerVerifiedLbl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        ProfilingEmployerVerifiedValLbl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        ProfilingLastCardLbl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        ProfilingLastCardValLbl.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
    End Sub

    Private Sub SetProfilingApiStatus(statusText As String, detailText As String, Optional statusColor As Color? = Nothing)
        Label38.Text = "Profiling API:"
        Label38.Visible = True
        Label39.Visible = True
        Label39.Text = If(String.IsNullOrWhiteSpace(detailText), statusText, $"{statusText} | {detailText}")
        Label39.ForeColor = If(statusColor.HasValue, statusColor.Value, Color.Black)
        If _profilingApiToolTip Is Nothing Then _profilingApiToolTip = New ToolTip()
        If Not String.IsNullOrWhiteSpace(detailText) Then
            _profilingApiToolTip.SetToolTip(Label39, detailText)
        Else
            _profilingApiToolTip.SetToolTip(Label39, "")
        End If
        ProfilingMissingLbl.Visible = False
        ProfilingMissingValLbl.Visible = False
        ProfilingNotVerifiedLbl.Visible = False
        ProfilingNotVerifiedValLbl.Visible = False
        ProfilingEmployerVerifiedLbl.Visible = False
        ProfilingEmployerVerifiedValLbl.Visible = False
        ProfilingLastCardLbl.Visible = False
        ProfilingLastCardValLbl.Visible = False
    End Sub

    ''' <summary>Shows Connected in its own label and the card stats in separate colored labels (move them in the Designer).</summary>
    Private Sub SetProfilingApiStatusDetailed(r As ExemplarProfileLookupResult)
        Label38.Text = "Profiling API:"
        Label38.Visible = True
        Label39.Visible = True
        Label39.Text = r.StatusText
        Label39.ForeColor = Color.DarkGreen
        If _profilingApiToolTip IsNot Nothing Then _profilingApiToolTip.SetToolTip(Label39, "")
        ProfilingMissingLbl.Text = "Cards not submitted/Outstanding:"
        ProfilingMissingLbl.ForeColor = Color.Red
        ProfilingMissingValLbl.Text = If(r.MissingWeeks.HasValue, r.MissingWeeks.Value.ToString(), "?")
        ProfilingMissingValLbl.ForeColor = Color.Black
        ProfilingNotVerifiedLbl.Text = "Cards Submitted (Not verified):"
        ProfilingNotVerifiedLbl.ForeColor = Color.Orange
        ProfilingNotVerifiedValLbl.Text = If(r.SubmittedNotVerified.HasValue, r.SubmittedNotVerified.Value.ToString(), "?")
        ProfilingNotVerifiedValLbl.ForeColor = Color.Black
        ProfilingEmployerVerifiedLbl.Text = "Cards submitted (Employer Verified):"
        ProfilingEmployerVerifiedLbl.ForeColor = Color.DarkGreen
        ProfilingEmployerVerifiedValLbl.Text = If(r.SubmittedEmployerVerified.HasValue, r.SubmittedEmployerVerified.Value.ToString(), "?")
        ProfilingEmployerVerifiedValLbl.ForeColor = Color.Black
        ProfilingLastCardLbl.Text = "Last Card Submission:"
        ProfilingLastCardLbl.ForeColor = Color.Blue
        ProfilingLastCardValLbl.Text = If(String.IsNullOrEmpty(r.LastCardFormatted), "?", r.LastCardFormatted)
        ProfilingLastCardValLbl.ForeColor = Color.Black
        ProfilingMissingLbl.Visible = True
        ProfilingMissingValLbl.Visible = True
        ProfilingNotVerifiedLbl.Visible = True
        ProfilingNotVerifiedValLbl.Visible = True
        ProfilingEmployerVerifiedLbl.Visible = True
        ProfilingEmployerVerifiedValLbl.Visible = True
        ProfilingLastCardLbl.Visible = True
        ProfilingLastCardValLbl.Visible = True
    End Sub

    Private Sub UpdateExemplarProfilingEmailButtonVisibility()
        Button11.Visible = StudentCB.SelectedIndex >= 0 AndAlso StudentCB.SelectedItem IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(StudentIDLBL.Text)
    End Sub

    Private Async Function RefreshSelectedStudentProfilingAsync() As Task
        Dim studentId As String = StudentIDLBL.Text?.Trim()
        Dim firstName As String = StudentFirstnameLBL.Text?.Trim()
        Dim lastName As String = StudentSurnameLBL.Text?.Trim()
        ' Agreement / student record email (Exemplar lookup default).
        Dim email As String = StudentEmailLBL.Text?.Trim()
        ' Prefer dbo.ExemplarProfilingStudentDB when present; otherwise use agreement email above.
        Dim overrideEmail As String = ExemplarEmailOverrides.GetOverride(studentId)
        Dim lookupEmail As String = If(String.IsNullOrWhiteSpace(overrideEmail), email, overrideEmail)

        If String.IsNullOrWhiteSpace(firstName) OrElse String.IsNullOrWhiteSpace(lastName) Then
            SetProfilingApiStatus("Waiting", "Select a student to query the profiling API.", Color.Black)
            Return
        End If

        SetProfilingApiStatus("Checking", $"{firstName} {lastName}", Color.SteelBlue)
        Dim lookupResult As ExemplarProfileLookupResult = Await ExemplarProfilingApi.LookupStudentProfileAsync(firstName, lastName, lookupEmail)

        If Not lookupResult.IsConfigured Then
            SetProfilingApiStatus("Not configured", ExemplarProfilingApi.GetNotConfiguredReason(), Color.DarkOrange)
            UpdateReconnectButtonVisibility()
            Return
        End If

        If lookupResult.IsSuccessful Then
            SetProfilingApiStatusDetailed(lookupResult)
        ElseIf lookupResult.StatusText = "Student not found" Then
            Dim profilingEmail As String = InputBox(
                "No matching Exemplar student was found for " & firstName & " " & lastName & "." & vbCrLf & vbCrLf &
                "Enter the student's Exemplar profiling email address to try again (or leave blank to skip):",
                "Profiling email",
                lookupEmail
            )
            If Not String.IsNullOrWhiteSpace(profilingEmail) Then
                ExemplarEmailOverrides.SetOverride(studentId, profilingEmail.Trim())
                SetProfilingApiStatus("Checking", "Retrying with profiling email...", Color.SteelBlue)
                lookupResult = Await ExemplarProfilingApi.LookupStudentProfileAsync(firstName, lastName, profilingEmail.Trim())
                If lookupResult.IsSuccessful Then
                    SetProfilingApiStatusDetailed(lookupResult)
                Else
                    SetProfilingApiStatus(lookupResult.StatusText, lookupResult.DetailText, Color.Maroon)
                End If
            Else
                SetProfilingApiStatus(lookupResult.StatusText, lookupResult.DetailText, Color.Maroon)
            End If
        Else
            SetProfilingApiStatus(lookupResult.StatusText, lookupResult.DetailText, Color.Maroon)
        End If

        ' Token may have been cleared (401 -> ClearCachedToken) during the lookup attempt.
        ' Refresh button visibility so reconnect appears/disappears immediately.
        UpdateReconnectButtonVisibility()
    End Function

    ''' <summary>Shows email editor. Returns Nothing if cancelled, otherwise trimmed text (may be empty to clear SQL override).</summary>
    Private Function ShowExemplarProfilingEmailSaveDialog(initialEmail As String, infoMessage As String) As String
        Using dlg As New Form()
            dlg.Text = "Exemplar profiling email"
            dlg.FormBorderStyle = FormBorderStyle.FixedDialog
            dlg.StartPosition = FormStartPosition.CenterParent
            dlg.MinimizeBox = False
            dlg.MaximizeBox = False
            dlg.ShowInTaskbar = False
            ' Give the buttons enough room under higher DPI/font scaling.
            ' Extra height so scaled fonts don't clip the buttons.
            dlg.ClientSize = New Size(540, 190)
            dlg.Font = Me.Font

            Dim infoLbl As New Label() With {
                .Left = 12,
                .Top = 12,
                .Width = 516,
                .Height = 56,
                .AutoSize = False,
                .Text = infoMessage
            }

            Dim emailTb As New TextBox() With {
                .Left = 12,
                .Top = 74,
                .Width = 516,
                .Text = If(initialEmail, "")
            }

            Dim saveBtn As New Button() With {
                .Text = "Save",
                .Left = 290,
                .Top = 112,
                .Width = 96,
                .Height = 32,
                .DialogResult = DialogResult.OK
            }
            Dim cancelBtn As New Button() With {
                .Text = "Cancel",
                .Left = 392,
                .Top = 112,
                .Width = 96,
                .Height = 32,
                .DialogResult = DialogResult.Cancel
            }

            dlg.Controls.Add(infoLbl)
            dlg.Controls.Add(emailTb)
            dlg.Controls.Add(saveBtn)
            dlg.Controls.Add(cancelBtn)
            dlg.AcceptButton = saveBtn
            dlg.CancelButton = cancelBtn

            If dlg.ShowDialog(Me) <> DialogResult.OK Then
                Return Nothing
            End If

            Return emailTb.Text.Trim()
        End Using
    End Function

    Private Async Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim studentId As String = StudentIDLBL.Text?.Trim()
        If String.IsNullOrWhiteSpace(studentId) Then
            MessageBox.Show("Select a student first.", "Exemplar profiling email", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim firstName As String = StudentFirstnameLBL.Text?.Trim()
        Dim lastName As String = StudentSurnameLBL.Text?.Trim()
        Dim agreementEmail As String = StudentEmailLBL.Text?.Trim()

        If String.IsNullOrWhiteSpace(firstName) OrElse String.IsNullOrWhiteSpace(lastName) Then
            MessageBox.Show("Student name is not loaded. Select a student from the list.", "Exemplar profiling email", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If Not ExemplarProfilingApi.IsConfigured() Then
            MessageBox.Show("Profiling API is not configured for this installation.", "Exemplar profiling email", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            UpdateReconnectButtonVisibility()
            Return
        End If

        ExemplarEmailOverrides.InvalidateCacheForStudent(studentId)
        Dim sqlOverride As String = ExemplarEmailOverrides.GetOverride(studentId)
        Dim lookupEmail As String = If(String.IsNullOrWhiteSpace(sqlOverride), agreementEmail, sqlOverride)

        Button11.Enabled = False
        Cursor = Cursors.WaitCursor
        Try
            Dim apiResult As ExemplarProfileLookupResult = Await ExemplarProfilingApi.LookupStudentProfileAsync(firstName, lastName, lookupEmail)

            If Not apiResult.IsConfigured Then
                MessageBox.Show(apiResult.DetailText, "Exemplar profiling email", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                UpdateReconnectButtonVisibility()
                Return
            End If

            Dim initialBoxText As String = ""
            Dim infoMessage As String

            If apiResult.IsSuccessful Then
                initialBoxText = If(String.IsNullOrWhiteSpace(apiResult.MatchedUserEmail), "", apiResult.MatchedUserEmail)
                infoMessage = "Enter Exemplar Profiling email address associated to the student." & vbCrLf &
                    "A matching Exemplar account was found. The email below is from the API — you can change it and click Save to store an override in the database."
            ElseIf String.Equals(apiResult.StatusText, "Student not found", StringComparison.OrdinalIgnoreCase) Then
                initialBoxText = ""
                infoMessage = "Enter Exemplar Profiling email address associated to the student." & vbCrLf &
                    "No Exemplar account was found with the current lookup. Enter the email to save as an override, then click Save."
            Else
                MessageBox.Show(
                    "Could not verify the student on Exemplar: " & If(String.IsNullOrWhiteSpace(apiResult.ErrorMessage), apiResult.DetailText, apiResult.ErrorMessage),
                    "Exemplar profiling email",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
                Return
            End If

            Cursor = Cursors.Default
            Dim savedEmail As String = ShowExemplarProfilingEmailSaveDialog(initialBoxText, infoMessage)
            If savedEmail Is Nothing Then
                Return
            End If

            ExemplarEmailOverrides.SetOverride(studentId, savedEmail)
            Await RefreshSelectedStudentProfilingAsync()

            If String.IsNullOrWhiteSpace(savedEmail) Then
                MessageBox.Show("Override cleared. Lookups will use the agreement email unless you save a profiling email again.", "Exemplar profiling email", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Exemplar profiling email saved for this student.", "Exemplar profiling email", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Finally
            Cursor = Cursors.Default
            Button11.Enabled = True
        End Try
    End Sub

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Await CheckForUpdatesAsync(showNoUpdateMessage:=True)
    End Sub

    Private Sub CheckVersionAndDisplayInfo()
        Dim w As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim h As Integer = Screen.PrimaryScreen.Bounds.Height
        'Dim FilePath As String = "F:\Student Reporting VisualBasic\VersionCheck.txt"
        Dim FilePath As String = "P:\VUPoly\MT&T\IT, Electrical and Engineering\Student Reporting Database\VersionCheck.txt"
        Dim TestStr As String = ""

        On Error Resume Next
        TestStr = Dir(FilePath)
        On Error GoTo 0

        If TestStr = "" Then
            MsgBox("Seems like you have an old version. A new version is available. Please update to the latest version. This application will now close.")
            Me.Close() ' Close the form if an old version is detected
        Else
            MsgBox(Version & " is the current and latest version." & vbCrLf & "Your screen resolution is " & Format(w, "#,##0") & " x " & Format(h, "#,##0"), vbInformation, "Monitor Size (width x height)")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        StudentAmendment.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        EmailSubjectHelp.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Close the SQL connection if it's open
        If connection.State = ConnectionState.Open Then
            connection.Close()
        End If
        System.Windows.Forms.Application.Exit()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)
        'Application.Restart()
    End Sub
    Private Sub UpdateLastReportDate(studentID As String, reportDate As Date)
        studentID = StudentIDLBL.Text
        Dim query As String = "UPDATE ElectrotechnologyReports.dbo.StudentLogs " &
                              "SET LastStudentReportDate = @ReportDate " &
                              "WHERE [Student ID] = @StudentID"

        Using connection As New SqlConnection(SQLCon.connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@ReportDate", reportDate)
                command.Parameters.AddWithValue("@StudentID", studentID)

                Try
                    connection.Open()
                    Dim rowsAffected As Integer = command.ExecuteNonQuery()
                    If rowsAffected > 0 Then
                        'MessageBox.Show("Last student report date updated successfully.")
                    Else
                        MessageBox.Show("Student ID not found.")
                    End If
                Catch ex As System.Exception
                    MessageBox.Show("Error updating last report date: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub

    Private Function GetLastReportDate(studentID As String) As Date?
        studentID = StudentIDLBL.Text
        Dim query As String = "SELECT LastStudentReportDate " &
                              "FROM ElectrotechnologyReports.dbo.StudentLogs " &
                              "WHERE [Student ID] = @StudentID"

        Using connection As New SqlConnection(SQLCon.connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@StudentID", studentID)

                Try
                    connection.Open()
                    Dim result As Object = command.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return DirectCast(result, Date)
                    Else
                        Return Nothing
                    End If
                Catch ex As System.Exception
                    MessageBox.Show("Error retrieving last report date: " & ex.Message)
                    Return Nothing
                End Try
            End Using
        End Using
    End Function
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Dim missingFields As String = ""
        Select Case ComboBox12.SelectedItem

            Case "Student Term Progress Report"

                If GetLastReportDate(studentID) Is Nothing Then
                    LastReportDatePicker.Show()
                    Exit Sub
                End If
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

                If DateTimePicker.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

                If ComboBox4.Text = "" Then
                    missingFields &= "- Attendance/Puncuality" & vbCrLf
                End If

                If ComboBox5.Text = "" Then
                    missingFields &= "- Class Room Engagement" & vbCrLf
                End If

                If ComboBox6.Text = "" Then
                    missingFields &= "- Course Progression" & vbCrLf
                End If




            ' Additional validation logic specific to ComboBoxOption1
            Case "2 Week Intention Letter"
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

                If DateTimePicker.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

            Case "4 Week Intention Letter"
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

                If DateTimePicker.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

            Case "Course Withdraw Notice"
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

            Case "Student Behaviour Notice"
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

                If DateTimePicker.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

            Case "Overdue Fees - Warning"
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

            Case "Overdue Fees - Sanction"
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

            Case "Unit Withdraw Notice"
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

                If DateTimePicker.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

                If ComboBox7.Text = "" Then
                    missingFields &= "- Unit" & vbCrLf
                End If

            Case "Absent Notice"
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

                If DateTimePicker.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

            Case "Late Arrival Notice"
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

                If DateTimePicker.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

                If TextBox1.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

            Case "Early Departure Notice"
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

                If DateTimePicker.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

                If TextBox1.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

            Case "Student Unit Report"
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

                If DateTimePicker.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

                If ComboBox7.Text = "" Then
                    missingFields &= "- Unit" & vbCrLf
                End If

                If ComboBox8.Text = "" Then
                    missingFields &= "- Pass/Fail" & vbCrLf
                End If

            Case "Sent Back to Work Notice"
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

                If DateTimePicker.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

                If TextBox1.Text = "" Then
                    missingFields &= "- Time" & vbCrLf
                End If

                If NotesTB.Text = "" Then
                    missingFields &= "- Notes" & vbCrLf
                End If

            Case "Student Investigation"
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

                If DateTimePicker.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

            Case "Class Commencement Reminder"
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

                If DateTimePicker.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

            Case "Yearly Student Report"

                If GetLastReportDate(studentID) Is Nothing Then
                    LastReportDatePicker.Show()
                    Exit Sub
                End If
                If ComboBox12.Text = "" Then
                    missingFields &= "- Email Subject" & vbCrLf
                End If

                If teacherNameComboBox.Text = "" Then
                    missingFields &= "- Teacher" & vbCrLf
                End If

                If DateTimePicker.Text = "" Then
                    missingFields &= "- Date" & vbCrLf
                End If

                If ComboBox4.Text = "" Then
                    missingFields &= "- Attendance/Puncuality" & vbCrLf
                End If

                If ComboBox5.Text = "" Then
                    missingFields &= "- Class Room Engagement" & vbCrLf
                End If

                If ComboBox6.Text = "" Then
                    missingFields &= "- Course Progression" & vbCrLf
                End If


                ' Additional validation logic specific to ComboBoxOption2
                ' Cases for ComboBoxOption3 to ComboBoxOption12...
        End Select



        ' If all fields are filled, proceed with submission

        SendOutlookEmail.SendOutlookEmail(StudentIDLBL.Text, StudentFirstnameLBL.Text, StudentSurnameLBL.Text, StudentEmailLBL.Text, EmployerFirstnameLBL.Text, EmployerSurnameLBL.Text, EmployerBusinessNameLBL.Text, EmployerEmailLBL.Text)
        Dim selectedStudent As String = StudentCB.SelectedItem.ToString()
        AbsentEarlyLateLog(selectedStudent)
        ComboBox12.Text = ""
        teacherNameComboBox.Text = ""
        DateTimePicker.Value = Date.Today ' Set to today's date or any default date you prefer
        ComboBox4.Text = ""
        TextBox1.Text = ""
        ComboBox5.Text = ""
        ComboBox6.Text = ""
        ComboBox7.Text = ""
        ComboBox8.Text = ""
        NotesTB.Text = ""
    End Sub

    Public Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If ComboBox12.Text = "Student Investigation" Then
            ' Declare variables to store responses
            Dim response1 As DialogResult
            Dim response2 As DialogResult
            Dim response3 As DialogResult
            Dim response3Text As String = ""

            ' Show message boxes to get responses
            response1 = MessageBox.Show("Has the student been emailed?", "Emailed Student?", MessageBoxButtons.YesNo)
            response2 = MessageBox.Show("Have you Rang the student?", "Rang Student?", MessageBoxButtons.YesNo)
            response3 = MessageBox.Show("Has any other form of contact been made?", "Other Contact?", MessageBoxButtons.YesNo)

            ' If response3 is Yes, prompt for additional information
            If response3 = DialogResult.Yes Then
                Dim inputForm As New System.Windows.Forms.Form()
                Dim textBox As New System.Windows.Forms.TextBox()
                Dim okButton As New System.Windows.Forms.Button()
                inputForm.Size = New Size(400, 200) ' Set form size
                textBox.Location = New System.Drawing.Point(40, 40) ' Set location to (40, 40)
                textBox.Size = New Size(320, 20) ' Set textbox size
                okButton.Location = New System.Drawing.Point(150, 90) ' Set location to (150, 90)
                okButton.Text = "OK"
                okButton.DialogResult = DialogResult.OK
                inputForm.Text = "What type of contact was made?"
                inputForm.FormBorderStyle = FormBorderStyle.FixedDialog
                inputForm.StartPosition = FormStartPosition.CenterScreen
                inputForm.Controls.Add(textBox)
                inputForm.Controls.Add(okButton)
                inputForm.AcceptButton = okButton

                If inputForm.ShowDialog() = DialogResult.OK Then
                    response3Text = textBox.Text
                End If

            End If

            ' Write responses to SQL database
            WriteToSQL(response1, response2, response3, response3Text, StudentIDLBL.Text)
        End If
        SendOutlookEmail.SendOutlookEmail(StudentIDLBL.Text, StudentFirstnameLBL.Text, StudentSurnameLBL.Text, StudentEmailLBL.Text, EmployerFirstnameLBL.Text, EmployerSurnameLBL.Text, EmployerBusinessNameLBL.Text, EmployerEmailLBL.Text)




    End Sub

    Public Sub WriteToSQL(response1 As DialogResult, response2 As DialogResult, response3 As DialogResult, response3Text As String, studentID As String)
        Try
            ' Check if the student ID exists in the table
            Dim queryCheck As String = "SELECT COUNT(*) FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"

            Using commandCheck As New SqlCommand(queryCheck, connection)
                commandCheck.Parameters.AddWithValue("@StudentID", studentID)
                Dim count As Integer = CInt(commandCheck.ExecuteScalar())

                If count > 0 Then
                    ' Student ID exists, so update the existing row
                    Dim queryUpdate As String = "UPDATE ElectrotechnologyReports.dbo.StudentLogs SET [Apptrain-Studentbeenemailed] = @Response1, [AppTrain-Haveyourangstudents] = @Response2, [AppTrain-OtherFormofcontact] = @Response3, [AppTrain-OtherText] = @Response3Text WHERE [Student ID] = @StudentID"
                    Using commandUpdate As New SqlCommand(queryUpdate, connection)
                        commandUpdate.Parameters.AddWithValue("@Response1", If(response1 = DialogResult.Yes, "Yes", "No"))
                        commandUpdate.Parameters.AddWithValue("@Response2", If(response2 = DialogResult.Yes, "Yes", "No"))
                        commandUpdate.Parameters.AddWithValue("@Response3", If(response3 = DialogResult.Yes, "Yes", "No"))
                        commandUpdate.Parameters.AddWithValue("@Response3Text", response3Text)
                        commandUpdate.Parameters.AddWithValue("@StudentID", studentID)
                        commandUpdate.ExecuteNonQuery()
                    End Using
                Else
                    ' Student ID does not exist, so insert a new row
                    Dim queryInsert As String = "INSERT INTO ElectrotechnologyReports.dbo.StudentLogs ([Student ID], [Apptrain-Studentbeenemailed], [AppTrain-Haveyourangstudents], [AppTrain-OtherFormofcontact], [AppTrain-OtherText]) VALUES (@StudentID, @Response1, @Response2, @Response3, @Response3Text)"
                    Using commandInsert As New SqlCommand(queryInsert, connection)
                        commandInsert.Parameters.AddWithValue("@StudentID", studentID)
                        commandInsert.Parameters.AddWithValue("@Response1", If(response1 = DialogResult.Yes, "Yes", "No"))
                        commandInsert.Parameters.AddWithValue("@Response2", If(response2 = DialogResult.Yes, "Yes", "No"))
                        commandInsert.Parameters.AddWithValue("@Response3", If(response3 = DialogResult.Yes, "Yes", "No"))
                        commandInsert.Parameters.AddWithValue("@Response3Text", response3Text)
                        commandInsert.ExecuteNonQuery()
                    End Using
                End If
            End Using

            MessageBox.Show("Responses written to SQL database successfully.")
        Catch ex As System.Exception
            ' Handle exceptions
            MessageBox.Show("Error writing responses to SQL database: " & ex.Message)
        End Try
    End Sub
    ' Method to retrieve responses from the database
    Public Function GetResponses(studentID As String) As Tuple(Of String, String, String, String)
        Dim query As String = "SELECT [Apptrain-Studentbeenemailed], [AppTrain-Haveyourangstudents], [AppTrain-OtherFormofcontact], [AppTrain-OtherText] FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"

        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@StudentID", studentID)
                Using reader As SqlDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        Dim response1 As String = reader.GetString(0)
                        Dim response2 As String = reader.GetString(1)
                        Dim response3 As String = reader.GetString(2)
                        Dim response3Text As String = reader.GetString(3)
                        Return New Tuple(Of String, String, String, String)(response1, response2, response3, response3Text)
                    Else
                        ' Handle case where student ID is not found
                        Return Nothing
                    End If
                End Using
            End Using
        End Using
    End Function



    Public Sub ResetInvestigation()
        ' Call the method to reset the date in the Excel sheet back to blank for the matching student ID
        ResetApptrainData()
    End Sub
    Private Sub ResetApptrainData()
        Try
            ' SQL query to reset the date and investigation report for the student
            Dim query As String = "UPDATE ElectrotechnologyReports.dbo.StudentLogs " &
                              "SET LastStudentReportDate = NULL, " &
                              "[Apptrain-Studentbeenemailed] = NULL, " &
                              "[AppTrain-Haveyourangstudents] = NULL, " &
                              "[Apptrain-OtherFormofcontact] = NULL, " &
                              "[AppTrain-OtherText] = '' " &
                              "WHERE LastStudentReportDate < DATEADD(day, -14, GETDATE())"

            ' Create a SqlCommand object with the query and connection
            Using command As New SqlCommand(query, connection)
                ' Execute the query
                command.ExecuteNonQuery()

                ' Show a message indicating success
                'MessageBox.Show("Date and investigation report reset successfully for all students with LastStudentReportDate exceeding 14 days.")
            End Using
        Catch ex As System.Exception
            ' Handle exceptions
            MessageBox.Show("Error resetting date and investigation report: " & ex.Message)
        End Try
    End Sub

    Private Sub UpdateInvestigationReport()

        Try
            ' SQL query to retrieve the investigation report for the student
            Dim query As String = "SELECT InvestigationReportColumn FROM ElectrotechnologyReports.dbo.StudentLogs WHERE StudentIDColumn = @StudentID"

            ' Create a SqlCommand object with the query and connection
            Using command As New SqlCommand(query, connection)
                ' Add parameter for the student ID
                command.Parameters.AddWithValue("@StudentID", StudentIDLBL.Text)

                ' Execute the query and retrieve the investigation report
                Dim investigationReport As Object = command.ExecuteScalar()

                ' Check if the investigation report is not null
                If investigationReport IsNot Nothing AndAlso Not DBNull.Value.Equals(investigationReport) Then
                    ' Update InvestigationLBL with the retrieved investigation report
                    InvestigationLBL.Text = investigationReport.ToString()
                Else
                    ' If the investigation report is null or DBNull, leave InvestigationLBL blank
                    InvestigationLBL.Text = ""
                End If
            End Using
        Catch ex As System.Exception
            ' Handle exceptions
            MessageBox.Show("Error retrieving investigation report: " & ex.Message)
        End Try
    End Sub


    Private Sub Button9_Click(sender As Object, e As EventArgs)
        SendOutlookEmail.SendOutlookEmail(studentID, studentFirstname, studentSurname, studentEmail, employerFirstname, employerSurname, employerBusinessName, employerEmail)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim Pword As String
        Dim adminPasswords() As String = {"Wpower84", "Admin123"} ' Add your additional passwords here

        Pword = InputBox("Enter your password?")
        If Pword = vbNullString Then
            MsgBox("Logged in as Normal User")
            Exit Sub
        End If

        Dim isAdmin As Boolean = False
        For Each pass As String In adminPasswords
            If Pword = pass Then
                isAdmin = True
                Exit For
            End If
        Next

        If isAdmin Then
            'Place Admin code here
            MsgBox("Logged in as Administrator")
            If Pword = "Wpower84" Then
                ' Open SettingsForm
                Dim settingsForm As New SettingsForm()
                settingsForm.Show()
            ElseIf Pword = "Admin123" Then
                ' Open AdminForm
                Dim adminForm As New Admin()
                Admin.Show()
            End If
        Else
            MsgBox("Incorrect Password, logged in as normal user")
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        StudentUnits.Show()
    End Sub
    Private Sub PopulateBlockGroupCB()
        Dim uniqueBlockGroupCodes As New HashSet(Of String)()

        ' Clear the ComboBox before repopulating only if it's not empty
        If BlockGroupCB.Items.Count > 0 Then
            BlockGroupCB.Items.Clear()
            StudentAmendment.ComboBox1.Items.Clear()
        End If

        Dim commandText As String = "SELECT DISTINCT [Block Group Code] FROM ElectrotechnologyReports.dbo.AgreementsDetails"
        Using command As New SqlCommand(commandText, connection)
            Using reader As SqlDataReader = command.ExecuteReader()
                While reader.Read()
                    If Not reader.IsDBNull(0) Then
                        Dim blockGroupCode As String = reader.GetString(0)
                        ' Check if the block group code is not already in the HashSet
                        If Not uniqueBlockGroupCodes.Contains(blockGroupCode) Then
                            BlockGroupCB.Items.Add(blockGroupCode)
                            StudentAmendment.ComboBox1.Items.Add(blockGroupCode)
                            uniqueBlockGroupCodes.Add(blockGroupCode)
                        End If
                    End If
                End While
            End Using
        End Using
    End Sub
    Private Sub BlockGroupCB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BlockGroupCB.SelectedIndexChanged
        Button11.Visible = False
        StudentCB.Text = ""
        Label19.Text = ""
        Label22.Text = ""
        Label23.Text = ""

        ' Show the loading form
        Dim loadingForm As New LoadingForm()
        loadingForm.Show()
        ' Define custom increments
        Dim totalSteps As Integer = 100 ' Total number of steps
        Dim currentStep As Integer = 5  ' Current step
        loadingForm.Label1.Text = "Loading..."

        ' Populate StudentCB ComboBox based on selected Block Group
        Dim selectedBlockGroup As String = BlockGroupCB.SelectedItem.ToString()
        PopulateStudentCB(selectedBlockGroup)
        currentStep = 50
        Label4.Visible = True
        StudentCB.Visible = True
        SelectedStudentLBL.Visible = True
        GroupBox1.Visible = True

        loadingForm.Label1.Text = "Loading Complete!"
        loadingForm.UpdateProgress(totalSteps)
        ' Simulate a delay
        System.Threading.Thread.Sleep(100)
        'CheckVersionAndDisplayInfo()
        ' Close the loading form once loading is finished
        loadingForm.Close()
        If BlockGroupCB.Text = "" Then
            PopulateBlockGroupCB()
        End If

    End Sub

    Private Sub PopulateStudentCB(blockGroup As String)
        ' Clear existing items in StudentCB
        StudentCB.Items.Clear()

        ' Query the database to retrieve students for the selected block group
        Dim commandText As String = "SELECT [Student ID], [Student Given Name], [Student Family Name], [Student Personal Email] FROM ElectrotechnologyReports.dbo.AgreementsDetails WHERE [Block Group Code] = @BlockGroup"
        Using command As New SqlCommand(commandText, connection)
            ' Add parameter for block group code
            command.Parameters.AddWithValue("@BlockGroup", blockGroup)

            Using reader As SqlDataReader = command.ExecuteReader()
                While reader.Read()
                    ' Initialize variables to store student information
                    Dim studentID As String = ""
                    Dim givenName As String = ""
                    Dim familyName As String = ""
                    Dim email As String = ""

                    ' Check for null values before retrieving data from reader
                    If Not reader.IsDBNull(0) Then studentID = Convert.ToString(reader.GetValue(0))
                    If Not reader.IsDBNull(1) Then givenName = reader.GetString(1)
                    If Not reader.IsDBNull(2) Then familyName = reader.GetString(2)
                    If Not reader.IsDBNull(3) Then email = reader.GetString(3)

                    ' Format the student information
                    Dim studentInfo As String = $"{studentID} - {givenName} {familyName} - {email}"

                    ' Add formatted student information to StudentCB
                    StudentCB.Items.Add(studentInfo)
                End While
            End Using
        End Using
    End Sub

    Private Async Sub StudentCB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles StudentCB.SelectedIndexChanged
        Label19.Text = ""
        Label22.Text = ""
        Label23.Text = ""

        If StudentCB.SelectedIndex < 0 OrElse StudentCB.SelectedItem Is Nothing Then
            Button11.Visible = False
            Return
        End If

        ' Show the loading form
        Dim loadingForm As New LoadingForm()
        loadingForm.Show()
        ' Define custom increments
        Dim totalSteps As Integer = 100 ' Total number of steps
        Dim currentStep As Integer = 5  ' Current step
        loadingForm.Label1.Text = "Loading..."
        ' Populate labels with selected student's information
        Dim selectedStudent As String = StudentCB.SelectedItem.ToString()


        UpdateLabels(selectedStudent)
        UpdateExemplarProfilingEmailButtonVisibility()
        ExemplarEmailOverrides.InvalidateCacheForStudent(StudentIDLBL.Text?.Trim())

        ' Update SelectedStudentLBL
        SelectedStudentLBL.Text = selectedStudent

        Dim studentID As String = StudentIDLBL.Text
        currentStep = 50
        Button10.Visible = True
        Label15.Visible = True
        Label19.Visible = True
        Label18.Visible = True
        Label16.Visible = True
        Label22.Visible = True
        Label20.Visible = True
        Label17.Visible = True
        Label23.Visible = True
        Label21.Visible = True
        Button3.Visible = True
        Button10.Visible = True
        Label5.Visible = True
        ComboBox12.Visible = True
        Label7.Visible = True
        teacherNameComboBox.Visible = True
        Button5.Visible = True
        ' Call the method to update the investigation report label when the selection in studentCB is changed
        'UpdateInvestigationReport()
        loadingForm.Label1.Text = "Loading Complete!"
        loadingForm.UpdateProgress(totalSteps)
        ' Simulate a delay
        System.Threading.Thread.Sleep(100)
        'CheckVersionAndDisplayInfo()
        ' Close the loading form once loading is finished
        loadingForm.Close()
        ' Check the order of checked items and update label
        Dim studentUnitsForm As New StudentUnits()

        ' Check the order of checked items and update label
        'CheckCompletionOrderAndUpdateLabel(StudentIDLBL.Text, UnitAlertLbl, StudentUnits.UnitAlertLbl1, StudentUnits.CheckedListBox1)
        CompletionChecker.LoadCheckBoxStates(studentID)
        ResitModule.CheckResit(StudentIDLBL.Text, resitLabel)
        If StudentCB.Text = "" Then
            PopulateBlockGroupCB()
        End If

        Await RefreshSelectedStudentProfilingAsync()
        UpdateExemplarProfilingEmailButtonVisibility()

        '-------------------Need to look at below----------------------------

    End Sub
    Private Sub UpdateLabels(selectedStudent As String)
        ' Check if there is data in StudentLogs table for the matching student ID
        Dim hasLogs As Boolean = CheckStudentLogs(selectedStudent)

        ' SQL query to retrieve student and employer information based on the selected student
        Dim query As String = "SELECT [Student ID], [Student Given Name], [Student Family Name], [Student Personal Email], " &
                      "[Employer Given Name], [Employer Name], [Employer Email], [Block Group Code], [Student Personal Mobile], [Apprenticeship Client ID]" &
                      " FROM ElectrotechnologyReports.dbo.AgreementsDetails WHERE [Student ID] = @StudentID"

        ' Create a SqlCommand object with the query and connection
        Using command As New SqlCommand(query, connection)
            ' Add parameters to the query
            command.Parameters.AddWithValue("@StudentID", selectedStudent.Split("-"c)(0).Trim())

            ' Execute the query and retrieve the data
            Using reader As SqlDataReader = command.ExecuteReader()
                If reader.Read() Then
                    ' Update labels with retrieved student and employer information
                    Dim studentID As Double = If(Not reader.IsDBNull(0), reader.GetDouble(0), 0) ' Check for DBNull
                    StudentIDLBL.Text = studentID.ToString() ' Convert to string

                    StudentFirstnameLBL.Text = If(Not reader.IsDBNull(1), reader.GetString(1), "")
                    StudentSurnameLBL.Text = If(Not reader.IsDBNull(2), reader.GetString(2), "")
                    StudentEmailLBL.Text = If(Not reader.IsDBNull(3), reader.GetString(3), "")
                    EmployerFirstnameLBL.Text = If(Not reader.IsDBNull(4), reader.GetString(4), "")
                    EmployerBusinessNameLBL.Text = If(Not reader.IsDBNull(5), reader.GetString(5), "")
                    EmployerEmailLBL.Text = If(Not reader.IsDBNull(6), reader.GetString(6), "")
                    BlockGroupLBL.Text = If(Not reader.IsDBNull(7), reader.GetString(7), "")
                    Dim mobileNumber As String = If(Not reader.IsDBNull(8), reader.GetDouble(8).ToString(), "")
                    ' Check if the string represents a mobile number (assuming a mobile number has 10 digits)
                    If mobileNumber.Length = 9 AndAlso mobileNumber.All(Function(c) Char.IsDigit(c)) Then
                        mobileNumber = "0" & mobileNumber ' Add "0" in front of the mobile number
                    End If
                    Label29.Text = mobileNumber

                    Label34.Text = If(Not reader.IsDBNull(9), reader.GetString(9), "")
                    '-----------------------------------------------------------------
                    'Enable once Address field is in SQL database, dont forget to add [Address] in the query field
                    'Label28.Text = If(Not reader.IsDBNull(10), reader.GetString(10), "")

                    '------------------------------------------------------------------
                    ' Update SelectedStudentLBL with selected student's information
                    SelectedStudentLBL.Text = selectedStudent

                    ' Make InvestigationLBL visible if there are logs for the selected student
                    InvestigationLBL.Visible = hasLogs
                End If
            End Using
        End Using
        ' Call AbsentEarlyLateLog to update absence, late arrival, and early departure logs
        AbsentEarlyLateLog(selectedStudent)
        Label28.Text = GetStudentAddress(selectedStudent)
    End Sub

    Private Function GetStudentAddress(selectedStudent As String) As String
        Dim addressList As New List(Of String)

        ' SQL query to retrieve student address from StudentLogs table based on the selected student ID
        Dim query As String = "SELECT StudentAddress FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"

        ' Create a SqlCommand object with the query and connection
        Using command As New SqlCommand(query, connection)
            ' Add parameters to the query
            ' Assuming selectedStudent is in the format "1234-5678", where "1234" is the float part
            Dim studentID As Double = Double.Parse(selectedStudent.Split("-"c)(0).Trim())
            command.Parameters.AddWithValue("@StudentID", studentID)

            ' Execute the query and retrieve the student address
            Using reader As SqlDataReader = command.ExecuteReader()
                While reader.Read()
                    Dim address As Object = reader("StudentAddress")
                    If address IsNot Nothing AndAlso address IsNot DBNull.Value Then
                        addressList.Add(address.ToString())
                    End If
                End While
            End Using
        End Using

        ' Return the first address found, if any
        If addressList.Count > 0 Then
            Return addressList(0)
        Else
            Return ""
        End If
    End Function





    Private Function AbsentEarlyLateLog(selectedStudent As String) As Boolean
        ' Check if there is data in StudentLogs table for the matching student ID
        Dim hasLogs As Boolean = CheckStudentLogs(selectedStudent)
        Dim query As String = "SELECT  [Absent], [Late Arrival], [Early Departure] " &
                      "FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Student ID] = @StudentID"
        ' Create a SqlCommand object with the query and connection
        Using command As New SqlCommand(query, connection)
            ' Add parameters to the query
            command.Parameters.AddWithValue("@StudentID", selectedStudent.Split("-"c)(0).Trim())
            Using reader As SqlDataReader = command.ExecuteReader()
                If reader.Read() Then
                    ' Update labels with retrieved student and employer information
                    Dim absent As Integer = If(Not reader.IsDBNull(0), reader.GetInt32(0), 0)
                    Dim lateArrival As Integer = If(Not reader.IsDBNull(1), reader.GetInt32(1), 0)
                    Dim earlyDeparture As Integer = If(Not reader.IsDBNull(2), reader.GetInt32(2), 0)

                    ' Update labels with log information
                    Label23.Text = absent
                    Label19.Text = lateArrival
                    Label22.Text = earlyDeparture
                End If
            End Using
        End Using
        Return hasLogs
    End Function

    Private Function CheckStudentLogs(selectedStudent As String) As Boolean
        Dim connectionString As String = SQLCon.connectionString
        Dim query As String = "SELECT COUNT(*) FROM ElectrotechnologyReports.dbo.StudentLogs WHERE [Apptrain-Studentbeenemailed] = 'Yes' AND [Student ID] = @StudentID"

        Using connectionLogs As New SqlConnection(connectionString)
            connectionLogs.Open()

            Using commandLogs As New SqlCommand(query, connectionLogs)
                commandLogs.Parameters.AddWithValue("@StudentID", selectedStudent.Split("-"c)(0).Trim())

                ' Execute the query and retrieve the count
                Dim count As Integer = 0
                Using reader As SqlDataReader = commandLogs.ExecuteReader()
                    If reader.Read() Then
                        count = Convert.ToInt32(reader(0))
                    End If
                End Using

                Return count > 0
            End Using
        End Using
    End Function




    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Close the SQL connection if it's open
        ' Check if the connection object is not null and its state is open before attempting to close it
        If connection IsNot Nothing AndAlso connection.State = ConnectionState.Open Then
            connection.Close()
        End If
    End Sub

    Private Sub PopulateEmailSubjectComboBox()
        ' Clear existing items in the ComboBox
        ComboBox12.Items.Clear()

        ' SQL query to retrieve all teachers' names
        Dim query As String = "SELECT EmailSubject FROM ElectrotechnologyReports.dbo.EmailTemplates ORDER BY EmailSubject ASC"

        ' Create a SqlCommand object with the query and connection
        Using command As New SqlCommand(query, connection)
            ' Execute the query and retrieve the data
            Using reader As SqlDataReader = command.ExecuteReader()
                While reader.Read()
                    ' Add each Email Subjects's name to the ComboBox
                    ComboBox12.Items.Add(reader.GetString(0))
                    EmailSubjectHelp.ComboBox1.Items.Add(reader.GetString(0))
                End While
            End Using
        End Using
    End Sub
    Private Sub PopulateTeacherComboBox()
        ' Clear existing items in the ComboBox
        teacherNameComboBox.Items.Clear()

        ' SQL query to retrieve all teachers' names
        Dim query As String = "SELECT Teacher_Full_Name FROM ElectrotechnologyReports.dbo.TeacherList WHERE Highest_Certificate_Taught = 'Certificate III' ORDER BY Teacher_Full_Name ASC"

        ' Create a SqlCommand object with the query and connection
        Using command As New SqlCommand(query, connection)
            ' Execute the query and retrieve the data
            Using reader As SqlDataReader = command.ExecuteReader()
                While reader.Read()
                    ' Add each teacher's name to the ComboBox
                    teacherNameComboBox.Items.Add(reader.GetString(0))
                    'StudentUnits.ComboBox1.Items.Add(reader.GetString(0))
                End While
            End Using
        End Using
    End Sub

    Private Sub teacherNameComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles teacherNameComboBox.SelectedIndexChanged
        ' Update labels when a teacher is selected from the ComboBox
        PopulateTeacherCB(teacherNameComboBox.SelectedItem.ToString())
    End Sub
    Private Sub PopulateTeacherCB(selectedTeacher As String)
        ' SQL query to retrieve teacher information based on the selected teacher name
        Dim query As String = "SELECT Teacher_Full_Name, Email FROM ElectrotechnologyReports.dbo.TeacherList WHERE Teacher_Full_Name = @TeacherName"

        ' Create a SqlCommand object with the query and connection
        Using command As New SqlCommand(query, connection)
            ' Add parameter for the selected teacher name
            command.Parameters.AddWithValue("@TeacherName", selectedTeacher)

            ' Execute the query and retrieve the data
            Using reader As SqlDataReader = command.ExecuteReader()
                If reader.Read() Then
                    ' Update labels with retrieved teacher information
                    TeacherNameLBL.Text = reader.GetString(0)
                    teacherEmailLabel.Text = reader.GetString(1)
                End If
            End Using
        End Using
    End Sub
    Private Sub Populateunit()
        Try
            ' SQL query to retrieve data from your database
            Dim query As String = "SELECT Unit_Code, Unit_Title FROM ElectrotechnologyReports.dbo.UEE30820units"

            ' Create a SqlCommand object with the query and connection
            Using command As New SqlCommand(query, connection)
                ' Execute the query and retrieve the data
                Using reader As SqlDataReader = command.ExecuteReader()
                    ' Loop through the result set and populate ComboBox7 with data from columns UnitCodeColumn and UnitNameColumn
                    While reader.Read()
                        Dim unitCode As String = reader.GetString(0) ' Assuming UnitCodeColumn is of type string
                        Dim unitName As String = reader.GetString(1) ' Assuming UnitNameColumn is of type string
                        ComboBox7.Items.Add(unitCode & " - " & unitName)
                    End While
                End Using
            End Using
        Catch ex As System.Exception
            ' Handle exceptions
            MessageBox.Show("Error populating ComboBox7: " & ex.Message)
        End Try
    End Sub


    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        ' Get the selected item from ComboBox7
        Dim selectedItem As String = ComboBox7.SelectedItem.ToString()

        ' Split the selected item into unit code and unit name
        Dim parts() As String = selectedItem.Split("-"c)
        Dim unitCode As String = parts(0).Trim()
        Dim unitName As String = parts(1).Trim()

        ' Update Label30 with unit code and Label31 with unit name
        Label30.Text = unitCode
        Label31.Text = unitName
    End Sub
    Private Sub PopulateWeekdays()
        ' Clear existing items in the ComboBox
        ' ComboBox1.Items.Clear()

        ' Get the current date
        ' Dim currentDate As Date = Date.Today

        ' Loop through each weekday for two weeks before and two weeks after the current date
        ' For i As Integer = -14 To 14 ' Two weeks before and two weeks after the current date
        'Dim day As Date = currentDate.AddDays(i)
        'Dim weekday As String = day.ToString("dd MMMM, yyyy")

        ' Add the weekday to the ComboBox
        'ComboBox1.Items.Add(weekday)
        ' Next

        ' Set the default selected item to today's date
        'ComboBox1.SelectedItem = currentDate.ToString("dd MMMM, yyyy")
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged
        DisplayMode()
    End Sub
    Private Sub DisplayMode()
        If ComboBox12.Text = "Student Term Progress Report" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = True
            Label6.Visible = False
            DateTimePicker.Visible = True
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = True
            ComboBox4.Visible = True
            Label9.Visible = True
            ComboBox5.Visible = True
            Label10.Visible = True
            ComboBox6.Visible = True
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            Label35.Visible = False
            'Button3.Visible = False
            'Button10.Visible = False
            Button7.Visible = False
            ' If GetLastReportDate(studentID) Is Nothing Then
            ' Last report date is NULL, show the LastReportDatePicker form
            'MessageBox.Show("It seems this student has never had a previous Student Report before, " & vbCrLf &
            '" Please Select the Date When this log started. " & vbCrLf & " 
            '(This is usually the start Of a term date and/or the First date of class the student was enrolled in)" & vbCrLf & "
            'This only needs to be set once.")
            '       End If
        ElseIf ComboBox12.Text = "2 Week Intention Letter" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = True
            Label6.Visible = False
            DateTimePicker.Visible = True
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            'Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = False
            Label35.Visible = False
        ElseIf ComboBox12.Text = "4 Week Intention Letter" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = True
            Label6.Visible = False
            DateTimePicker.Visible = True
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            'Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = False
            Label35.Visible = False
        ElseIf ComboBox12.Text = "Course Withdraw Notice" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = False
            Label6.Visible = False
            DateTimePicker.Visible = False
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            'Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = False
            Label35.Visible = False
        ElseIf ComboBox12.Text = "Student Behaviour Notice" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = True
            Label6.Visible = False
            DateTimePicker.Visible = True
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            'Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = False
            Label35.Visible = False
        ElseIf ComboBox12.Text = "Overdue Fees - Warning" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = False
            Label6.Visible = False
            DateTimePicker.Visible = False
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = False
            Label35.Visible = False
        ElseIf ComboBox12.Text = "Overdue Fees - Sanction" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = False
            Label6.Visible = False
            DateTimePicker.Visible = False
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            'Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = False
            Label35.Visible = False
        ElseIf ComboBox12.Text = "Unit Withdraw Notice" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = True
            Label6.Visible = False
            DateTimePicker.Visible = True
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = True
            ComboBox7.Visible = True
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            'Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = False
            Label35.Visible = False
        ElseIf ComboBox12.Text = "Absent Notice" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = True
            Label6.Visible = False
            DateTimePicker.Visible = True
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            'Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = False
            Label35.Visible = False
        ElseIf ComboBox12.Text = "Late Arrival Notice" Then
            Label24.Visible = True
            Label25.Visible = False
            Label33.Visible = False
            Label6.Visible = False
            DateTimePicker.Visible = True
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = True
            TextBox1.Visible = True
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            Label35.Visible = True
            'Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = False
        ElseIf ComboBox12.Text = "Early Departure Notice" Then
            Label24.Visible = False
            Label25.Visible = True
            Label33.Visible = False
            Label6.Visible = False
            DateTimePicker.Visible = True
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = True
            TextBox1.Visible = True
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            'Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = False
            Label35.Visible = True
        ElseIf ComboBox12.Text = "Sent Back to Work Notice" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = True
            Label6.Visible = False
            DateTimePicker.Visible = True
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = True
            TextBox1.Visible = True
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            'Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = False
            Label35.Visible = True
        ElseIf ComboBox12.Text = "Student Unit Report" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = True
            Label6.Visible = False
            DateTimePicker.Visible = True
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = True
            ComboBox7.Visible = True
            Label13.Visible = True
            ComboBox8.Visible = True
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            'Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = False
            Label35.Visible = False
        ElseIf ComboBox12.Text = "Student Investigation" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = True
            Label6.Visible = False
            DateTimePicker.Visible = True
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = False
            NotesTB.Visible = False
            Button8.Visible = False
            'Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = True
            Label35.Visible = False
        ElseIf ComboBox12.Text = "Class Commencement Reminder" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = False
            Label6.Visible = True
            DateTimePicker.Visible = True
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = True
            ComboBox7.Visible = True
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            'Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = False
            Label35.Visible = False
        ElseIf ComboBox12.Text = "Yearly Student Report" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = True
            Label6.Visible = False
            DateTimePicker.Visible = True
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = True
            ComboBox4.Visible = True
            Label9.Visible = True
            ComboBox5.Visible = True
            Label10.Visible = True
            ComboBox6.Visible = True
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            Label35.Visible = False
            Button7.Visible = False
        ElseIf ComboBox12.Text = "Exemplar Profiling Outstanding Alert" Then
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = False
            Label6.Visible = False
            DateTimePicker.Visible = False
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            'Button3.Visible = True
            'Button10.Visible = False
            Button7.Visible = False
            Label35.Visible = False
        Else
            Label24.Visible = False
            Label25.Visible = False
            Label33.Visible = False
            Label6.Visible = False
            DateTimePicker.Visible = False
            Label15.Visible = True
            Label19.Visible = True
            Label18.Visible = True
            Label16.Visible = True
            Label22.Visible = True
            Label20.Visible = True
            Label17.Visible = True
            Label23.Visible = True
            Label21.Visible = True
            Label8.Visible = False
            ComboBox4.Visible = False
            Label9.Visible = False
            ComboBox5.Visible = False
            Label10.Visible = False
            ComboBox6.Visible = False
            Label11.Visible = False
            TextBox1.Visible = False
            Label12.Visible = False
            ComboBox7.Visible = False
            Label13.Visible = False
            ComboBox8.Visible = False
            Label32.Visible = True
            NotesTB.Visible = True
            Button8.Visible = True
            Button3.Visible = False
            Button10.Visible = False
            Button7.Visible = False
            Label35.Visible = False




        End If
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs)
        ' Close the connection
        CloseConnection(connection)
        UpdateReconnectButtonVisibility()
    End Sub

    Private Sub MainFrm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        ' Check if the connection object is not null and its state is open before attempting to close it
        If connection IsNot Nothing AndAlso connection.State = ConnectionState.Open Then
            connection.Close()
        End If
        ' Store the state of the checkbox in application settings
        My.Settings.MassEmail = SettingsForm.MassEmailChkBx.Checked
        ' Save the settings
        My.Settings.Save()
    End Sub


    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles StudentIDTextBox.TextChanged
        Dim searchCriteria As String = StudentIDTextBox.Text.Trim()
        Dim query As String = "SELECT DISTINCT [Block Group Code] AS BlockGroup, [Student ID] AS StudentID FROM ElectrotechnologyReports.dbo.AgreementsDetails WHERE [Student ID] LIKE @searchCriteria OR [Block Group Code] IN (SELECT [Block Group Code] FROM ElectrotechnologyReports.dbo.AgreementsDetails WHERE [Student ID] LIKE @searchCriteria)"

        Using connection As New SqlConnection(SQLCon.connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@searchCriteria", "%" & searchCriteria & "%")
                connection.Open()
                Using reader As SqlDataReader = command.ExecuteReader()
                    BlockGroupCB.Items.Clear()
                    StudentCB.Items.Clear()
                    While reader.Read()
                        BlockGroupCB.Items.Add(reader("BlockGroup").ToString())
                        StudentCB.Items.Add(reader("StudentID").ToString())
                    End While
                End Using
            End Using
        End Using
        ' Check if the TextBox is empty
        If StudentIDTextBox.Text = "" Then
            PopulateBlockGroupCB()
        End If

    End Sub
    Private Sub Button6_Click_2(sender As Object, e As EventArgs) Handles Button6.Click
        Dim searchCriteria As String = StudentIDTextBox.Text.Trim()
        Dim searchCriteriaFloat As Single

        If Single.TryParse(searchCriteria, searchCriteriaFloat) Then
            Dim blockGroupQuery As String = "SELECT DISTINCT [Block Group Code] AS BlockGroup FROM ElectrotechnologyReports.dbo.AgreementsDetails WHERE [Student ID] = @searchCriteria"
            Dim studentQuery As String = "SELECT DISTINCT CAST([Student ID] AS NVARCHAR(255)) AS StudentID FROM ElectrotechnologyReports.dbo.AgreementsDetails WHERE CAST([Student ID] AS FLOAT) = @searchCriteriaFloat"

            Using connection As New SqlConnection(SQLCon.connectionString)
                connection.Open()

                ' Retrieve distinct block group codes
                Using blockGroupCommand As New SqlCommand(blockGroupQuery, connection)
                    blockGroupCommand.Parameters.AddWithValue("@searchCriteria", searchCriteriaFloat)
                    Using blockGroupReader As SqlDataReader = blockGroupCommand.ExecuteReader()
                        BlockGroupCB.Items.Clear()
                        While blockGroupReader.Read()
                            Dim blockGroup As String = blockGroupReader("BlockGroup").ToString()
                            BlockGroupCB.Items.Add(blockGroup)
                        End While
                    End Using
                End Using

                ' Retrieve distinct student IDs matching the StudentIDTextBox value
                Using studentCommand As New SqlCommand(studentQuery, connection)
                    studentCommand.Parameters.AddWithValue("@searchCriteriaFloat", searchCriteriaFloat)
                    Using studentReader As SqlDataReader = studentCommand.ExecuteReader()
                        StudentCB.Items.Clear()
                        While studentReader.Read()
                            Dim studentID As String = studentReader("StudentID").ToString()
                            StudentCB.Items.Add(studentID)
                        End While
                    End Using
                End Using
            End Using

            ' Automatically select the first item if available
            If BlockGroupCB.Items.Count > 0 Then
                BlockGroupCB.SelectedIndex = 0
            End If
            Dim searchText As String = searchCriteria.Split(" "c)(0)
            Dim matchingIndex As Integer = -1
            For i As Integer = 0 To StudentCB.Items.Count - 1
                Dim studentIDFromCB As String = StudentCB.Items(i).ToString().Split(" "c)(0)
                If searchText = studentIDFromCB Then
                    matchingIndex = i
                    Exit For
                End If
            Next

            If matchingIndex <> -1 Then
                StudentCB.SelectedIndex = matchingIndex
            Else
                BlockGroupCB.Items.Clear()
                BlockGroupCB.Text = ""
                BlockGroupCB.SelectedIndex = -1 ' Clear selected value
                StudentCB.Items.Clear()
                StudentCB.Text = ""
                StudentCB.SelectedIndex = -1 ' Clear selected value
                Button11.Visible = False
                MessageBox.Show("Student ID Doesn't Exist")
            End If
        Else
            MessageBox.Show("Please enter a valid integer value for Student ID.")
        End If
        ResitModule.CheckResit(StudentIDLBL.Text, resitLabel)
    End Sub






    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        BugReport.Show()
    End Sub

    Private Sub VersionLBL_Click(sender As Object, e As EventArgs) Handles VersionLBL.Click

    End Sub

    Private Sub MassEmailBtn_Click(sender As Object, e As EventArgs) Handles MassEmailBtn.Click
        Dim OutApp As Object
        Dim OutMail As Object
        Dim body As String
        Dim body1 As String
        Dim imageData As Byte() = RetrieveImageDataFromDatabase()
        If BlockGroupCB.SelectedItem Is Nothing Then
            MsgBox("Please select a Blockgroup for Mass Email")
            Exit Sub
        Else
            ' Prompt the user for the email body message
            Dim bodyMessage As String = InputBox("Enter the email body message:", "Email Body")

            ' Check if the user cancelled the input
            If String.IsNullOrEmpty(bodyMessage) Then
                MsgBox("No email body message entered. Mass email cancelled.")
                Exit Sub
            End If
            Dim blockGroupCode As String = BlockGroupCB.SelectedItem.ToString()

            ' Extract the text after the underscore
            Dim blockGroupText As String = ""
            Dim parts As String() = blockGroupCode.Split("_"c)
            If parts.Length > 1 Then
                blockGroupText = parts(1)
            End If

            ' SQL connection string
            Dim connectionString As String = SQLCon.connectionString

            ' SQL query to fetch student email addresses based on block group code
            Dim query As String = "SELECT [Student Personal Email] FROM AgreementsDetails WHERE [Block Group Code] = @BlockGroupCode"

            Try
                ' Create SQL connection
                Using connection As New SqlConnection(connectionString)
                    connection.Open()

                    ' Create SQL command
                    Using command As New SqlCommand(query, connection)
                        ' Add parameter for block group code
                        command.Parameters.AddWithValue("@BlockGroupCode", blockGroupCode)

                        ' Execute SQL command
                        Using reader As SqlDataReader = command.ExecuteReader()
                            Dim outlookApp As New Application()
                            Dim mail As MailItem = outlookApp.CreateItem(OlItemType.olMailItem)

                            ' Initialize StringBuilder for BCC field
                            Dim bccBuilder As New StringBuilder()

                            ' Add student email addresses to StringBuilder
                            While reader.Read()
                                ' Add each email address to the StringBuilder
                                Dim emailAddress As String = reader("Student Personal Email").ToString()
                                If Not String.IsNullOrEmpty(emailAddress) Then
                                    If bccBuilder.Length > 0 Then
                                        bccBuilder.Append("; ")
                                    End If
                                    bccBuilder.Append(emailAddress)
                                End If
                            End While

                            ' Create a new instance of Outlook Application
                            'Dim outlookApp As New Outlook.Application()

                            OutApp = CreateObject("Outlook.Application")
                            ' Create a new email item
                            OutMail = OutApp.CreateItem(0)

                            ' Set email properties
                            With OutMail
                                .To = ""
                                .cc = ""
                                .bcc = bccBuilder.ToString()
                                .Subject = "Attention Class: " & blockGroupText
                                body = "Attention All Students In Class " & blockGroupText & "<BR><BR>"
                                body1 = bodyMessage
                                .HTMLbody = body & body1 & "<br><br><br><br><img src='data:image/jpeg;base64," & Convert.ToBase64String(imageData) & "' width='90%'> " & Me.VersionLBL.Text
                                .Display ' Display the email
                            End With

                            ' Release COM objects
                            Marshal.ReleaseComObject(OutMail)
                            Marshal.ReleaseComObject(OutApp)

                            ' Clean up
                            OutMail = Nothing
                            OutApp = Nothing

                        End Using
                    End Using
                End Using
            Catch ex As System.Exception
                ' Handle any exceptions
                MsgBox("An error occurred: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub Label34_Click(sender As Object, e As EventArgs) Handles Label34.Click

    End Sub

    Private Sub StudentIDTextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles StudentIDTextBox.KeyPress
        ' Check if the pressed key is Enter
        If e.KeyChar = Convert.ToChar(Keys.Enter) Then
            ' Trigger the click event of the Search button
            Button6.PerformClick()
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

    End Sub
End Class