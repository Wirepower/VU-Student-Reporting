<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SettingsForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        txtAdminEmail = New TextBox()
        txtApptrainEmail = New TextBox()
        Label1 = New Label()
        Label2 = New Label()
        btnSave = New Button()
        btnCancel = New Button()
        Label3 = New Label()
        Button1 = New Button()
        Button2 = New Button()
        Button5 = New Button()
        DataGridView1 = New DataGridView()
        TradesAdminTB = New TextBox()
        Label4 = New Label()
        TextBox1 = New TextBox()
        TextBox2 = New TextBox()
        TextBox3 = New TextBox()
        TextBox4 = New TextBox()
        ComboBox1 = New ComboBox()
        ComboBox2 = New ComboBox()
        ComboBox3 = New ComboBox()
        Label5 = New Label()
        Label6 = New Label()
        Label7 = New Label()
        Label8 = New Label()
        Label9 = New Label()
        Label10 = New Label()
        Label11 = New Label()
        Button3 = New Button()
        ComboBox4 = New ComboBox()
        Button4 = New Button()
        Button6 = New Button()
        Button7 = New Button()
        Button8 = New Button()
        MassEmailChkBx = New CheckBox()
        Button9 = New Button()
        TextBox5 = New TextBox()
        Label12 = New Label()
        Button10 = New Button()
        CType(DataGridView1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' txtAdminEmail
        ' 
        txtAdminEmail.Location = New Point(179, 147)
        txtAdminEmail.Name = "txtAdminEmail"
        txtAdminEmail.Size = New Size(261, 23)
        txtAdminEmail.TabIndex = 0
        ' 
        ' txtApptrainEmail
        ' 
        txtApptrainEmail.Location = New Point(179, 197)
        txtApptrainEmail.Name = "txtApptrainEmail"
        txtApptrainEmail.Size = New Size(261, 23)
        txtApptrainEmail.TabIndex = 1
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(95, 150)
        Label1.Name = "Label1"
        Label1.Size = New Size(78, 15)
        Label1.TabIndex = 2
        Label1.Text = "Admin Email:"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(79, 200)
        Label2.Name = "Label2"
        Label2.Size = New Size(94, 15)
        Label2.TabIndex = 3
        Label2.Text = "App-Train Email:"
        ' 
        ' btnSave
        ' 
        btnSave.Location = New Point(179, 292)
        btnSave.Name = "btnSave"
        btnSave.Size = New Size(125, 43)
        btnSave.TabIndex = 4
        btnSave.Text = "Update Emails"
        btnSave.UseVisualStyleBackColor = True
        ' 
        ' btnCancel
        ' 
        btnCancel.Location = New Point(630, 687)
        btnCancel.Name = "btnCancel"
        btnCancel.Size = New Size(177, 61)
        btnCancel.TabIndex = 5
        btnCancel.Text = "Close"
        btnCancel.UseVisualStyleBackColor = True
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Font = New Font("Segoe UI", 24F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label3.Location = New Point(420, 22)
        Label3.Name = "Label3"
        Label3.Size = New Size(141, 45)
        Label3.TabIndex = 6
        Label3.Text = "Settings"
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(79, 687)
        Button1.Name = "Button1"
        Button1.Size = New Size(95, 61)
        Button1.TabIndex = 7
        Button1.TabStop = False
        Button1.Text = "Manual Reset AppTrain Data > 14days"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' Button2
        ' 
        Button2.Location = New Point(315, 292)
        Button2.Name = "Button2"
        Button2.Size = New Size(125, 43)
        Button2.TabIndex = 10
        Button2.Text = "Upload New Email Signature"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' Button5
        ' 
        Button5.Location = New Point(992, 85)
        Button5.Name = "Button5"
        Button5.Size = New Size(191, 43)
        Button5.TabIndex = 12
        Button5.Text = "Upload and Load CSV Data"
        Button5.UseVisualStyleBackColor = True
        ' 
        ' DataGridView1
        ' 
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridView1.Location = New Point(837, 147)
        DataGridView1.Name = "DataGridView1"
        DataGridView1.Size = New Size(489, 464)
        DataGridView1.TabIndex = 13
        ' 
        ' TradesAdminTB
        ' 
        TradesAdminTB.Location = New Point(179, 253)
        TradesAdminTB.Name = "TradesAdminTB"
        TradesAdminTB.Size = New Size(261, 23)
        TradesAdminTB.TabIndex = 14
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(59, 256)
        Label4.Name = "Label4"
        Label4.Size = New Size(114, 15)
        Label4.TabIndex = 15
        Label4.Text = "Trades Admin Email:"
        ' 
        ' TextBox1
        ' 
        TextBox1.Location = New Point(630, 144)
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New Size(177, 23)
        TextBox1.TabIndex = 16
        ' 
        ' TextBox2
        ' 
        TextBox2.Location = New Point(630, 171)
        TextBox2.Name = "TextBox2"
        TextBox2.Size = New Size(177, 23)
        TextBox2.TabIndex = 17
        ' 
        ' TextBox3
        ' 
        TextBox3.Location = New Point(630, 200)
        TextBox3.Name = "TextBox3"
        TextBox3.Size = New Size(177, 23)
        TextBox3.TabIndex = 18
        ' 
        ' TextBox4
        ' 
        TextBox4.Location = New Point(630, 229)
        TextBox4.Name = "TextBox4"
        TextBox4.Size = New Size(177, 23)
        TextBox4.TabIndex = 19
        ' 
        ' ComboBox1
        ' 
        ComboBox1.FormattingEnabled = True
        ComboBox1.Items.AddRange(New Object() {"Electrotechnology", "Engineering"})
        ComboBox1.Location = New Point(630, 258)
        ComboBox1.Name = "ComboBox1"
        ComboBox1.Size = New Size(177, 23)
        ComboBox1.TabIndex = 20
        ' 
        ' ComboBox2
        ' 
        ComboBox2.FormattingEnabled = True
        ComboBox2.Items.AddRange(New Object() {"Certificate II", "Certificate III", "Certificate IV", "Diploma", "Advanced Diploma"})
        ComboBox2.Location = New Point(630, 287)
        ComboBox2.Name = "ComboBox2"
        ComboBox2.Size = New Size(177, 23)
        ComboBox2.TabIndex = 21
        ' 
        ' ComboBox3
        ' 
        ComboBox3.FormattingEnabled = True
        ComboBox3.Items.AddRange(New Object() {"Teacher", "Senior Educator", "Manager ", "Administator"})
        ComboBox3.Location = New Point(630, 316)
        ComboBox3.Name = "ComboBox3"
        ComboBox3.Size = New Size(177, 23)
        ComboBox3.TabIndex = 22
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(520, 145)
        Label5.Name = "Label5"
        Label5.Size = New Size(104, 15)
        Label5.TabIndex = 23
        Label5.Text = "Teacher FullName:"
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Location = New Point(561, 174)
        Label6.Name = "Label6"
        Label6.Size = New Size(63, 15)
        Label6.TabIndex = 24
        Label6.Text = "E Number:"
        ' 
        ' Label7
        ' 
        Label7.AutoSize = True
        Label7.Location = New Point(542, 203)
        Label7.Name = "Label7"
        Label7.Size = New Size(82, 15)
        Label7.TabIndex = 25
        Label7.Text = "Teacher Email:"
        ' 
        ' Label8
        ' 
        Label8.AutoSize = True
        Label8.Location = New Point(525, 232)
        Label8.Name = "Label8"
        Label8.Size = New Size(99, 15)
        Label8.TabIndex = 26
        Label8.Text = "Contact Number:"
        ' 
        ' Label9
        ' 
        Label9.AutoSize = True
        Label9.Location = New Point(550, 261)
        Label9.Name = "Label9"
        Label9.Size = New Size(73, 15)
        Label9.TabIndex = 27
        Label9.Text = "Department:"
        ' 
        ' Label10
        ' 
        Label10.AutoSize = True
        Label10.Location = New Point(482, 290)
        Label10.Name = "Label10"
        Label10.Size = New Size(147, 15)
        Label10.TabIndex = 28
        Label10.Text = "Highest Certificate Taught:"
        ' 
        ' Label11
        ' 
        Label11.AutoSize = True
        Label11.Location = New Point(571, 319)
        Label11.Name = "Label11"
        Label11.Size = New Size(53, 15)
        Label11.TabIndex = 29
        Label11.Text = "Position:"
        ' 
        ' Button3
        ' 
        Button3.DialogResult = DialogResult.TryAgain
        Button3.Location = New Point(630, 345)
        Button3.Name = "Button3"
        Button3.Size = New Size(177, 29)
        Button3.TabIndex = 30
        Button3.Text = "Add Teacher"
        Button3.UseVisualStyleBackColor = True
        ' 
        ' ComboBox4
        ' 
        ComboBox4.FormattingEnabled = True
        ComboBox4.Location = New Point(630, 416)
        ComboBox4.Name = "ComboBox4"
        ComboBox4.Size = New Size(177, 23)
        ComboBox4.TabIndex = 31
        ' 
        ' Button4
        ' 
        Button4.Location = New Point(630, 452)
        Button4.Name = "Button4"
        Button4.Size = New Size(177, 29)
        Button4.TabIndex = 32
        Button4.Text = "Remove Teacher"
        Button4.UseVisualStyleBackColor = True
        ' 
        ' Button6
        ' 
        Button6.Location = New Point(179, 341)
        Button6.Name = "Button6"
        Button6.Size = New Size(261, 43)
        Button6.TabIndex = 33
        Button6.Text = "Email Templates"
        Button6.UseVisualStyleBackColor = True
        ' 
        ' Button7
        ' 
        Button7.Location = New Point(630, 380)
        Button7.Name = "Button7"
        Button7.Size = New Size(177, 29)
        Button7.TabIndex = 34
        Button7.Text = "Save/Update"
        Button7.UseVisualStyleBackColor = True
        ' 
        ' Button8
        ' 
        Button8.Location = New Point(630, 487)
        Button8.Name = "Button8"
        Button8.Size = New Size(178, 53)
        Button8.TabIndex = 35
        Button8.Text = "Reset"
        Button8.UseVisualStyleBackColor = True
        ' 
        ' MassEmailChkBx
        ' 
        MassEmailChkBx.AutoSize = True
        MassEmailChkBx.Location = New Point(211, 399)
        MassEmailChkBx.Name = "MassEmailChkBx"
        MassEmailChkBx.Size = New Size(208, 19)
        MassEmailChkBx.TabIndex = 36
        MassEmailChkBx.Text = "Enable MASS Blockgroup Emailing"
        MassEmailChkBx.UseVisualStyleBackColor = True
        ' 
        ' Button9
        ' 
        Button9.Location = New Point(180, 687)
        Button9.Name = "Button9"
        Button9.Size = New Size(91, 61)
        Button9.TabIndex = 37
        Button9.Text = "Reset Yearly Logs"
        Button9.UseVisualStyleBackColor = True
        ' 
        ' TextBox5
        ' 
        TextBox5.Location = New Point(79, 639)
        TextBox5.Name = "TextBox5"
        TextBox5.Size = New Size(1247, 23)
        TextBox5.TabIndex = 38
        TextBox5.TextAlign = HorizontalAlignment.Center
        ' 
        ' Label12
        ' 
        Label12.AutoSize = True
        Label12.Location = New Point(79, 621)
        Label12.Name = "Label12"
        Label12.Size = New Size(127, 15)
        Label12.TabIndex = 39
        Label12.Text = "SQL Connection String"
        ' 
        ' Button10
        ' 
        Button10.Location = New Point(353, 610)
        Button10.Name = "Button10"
        Button10.Size = New Size(229, 26)
        Button10.TabIndex = 40
        Button10.Text = "Update String and Restart Application"
        Button10.UseVisualStyleBackColor = True
        ' 
        ' SettingsForm
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1338, 760)
        Controls.Add(Button10)
        Controls.Add(Label12)
        Controls.Add(TextBox5)
        Controls.Add(Button9)
        Controls.Add(MassEmailChkBx)
        Controls.Add(Button8)
        Controls.Add(Button7)
        Controls.Add(Button6)
        Controls.Add(Button4)
        Controls.Add(ComboBox4)
        Controls.Add(Button3)
        Controls.Add(Label11)
        Controls.Add(Label10)
        Controls.Add(Label9)
        Controls.Add(Label8)
        Controls.Add(Label7)
        Controls.Add(Label6)
        Controls.Add(Label5)
        Controls.Add(ComboBox3)
        Controls.Add(ComboBox2)
        Controls.Add(ComboBox1)
        Controls.Add(TextBox4)
        Controls.Add(TextBox3)
        Controls.Add(TextBox2)
        Controls.Add(TextBox1)
        Controls.Add(Label4)
        Controls.Add(TradesAdminTB)
        Controls.Add(DataGridView1)
        Controls.Add(Button5)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Controls.Add(Label3)
        Controls.Add(btnCancel)
        Controls.Add(btnSave)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Controls.Add(txtApptrainEmail)
        Controls.Add(txtAdminEmail)
        Name = "SettingsForm"
        Text = "Settings"
        CType(DataGridView1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents txtAdminEmail As TextBox
    Friend WithEvents txtApptrainEmail As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents btnSave As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents TradesAdminTB As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents ComboBox3 As ComboBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents Button3 As Button
    Friend WithEvents ComboBox4 As ComboBox
    Friend WithEvents Button4 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents Button8 As Button
    Friend WithEvents MassEmailChkBx As CheckBox
    Friend WithEvents Button9 As Button
    Friend WithEvents TextBox5 As TextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents Button10 As Button
End Class
