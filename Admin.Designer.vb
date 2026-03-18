<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Admin
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
        AdminLbl = New Label()
        Label4 = New Label()
        TradesAdminTB = New TextBox()
        Label2 = New Label()
        Label1 = New Label()
        txtApptrainEmail = New TextBox()
        txtAdminEmail = New TextBox()
        DataGridView1 = New DataGridView()
        Button5 = New Button()
        btnCancel = New Button()
        Label3 = New Label()
        Label5 = New Label()
        CType(DataGridView1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' AdminLbl
        ' 
        AdminLbl.AutoSize = True
        AdminLbl.Font = New Font("Segoe UI", 21.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        AdminLbl.Location = New Point(316, 42)
        AdminLbl.Name = "AdminLbl"
        AdminLbl.Size = New Size(342, 40)
        AdminLbl.TabIndex = 0
        AdminLbl.Text = "Administration Settings"
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(48, 309)
        Label4.Name = "Label4"
        Label4.Size = New Size(114, 15)
        Label4.TabIndex = 21
        Label4.Text = "Trades Admin Email:"
        ' 
        ' TradesAdminTB
        ' 
        TradesAdminTB.Location = New Point(168, 306)
        TradesAdminTB.Name = "TradesAdminTB"
        TradesAdminTB.Size = New Size(261, 23)
        TradesAdminTB.TabIndex = 20
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(68, 253)
        Label2.Name = "Label2"
        Label2.Size = New Size(94, 15)
        Label2.TabIndex = 19
        Label2.Text = "App-Train Email:"
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(84, 203)
        Label1.Name = "Label1"
        Label1.Size = New Size(78, 15)
        Label1.TabIndex = 18
        Label1.Text = "Admin Email:"
        ' 
        ' txtApptrainEmail
        ' 
        txtApptrainEmail.Location = New Point(168, 250)
        txtApptrainEmail.Name = "txtApptrainEmail"
        txtApptrainEmail.Size = New Size(261, 23)
        txtApptrainEmail.TabIndex = 17
        ' 
        ' txtAdminEmail
        ' 
        txtAdminEmail.Location = New Point(168, 200)
        txtAdminEmail.Name = "txtAdminEmail"
        txtAdminEmail.Size = New Size(261, 23)
        txtAdminEmail.TabIndex = 16
        ' 
        ' DataGridView1
        ' 
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridView1.Location = New Point(465, 184)
        DataGridView1.Name = "DataGridView1"
        DataGridView1.Size = New Size(465, 464)
        DataGridView1.TabIndex = 23
        ' 
        ' Button5
        ' 
        Button5.Location = New Point(608, 135)
        Button5.Name = "Button5"
        Button5.Size = New Size(191, 43)
        Button5.TabIndex = 22
        Button5.Text = "Upload and Load CSV Data"
        Button5.UseVisualStyleBackColor = True
        ' 
        ' btnCancel
        ' 
        btnCancel.Location = New Point(168, 599)
        btnCancel.Name = "btnCancel"
        btnCancel.Size = New Size(261, 49)
        btnCancel.TabIndex = 24
        btnCancel.Text = "Close"
        btnCancel.UseVisualStyleBackColor = True
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(363, 82)
        Label3.Name = "Label3"
        Label3.Size = New Size(203, 15)
        Label3.TabIndex = 25
        Label3.Text = "Student Units Database Current as of:"
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(572, 82)
        Label5.Name = "Label5"
        Label5.Size = New Size(31, 15)
        Label5.TabIndex = 26
        Label5.Text = "Date"
        ' 
        ' Admin
        ' 
        AutoScaleDimensions = New SizeF(96F, 96F)
        AutoScaleMode = AutoScaleMode.Dpi
        AutoScroll = True
        ClientSize = New Size(964, 698)
        Controls.Add(Label5)
        Controls.Add(Label3)
        Controls.Add(btnCancel)
        Controls.Add(DataGridView1)
        Controls.Add(Button5)
        Controls.Add(Label4)
        Controls.Add(TradesAdminTB)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Controls.Add(txtApptrainEmail)
        Controls.Add(txtAdminEmail)
        Controls.Add(AdminLbl)
        Name = "Admin"
        Text = "Admin"
        CType(DataGridView1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents AdminLbl As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents TradesAdminTB As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents txtApptrainEmail As TextBox
    Friend WithEvents txtAdminEmail As TextBox
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button5 As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents Label5 As Label
End Class


