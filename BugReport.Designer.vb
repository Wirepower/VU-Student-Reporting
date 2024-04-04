<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BugReport
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
        Button1 = New Button()
        Button2 = New Button()
        Label1 = New Label()
        Label2 = New Label()
        Label3 = New Label()
        IssueTB = New TextBox()
        HowTB = New TextBox()
        OtherTB = New TextBox()
        Label4 = New Label()
        Label5 = New Label()
        ReportByTB = New ComboBox()
        SuspendLayout()
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(296, 594)
        Button1.Name = "Button1"
        Button1.Size = New Size(254, 62)
        Button1.TabIndex = 0
        Button1.Text = "Submit"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' Button2
        ' 
        Button2.Location = New Point(296, 671)
        Button2.Name = "Button2"
        Button2.Size = New Size(254, 62)
        Button2.TabIndex = 1
        Button2.Text = "Cancel"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Microsoft Sans Serif", 8.25F)
        Label1.Location = New Point(124, 127)
        Label1.Name = "Label1"
        Label1.Size = New Size(91, 13)
        Label1.TabIndex = 2
        Label1.Text = "What is the issue:"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(42, 250)
        Label2.Name = "Label2"
        Label2.Size = New Size(179, 15)
        Label2.TabIndex = 3
        Label2.Text = "How can this issue be replicated:"
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(42, 381)
        Label3.Name = "Label3"
        Label3.Size = New Size(173, 15)
        Label3.TabIndex = 4
        Label3.Text = "Any other relevant information:"
        ' 
        ' IssueTB
        ' 
        IssueTB.Location = New Point(221, 123)
        IssueTB.Multiline = True
        IssueTB.Name = "IssueTB"
        IssueTB.Size = New Size(441, 106)
        IssueTB.TabIndex = 5
        ' 
        ' HowTB
        ' 
        HowTB.Location = New Point(221, 250)
        HowTB.Multiline = True
        HowTB.Name = "HowTB"
        HowTB.Size = New Size(441, 106)
        HowTB.TabIndex = 6
        ' 
        ' OtherTB
        ' 
        OtherTB.Location = New Point(221, 378)
        OtherTB.Multiline = True
        OtherTB.Name = "OtherTB"
        OtherTB.Size = New Size(441, 106)
        OtherTB.TabIndex = 7
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Font = New Font("Segoe UI Black", 21.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label4.Location = New Point(349, 40)
        Label4.Name = "Label4"
        Label4.Size = New Size(180, 40)
        Label4.TabIndex = 8
        Label4.Text = "Bug Report"
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(67, 511)
        Label5.Name = "Label5"
        Label5.Size = New Size(151, 15)
        Label5.TabIndex = 9
        Label5.Text = "Who is making this report?:"
        ' 
        ' ReportByTB
        ' 
        ReportByTB.FormattingEnabled = True
        ReportByTB.Location = New Point(224, 511)
        ReportByTB.Name = "ReportByTB"
        ReportByTB.Size = New Size(438, 23)
        ReportByTB.TabIndex = 10
        ' 
        ' BugReport
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 747)
        Controls.Add(ReportByTB)
        Controls.Add(Label5)
        Controls.Add(Label4)
        Controls.Add(OtherTB)
        Controls.Add(HowTB)
        Controls.Add(IssueTB)
        Controls.Add(Label3)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Name = "BugReport"
        Text = "BugReport"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents IssueTB As TextBox
    Friend WithEvents HowTB As TextBox
    Friend WithEvents OtherTB As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents ReportByTB As ComboBox
End Class
