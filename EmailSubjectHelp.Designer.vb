<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EmailSubjectHelp
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
        ComboBox1 = New ComboBox()
        Label1 = New Label()
        Label2 = New Label()
        Button2 = New Button()
        SuspendLayout()
        ' 
        ' ComboBox1
        ' 
        ComboBox1.FormattingEnabled = True
        ComboBox1.Location = New Point(383, 46)
        ComboBox1.Name = "ComboBox1"
        ComboBox1.Size = New Size(361, 23)
        ComboBox1.TabIndex = 0
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        Label1.Location = New Point(60, 44)
        Label1.Name = "Label1"
        Label1.Size = New Size(317, 21)
        Label1.TabIndex = 1
        Label1.Text = "What ""Email Subject"" will you like help with?"
        ' 
        ' Label2
        ' 
        Label2.BackColor = SystemColors.HighlightText
        Label2.BorderStyle = BorderStyle.Fixed3D
        Label2.Font = New Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        Label2.Location = New Point(60, 104)
        Label2.Name = "Label2"
        Label2.Size = New Size(684, 224)
        Label2.TabIndex = 2
        ' 
        ' Button2
        ' 
        Button2.Location = New Point(273, 354)
        Button2.Name = "Button2"
        Button2.Size = New Size(241, 81)
        Button2.TabIndex = 4
        Button2.Text = "Return"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' EmailSubjectHelp
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 462)
        Controls.Add(Button2)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Controls.Add(ComboBox1)
        Name = "EmailSubjectHelp"
        Text = "Email Subject Help"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Button2 As Button
End Class
