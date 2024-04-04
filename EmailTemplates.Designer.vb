<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EmailTemplates
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
        TextBox1 = New TextBox()
        Button1 = New Button()
        Button2 = New Button()
        Button3 = New Button()
        Button4 = New Button()
        Button5 = New Button()
        SuspendLayout()
        ' 
        ' ComboBox1
        ' 
        ComboBox1.FormattingEnabled = True
        ComboBox1.Location = New Point(264, 78)
        ComboBox1.Name = "ComboBox1"
        ComboBox1.Size = New Size(447, 23)
        ComboBox1.TabIndex = 0
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Segoe UI", 15.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label1.Location = New Point(389, 33)
        Label1.Name = "Label1"
        Label1.Size = New Size(225, 30)
        Label1.TabIndex = 1
        Label1.Text = "Select Email Template"
        ' 
        ' TextBox1
        ' 
        TextBox1.Location = New Point(25, 127)
        TextBox1.Multiline = True
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New Size(925, 537)
        TextBox1.TabIndex = 2
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(277, 670)
        Button1.Name = "Button1"
        Button1.Size = New Size(154, 50)
        Button1.TabIndex = 3
        Button1.Text = "Update"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' Button2
        ' 
        Button2.Location = New Point(25, 671)
        Button2.Name = "Button2"
        Button2.Size = New Size(154, 50)
        Button2.TabIndex = 4
        Button2.Text = "Add New Template"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' Button3
        ' 
        Button3.Location = New Point(796, 671)
        Button3.Name = "Button3"
        Button3.Size = New Size(154, 49)
        Button3.TabIndex = 5
        Button3.Text = "Delete Template"
        Button3.UseVisualStyleBackColor = True
        ' 
        ' Button4
        ' 
        Button4.Location = New Point(535, 669)
        Button4.Name = "Button4"
        Button4.Size = New Size(154, 51)
        Button4.TabIndex = 6
        Button4.Text = "Reset"
        Button4.UseVisualStyleBackColor = True
        ' 
        ' Button5
        ' 
        Button5.Location = New Point(857, 33)
        Button5.Name = "Button5"
        Button5.Size = New Size(93, 67)
        Button5.TabIndex = 7
        Button5.Text = "Close"
        Button5.UseVisualStyleBackColor = True
        ' 
        ' EmailTemplates
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(983, 746)
        Controls.Add(Button5)
        Controls.Add(Button4)
        Controls.Add(Button3)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Controls.Add(TextBox1)
        Controls.Add(Label1)
        Controls.Add(ComboBox1)
        Name = "EmailTemplates"
        Text = "EmailTemplates"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
End Class
