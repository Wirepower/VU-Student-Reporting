<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SQLError
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Button10 = New Button()
        Label12 = New Label()
        TextBox5 = New TextBox()
        Label1 = New Label()
        SuspendLayout()
        ' 
        ' Button10
        ' 
        Button10.Location = New Point(546, 207)
        Button10.Name = "Button10"
        Button10.Size = New Size(137, 44)
        Button10.TabIndex = 43
        Button10.Text = "Update String"
        Button10.UseVisualStyleBackColor = True
        ' 
        ' Label12
        ' 
        Label12.AutoSize = True
        Label12.Location = New Point(12, 160)
        Label12.Name = "Label12"
        Label12.Size = New Size(127, 15)
        Label12.TabIndex = 42
        Label12.Text = "SQL Connection String"
        ' 
        ' TextBox5
        ' 
        TextBox5.Location = New Point(12, 178)
        TextBox5.Name = "TextBox5"
        TextBox5.Size = New Size(1247, 23)
        TextBox5.TabIndex = 41
        TextBox5.TextAlign = HorizontalAlignment.Center
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Segoe UI", 21.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label1.Location = New Point(458, 45)
        Label1.Name = "Label1"
        Label1.Size = New Size(312, 40)
        Label1.TabIndex = 44
        Label1.Text = "SQL Connection Error"
        ' 
        ' SQLError
        ' 
        AutoScaleDimensions = New SizeF(7.0F, 15.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1295, 263)
        Controls.Add(Label1)
        Controls.Add(Button10)
        Controls.Add(Label12)
        Controls.Add(TextBox5)
        Name = "SQLError"
        Text = "SQLError"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Button10 As Button
    Friend WithEvents Label12 As Label
    Friend WithEvents TextBox5 As TextBox
    Friend WithEvents Label1 As Label
End Class
