<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class StudentInvestigation
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
        Label1 = New Label()
        Button1 = New Button()
        SuspendLayout()
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Segoe UI", 24F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        Label1.Location = New Point(236, 27)
        Label1.Name = "Label1"
        Label1.Size = New Size(322, 45)
        Label1.TabIndex = 0
        Label1.Text = "Student Investigation"
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(236, 539)
        Button1.Name = "Button1"
        Button1.Size = New Size(363, 54)
        Button1.TabIndex = 1
        Button1.Text = "Close"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' StudentInvestigation
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 605)
        Controls.Add(Button1)
        Controls.Add(Label1)
        Name = "StudentInvestigation"
        Text = "StudentInvestigation"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Button1 As Button
End Class
