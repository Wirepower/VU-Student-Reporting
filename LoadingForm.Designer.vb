<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LoadingForm
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
        ProgressBar1 = New ProgressBar()
        SuspendLayout()
        ' 
        ' Label1
        ' 
        Label1.Font = New Font("Segoe UI", 18F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        Label1.Location = New Point(101, 51)
        Label1.Name = "Label1"
        Label1.Size = New Size(445, 32)
        Label1.TabIndex = 0
        Label1.Text = "Loading Please Wait"
        Label1.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' ProgressBar1
        ' 
        ProgressBar1.Location = New Point(101, 132)
        ProgressBar1.Name = "ProgressBar1"
        ProgressBar1.Size = New Size(445, 51)
        ProgressBar1.TabIndex = 1
        ProgressBar1.Value = 5
        ' 
        ' LoadingForm
        ' 
        AutoScaleDimensions = New SizeF(96F, 96F)
        AutoScaleMode = AutoScaleMode.Dpi
        ClientSize = New Size(634, 236)
        Controls.Add(ProgressBar1)
        Controls.Add(Label1)
        Name = "LoadingForm"
        StartPosition = FormStartPosition.CenterScreen
        Text = "Loading....."
        ResumeLayout(False)
    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents ProgressBar1 As ProgressBar
End Class


