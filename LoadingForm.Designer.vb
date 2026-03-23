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
        VUExemplar = New PictureBox()
        Label2 = New Label()
        Label3 = New Label()
        CType(VUExemplar, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' Label1
        ' 
        Label1.Font = New Font("Segoe UI", 18F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        Label1.Location = New Point(96, 346)
        Label1.Name = "Label1"
        Label1.Size = New Size(445, 32)
        Label1.TabIndex = 0
        Label1.Text = "Loading Please Wait"
        Label1.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' ProgressBar1
        ' 
        ProgressBar1.Location = New Point(96, 396)
        ProgressBar1.Name = "ProgressBar1"
        ProgressBar1.Size = New Size(445, 51)
        ProgressBar1.TabIndex = 1
        ProgressBar1.Value = 5
        ' 
        ' VUExemplar
        ' 
        VUExemplar.Image = My.Resources.Resources.Victoria_UniversityExemplar
        VUExemplar.Location = New Point(96, 72)
        VUExemplar.Name = "VUExemplar"
        VUExemplar.Size = New Size(445, 210)
        VUExemplar.SizeMode = PictureBoxSizeMode.Zoom
        VUExemplar.TabIndex = 2
        VUExemplar.TabStop = False
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Font = New Font("Segoe UI Semibold", 18F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label2.Location = New Point(130, 27)
        Label2.Name = "Label2"
        Label2.Size = New Size(386, 32)
        Label2.TabIndex = 3
        Label2.Text = "VU Student Attendance Reporting"
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(245, 285)
        Label3.Name = "Label3"
        Label3.Size = New Size(126, 15)
        Label3.TabIndex = 4
        Label3.Text = "Created by Frank Offer"
        ' 
        ' LoadingForm
        ' 
        AutoScaleDimensions = New SizeF(96F, 96F)
        AutoScaleMode = AutoScaleMode.Dpi
        ClientSize = New Size(634, 481)
        Controls.Add(Label3)
        Controls.Add(Label2)
        Controls.Add(VUExemplar)
        Controls.Add(ProgressBar1)
        Controls.Add(Label1)
        Name = "LoadingForm"
        StartPosition = FormStartPosition.CenterScreen
        Text = "Loading....."
        CType(VUExemplar, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents VUExemplar As PictureBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
End Class


