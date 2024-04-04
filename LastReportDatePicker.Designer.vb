<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LastReportDatePicker
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
        DateTimePicker1 = New DateTimePicker()
        OKbutton = New Button()
        CancelButton = New Button()
        Label1 = New Label()
        Label2 = New Label()
        SuspendLayout()
        ' 
        ' DateTimePicker1
        ' 
        DateTimePicker1.Location = New Point(237, 225)
        DateTimePicker1.Name = "DateTimePicker1"
        DateTimePicker1.Size = New Size(346, 23)
        DateTimePicker1.TabIndex = 0
        ' 
        ' OKbutton
        ' 
        OKbutton.Location = New Point(237, 315)
        OKbutton.Name = "OKbutton"
        OKbutton.Size = New Size(112, 65)
        OKbutton.TabIndex = 1
        OKbutton.Text = "Submit"
        OKbutton.UseVisualStyleBackColor = True
        ' 
        ' CancelButton
        ' 
        CancelButton.Location = New Point(471, 315)
        CancelButton.Name = "CancelButton"
        CancelButton.Size = New Size(112, 65)
        CancelButton.TabIndex = 2
        CancelButton.Text = "Cancel"
        CancelButton.UseVisualStyleBackColor = True
        ' 
        ' Label1
        ' 
        Label1.Location = New Point(181, 127)
        Label1.Name = "Label1"
        Label1.Size = New Size(456, 76)
        Label1.TabIndex = 3
        Label1.Text = "It seems this student has never had a previous Student Report before, Please select the date when this log started. (usually the start of a term, first date of class)"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Font = New Font("Segoe UI", 18F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label2.Location = New Point(250, 45)
        Label2.Name = "Label2"
        Label2.Size = New Size(306, 32)
        Label2.TabIndex = 4
        Label2.Text = "Enter Date of Start of Log"
        ' 
        ' LastReportDatePicker
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Controls.Add(CancelButton)
        Controls.Add(OKbutton)
        Controls.Add(DateTimePicker1)
        Name = "LastReportDatePicker"
        Text = "LastReportDatePicker"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents OKbutton As Button
    Friend WithEvents CancelButton As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
End Class
