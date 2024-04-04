Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class LoadingForm
    ' Public method to update the progress bar value
    Public Sub UpdateProgress(value As Integer)
        ProgressBar1.Value = value
    End Sub
End Class