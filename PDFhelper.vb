Imports System.Runtime.InteropServices

Public Class PdfHelper
    <DllImport("shell32.dll", EntryPoint:="ShellExecute", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function ShellExecute(ByVal hwnd As IntPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As IntPtr
    End Function

    Public Shared Sub OpenPdfWithDefaultViewer(ByVal filePath As String)
        ShellExecute(IntPtr.Zero, "open", filePath, Nothing, Nothing, 1)
    End Sub
End Class
