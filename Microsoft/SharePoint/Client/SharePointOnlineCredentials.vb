Imports System.Security

Namespace Microsoft.SharePoint.Client
    Friend Class SharePointOnlineCredentials
        Private userName As String
        Private securePassword As SecureString

        Public Sub New(userName As String, securePassword As SecureString)
            Me.userName = userName
            Me.securePassword = securePassword
        End Sub
    End Class
End Namespace
