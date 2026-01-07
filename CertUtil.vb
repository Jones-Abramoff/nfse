Option Strict On
Option Explicit On

Imports System
Imports System.Security.Cryptography.X509Certificates

Public NotInheritable Class CertUtil
    Private Sub New()
    End Sub

    ' Busca no "CurrentUser\My" e "LocalMachine\My" por Subject contendo o texto
    ' ou por thumbprint exato (sem espaços).
    Public Shared Function BuscaPorSubjectOuThumbprint(chave As String) As X509Certificate2
        If String.IsNullOrWhiteSpace(chave) Then Return Nothing
        Dim k = chave.Trim()
        Dim thumb = k.Replace(" ", "").ToUpperInvariant()

        Dim cert = Busca(StoreLocation.CurrentUser, k, thumb)
        If cert IsNot Nothing Then Return cert

        Return Busca(StoreLocation.LocalMachine, k, thumb)
    End Function

    Private Shared Function Busca(loc As StoreLocation, subjectLike As String, thumb As String) As X509Certificate2
        Using store As New X509Store(StoreName.My, loc)
            store.Open(OpenFlags.ReadOnly)
            For Each c In store.Certificates
                If c.Thumbprint IsNot Nothing AndAlso c.Thumbprint.Replace(" ", "").ToUpperInvariant() = thumb Then
                    Return c
                End If
                If c.Subject IsNot Nothing AndAlso c.Subject.ToUpperInvariant().Contains(subjectLike.ToUpperInvariant()) Then
                    Return c
                End If
            Next
        End Using
        Return Nothing
    End Function
End Class
