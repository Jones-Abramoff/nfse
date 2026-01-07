Imports System
Imports System.Configuration

''' <summary>
''' Decide qual emissor usar com mínima interferência no ABRASF.
''' Regras:
''' 1) Se a filial estiver em São Paulo (cMun = 3550308) -> Paulistana (SOAP/XML)
''' 2) Se appSettings("NFSE_PADRAO") = "NACIONAL" -> Nacional (JSON)
''' 3) Caso contrário -> ABRASF (fluxo existente)
''' </summary>
Public NotInheritable Class NfseRouter

    Public Const ROTA_ABRASF As String = "ABRASF"
    Public Const ROTA_SAO_PAULO As String = "SAO_PAULO"
    Public Const ROTA_NACIONAL As String = "NACIONAL"

    Private Sub New()
    End Sub

    Public Shared Function DetectarRota(ByVal sEmpresa As String, ByVal iFilialEmpresa As Integer) As String
        'Detecção por município do prestador
        Try
            Dim db As New SGEDadosDataContext()
            Dim filial = (From f In db.FiliaisEmpresas Where f.FilialEmpresa = iFilialEmpresa Select f).FirstOrDefault()
            If filial Is Nothing Then Return ROTA_NACIONAL

            Dim endereco = (From e In db.Enderecos Where e.Codigo = filial.Endereco Select e).FirstOrDefault()
            If endereco Is Nothing Then Return ROTA_NACIONAL
            If endereco.Cidade = "São Paulo" Then
                Return ROTA_SAO_PAULO
            End If
        Catch
            'Se falhar, cai no ROTA_NACIONAL
        End Try

        Return ROTA_NACIONAL
    End Function

End Class
