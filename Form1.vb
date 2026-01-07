Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.Schema
Imports System.Net
Imports System.Net.Security
Imports Org.BouncyCastle.Crypto.Tls

Public Class Form1

    Public Const NFE_AMBIENTE_HOMOLOGACAO As Integer = 2
    Public Const NFE_AMBIENTE_PRODUCAO As Integer = 1

    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

    'Define a custom delegate that always returns true
    Private Function AlwaysTrue(sender As Object, certificate As X509Certificate, chain As X509Chain, sslPolicyErrors As SslPolicyErrors) As Boolean
        Return True
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        Net.ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf AlwaysTrue)


        Timer1.Interval = 1000
        Timer1.Start()

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick


        Dim sEmpresa As String
        Dim lLote As Long
        Dim sOperacao As String
        Dim lErro As Long

        Dim lNumIntNF As Long
        Dim sMotivo As String
        Dim iFilialEmpresa As Integer

        Dim objEnvioRPS As ClassEnvioRPS = New ClassEnvioRPS
        Dim objCancelaNFSE As ClassCancelaNFSE = New ClassCancelaNFSE
        Dim objConsultaLoteNFSE As ClassConsultaLoteNFSE = New ClassConsultaLoteNFSE

        Dim arguments As [String]() = Environment.GetCommandLineArgs()
        Dim iTamanho As Integer
        Dim sRetorno As String
        Dim iDebug As Integer

        Try

            Timer1.Stop()

            'os valores abaixo sao setados para depuracao  
            'simulando a chamada pela aplicacao vb6.
            'De acordo com o tipo da operacao descomente e atribua os valores devidos,
            'sempre preencha sOperacao, sEmpresa e iFilialEmpresa
            'Para sOperacao "Envio" ou "Consulta" preencha o lote
            'Para sOperacao "Cancela" preencha o NumIntNF e o motivo

            'p/todas as operacoes
            sOperacao = "Envio"       ' "Inutiliza" '"Envio"  '"Consulta"   'ou Envio ou Cancela(NF)
            sEmpresa = 1
            iFilialEmpresa = 1

            '' '' '' ' '' ''  ''p/envio ou consulta de lote
            lLote = 481
            ' ''p/cancelamento
            'lNumIntNF = 5924
            'sMotivo = "erro na emissao"

            'MsgBox("vai ler os parametros")

            'os valores abaixo vem da aplicacao normal em vb6
            'comente as linhas abaixo para depuracao
            'sOperacao = arguments(1)
            'sEmpresa = arguments(2)
            'iFilialEmpresa = CInt(arguments(3))

            'MsgBox("leu os parametros")

            'MsgBox("sOperacao = " & sOperacao)
            'MsgBox("sEmpresa = " & sEmpresa)
            'MsgBox("iFilialEmpresa = " & iFilialEmpresa)

            iTamanho = 255
            sRetorno = StrDup(iTamanho, Chr(0))

            iDebug = 0

            Call GetPrivateProfileString("Geral", "Debug", 0, sRetorno, iTamanho, "Adm100.ini")

            If IsNumeric(sRetorno) Then
                iDebug = CInt(sRetorno)
            End If


            If iDebug = 1 Then MsgBox("sOperacao  = " & sOperacao)
            If iDebug = 1 Then MsgBox("sEmpresa = " & sEmpresa)
            If iDebug = 1 Then MsgBox("iFilialEmpresa = " & iFilialEmpresa)

            If sOperacao = "Envio" Then

                lLote = CLng(arguments(4))

                Lote.Text = lLote

                If iDebug = 1 Then MsgBox("Lote = " & lLote)

                lErro = objEnvioRPS.Envia_Lote_RPS(sEmpresa, lLote, iFilialEmpresa)

                If lErro = ADM.SUCESSO Then

#If TATUI Then
    'coloca para esperar a consulta por 5s pois o servidor de tatui é lento
    Sleep (5000)
#End If

                    objConsultaLoteNFSE.Consulta_Lote_NFSE(sEmpresa, lLote, iFilialEmpresa)
                End If

            ElseIf sOperacao = "Cancela" Then

                lNumIntNF = CLng(arguments(4))

                sMotivo = arguments(5)

                objCancelaNFSE.Cancela_NFSE(sEmpresa, lNumIntNF, sMotivo, iFilialEmpresa)


            ElseIf sOperacao = "Consulta" Then

                lLote = CLng(arguments(4))

                Lote.Text = lLote

                If iDebug = 1 Then MsgBox("Lote = " & lLote)

                objConsultaLoteNFSE.Consulta_Lote_NFSE(sEmpresa, lLote, iFilialEmpresa)

            End If

        Catch ex As Exception
            Msg.Items.Add("Erro na execucao: " & ex.Message)
        End Try

    End Sub

End Class
