Option Strict On
Option Explicit On

Imports System

'Module Program
'    ' Exemplo de uso (Console)
'    Sub Main(args As String())
'        Dim empresa As String = If(args.Length > 0, args(0), "01")
'        Dim filial As Integer = If(args.Length > 1, Integer.Parse(args(1)), 1)
'        Dim numIntDoc As Long = If(args.Length > 2, Long.Parse(args(2)), 0)

'        If numIntDoc = 0 Then
'            Console.WriteLine("Uso: app.exe <Empresa> <FilialEmpresa> <NumIntDoc>")
'            Return
'        End If

'        Dim svc As New NfseNacionalEmissor()
'        Dim result = svc.EmitirNfse(empresa, numIntDoc, filial)

'        Console.WriteLine("OK=" & result.Sucesso)
'        Console.WriteLine("ChaveAcesso=" & result.ChaveAcesso)
'        Console.WriteLine("IdDps=" & result.IdDps)
'        If result.Alertas IsNot Nothing AndAlso result.Alertas.Count > 0 Then
'            For Each a In result.Alertas
'                Console.WriteLine($"ALERTA {a.codigo}: {a.mensagem} ({a.descricao}) {a.complemento}")
'            Next
'        End If
'    End Sub
'End Module
