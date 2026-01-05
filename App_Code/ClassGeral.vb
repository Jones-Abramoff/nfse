Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.Schema
Imports System.Data.Odbc
Imports Microsoft.Win32
Imports System.Net



Public Class ADM

    Public Const SUCESSO As Integer = 0


    Public Shared Sub Formata_String_Numero(ByVal sStringRecebe As String, ByRef sStringRetorna As String)

        Dim iTamanho As Integer
        Dim sCaracter As String
        Dim iIndice As Integer

        iTamanho = Len(Trim(sStringRecebe))

        sStringRetorna = ""

        For iIndice = 1 To iTamanho

            sCaracter = Mid(sStringRecebe, iIndice, 1)

            If IsNumeric(sCaracter) Then
                sStringRetorna = sStringRetorna & sCaracter
            End If

        Next

    End Sub

    Public Shared Sub Formata_Sem_Espaco(ByVal sStringRecebe As String, ByRef sStringRetorna As String)

        Dim iTamanho As Integer
        Dim sCaracter As String
        Dim iIndice As Integer

        iTamanho = Len(Trim(sStringRecebe))

        sStringRetorna = ""

        For iIndice = 1 To iTamanho

            sCaracter = Mid(sStringRecebe, iIndice, 1)

            If sCaracter = " " Then
                sStringRetorna = sStringRetorna & "_"
            Else
                sStringRetorna = sStringRetorna & sCaracter
            End If

        Next

    End Sub

    Public Shared Function DesacentuaTexto(ByVal sTexto As String) As String

        'retorna uma copia do texto com a troca dos caracteres acentuados por nao acentuados

        Dim iIndice As Integer
        Dim sCaracter As String
        Dim sGuardaTexto As String
        Dim iCodigo As Integer

        sTexto = Trim(sTexto)

        'Para cada Caracter do Texto
        For iIndice = 1 To Len(sTexto)

            'Seleciona caracter da posição iIndice
            sCaracter = Mid(sTexto, iIndice, 1)

            'Pega codigo ASC do caracter da selecionado acima
            iCodigo = Asc(sCaracter)

            'Verifica se caracter é acentuado
            Select Case iCodigo

                Case 186
                    sCaracter = "."

                Case 192 To 197
                    sCaracter = Chr(65)

                Case 199
                    sCaracter = Chr(67)

                Case 200 To 203
                    sCaracter = Chr(69)

                Case 204 To 207
                    sCaracter = Chr(73)

                Case 210 To 214
                    sCaracter = Chr(79)

                Case 217 To 220
                    sCaracter = Chr(85)

                Case 224 To 229
                    sCaracter = Chr(97)

                Case 231
                    sCaracter = Chr(99)

                Case 232 To 235
                    sCaracter = Chr(101)

                Case 236 To 239
                    sCaracter = Chr(105)

                Case 242 To 246
                    sCaracter = Chr(111)

                Case 249 To 252
                    sCaracter = Chr(117)

            End Select

            If sCaracter <> "." Then
                sGuardaTexto = sGuardaTexto & sCaracter
            End If

        Next

        DesacentuaTexto = sGuardaTexto


    End Function

    Public Shared Function Serie_Sem_E(ByVal sSerie As String) As String
        'retira -E da serie

        Dim sSerieNova As String
        Dim iPos As Integer

        If sSerie = "UN" Then sSerie = "1"

        iPos = InStr(sSerie, "-e")

        If iPos <> 0 Then
            sSerieNova = Mid(sSerie, 1, iPos - 1)
        Else
            sSerieNova = sSerie
        End If

        Serie_Sem_E = sSerieNova


    End Function

    Public Shared Sub Formata_String_AlfaNumerico(ByVal sStringRecebe As String, ByRef sStringRetorna As String)

        Dim iTamanho As Integer
        Dim sCaracter As String
        Dim iIndice As Integer

        iTamanho = Len(Trim(sStringRecebe))

        sStringRetorna = ""

        For iIndice = 1 To iTamanho

            sCaracter = Mid(sStringRecebe, iIndice, 1)

            If IsNumeric(sCaracter) Then
                sStringRetorna = sStringRetorna & sCaracter
            End If

            If UCase(sCaracter) >= "A" And UCase(sCaracter) <= "Z" Then
                sStringRetorna = sStringRetorna & sCaracter
            End If

        Next

    End Sub

End Class

Public Class GetWebRequest_bhHomologacao


    Inherits br.gov.pbh.bhisshomologa.NfseWSService

    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)


    Protected Overrides Function GetWebRequest(ByVal uri As Uri) As System.Net.WebRequest
        Dim webRequest As System.Net.HttpWebRequest

        webRequest = CType(MyBase.GetWebRequest(uri), System.Net.HttpWebRequest)
        Dim currentServicePoint As ServicePoint = webRequest.ServicePoint

        Sleep(100)

        'Setting KeepAlive to false 
        webRequest.KeepAlive = False

        GetWebRequest = webRequest
        currentServicePoint.MaxIdleTime = 1
        currentServicePoint.SetTcpKeepAlive(False, 1, 1)
        currentServicePoint.ConnectionLeaseTimeout = 1


    End Function
End Class

Public Class GetWebRequest_bhProducao

    Inherits bhproducao.NfseWSService

    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)


    Protected Overrides Function GetWebRequest(ByVal uri As Uri) As System.Net.WebRequest
        Dim webRequest As System.Net.HttpWebRequest

        webRequest = CType(MyBase.GetWebRequest(uri), System.Net.HttpWebRequest)
        Dim currentServicePoint As ServicePoint = webRequest.ServicePoint

        Sleep(100)

        'Setting KeepAlive to false 
        webRequest.KeepAlive = False

        GetWebRequest = webRequest
        currentServicePoint.MaxIdleTime = 1
        currentServicePoint.SetTcpKeepAlive(False, 1, 1)
        currentServicePoint.ConnectionLeaseTimeout = 1

    End Function


End Class
