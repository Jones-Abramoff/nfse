Imports System
Imports System.Configuration
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Xml

''' <summary>
''' Cancelamento (São Paulo - Paulistana).
'''
''' Observação:
''' - O pedido de cancelamento do município de São Paulo varia por versão.
''' - Como você não anexou um XML de exemplo de cancelamento, aqui deixei um esqueleto
'''   que já faz o POST SOAP e registra logs em RPSWEBLoteLog.
''' - Você só precisa ajustar o XML montado em <see cref="MontarPedidoCancelamento"/>
'''   de acordo com o Manual 3.3.4.
''' </summary>
Public Class NfsePaulistanaCancelamento

    Public Function CancelarNfsePaulistana(ByVal iFilialEmpresa As Integer, ByVal numeroNFe As String, ByVal codigoVerificacao As String, ByVal motivo As String) As String
        Dim db As New SGEDadosDataContext()

        Dim ambiente As String = GetAppSetting("NFSE_AMBIENTE", "HOMOLOG")
        Dim url As String = If(ambiente.ToUpperInvariant().Contains("PROD"),
                               GetAppSetting("NFSE_SP_URL_PRODUCAO", ""),
                               GetAppSetting("NFSE_SP_URL_HOMOLOGACAO", ""))
        If String.IsNullOrWhiteSpace(url) Then
            Throw New Exception("Configurar NFSE_SP_URL_HOMOLOGACAO e/ou NFSE_SP_URL_PRODUCAO no app.config")
        End If

        Dim xmlPedido As String = MontarPedidoCancelamento(numeroNFe, codigoVerificacao, motivo)
        Dim xmlResp As String = EnviarSoap(url, "CancelamentoNFe", xmlPedido)

        'Log
        Try
            Dim log As New RPSWEBLoteLog()
            log.FilialEmpresa = CShort(iFilialEmpresa)
            log.Lote = 0
            log.Data = Date.Today
            log.Hora = CDbl(Date.Now.ToString("HHmmss"))
            log.status = "CANC_SP_XML"
            log.NumIntNF = 0
            log.NumIntDoc = 0
            log.Ambiente = 2
            db.RPSWEBLoteLogs.InsertOnSubmit(log)
            db.SubmitChanges()
        Catch
        End Try

        Return xmlResp
    End Function

    Private Function MontarPedidoCancelamento(ByVal numeroNFe As String, ByVal codigoVerificacao As String, ByVal motivo As String) As String
        'TODO: Ajustar conforme Manual 3.3.4.
        'Modelo mínimo (PLACEHOLDER):
        Dim doc As New XmlDocument()
        doc.PreserveWhitespace = False
        Dim ns As String = "http://www.prefeitura.sp.gov.br/nfe"

        Dim root = doc.CreateElement("PedidoCancelamentoNFe", ns)
        doc.AppendChild(root)

        Dim cab = doc.CreateElement("Cabecalho")
        cab.SetAttribute("Versao", "2")
        root.AppendChild(cab)

        Dim chave = doc.CreateElement("ChaveNFe")
        root.AppendChild(chave)

        Dim n1 = doc.CreateElement("NumeroNFe")
        n1.InnerText = numeroNFe
        chave.AppendChild(n1)

        Dim n2 = doc.CreateElement("CodigoVerificacao")
        n2.InnerText = codigoVerificacao
        chave.AppendChild(n2)

        Dim m = doc.CreateElement("Motivo")
        m.InnerText = If(motivo, "")
        root.AppendChild(m)

        Return doc.OuterXml
    End Function

    Private Function EnviarSoap(ByVal url As String, ByVal metodo As String, ByVal xmlConteudo As String) As String
        Dim soap As String =
    "<?xml version=""1.0"" encoding=""utf-8""?>" &
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " &
    "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " &
    "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &
    "<soap:Body>" &
    "<" & metodo & " xmlns=""http://www.prefeitura.sp.gov.br/nfe"">" &
    "<xml><![CDATA[" & xmlConteudo & "]]></xml>" &
    "</" & metodo & ">" &
    "</soap:Body>" &
    "</soap:Envelope>"

        Dim req = CType(WebRequest.Create(url), HttpWebRequest)
        req.Method = "POST"
        req.ContentType = "text/xml; charset=utf-8"
        req.Headers.Add("SOAPAction", """http://www.prefeitura.sp.gov.br/nfe/" & metodo & """")

        Dim bytes = Encoding.UTF8.GetBytes(soap)
        req.ContentLength = bytes.Length
        Using rs = req.GetRequestStream()
            rs.Write(bytes, 0, bytes.Length)
        End Using

        Using resp = CType(req.GetResponse(), HttpWebResponse)
            Using sr As New StreamReader(resp.GetResponseStream(), Encoding.UTF8)
                Return sr.ReadToEnd()
            End Using
        End Using
    End Function

    Private Function GetAppSetting(ByVal key As String, ByVal defaultValue As String) As String
        Try
            Dim v = ConfigurationManager.AppSettings(key)
            If String.IsNullOrEmpty(v) Then Return defaultValue
            Return v
        Catch
            Return defaultValue
        End Try
    End Function

End Class
