Imports System
Imports System.Configuration
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Xml

''' <summary>
''' Emissor para São Paulo (Nota Fiscal Paulistana)
''' - Envio em lote via SOAP
''' - XML conforme exemplo (PedidoEnvioLoteRPS) e extensão Reforma (IBSCBS/valores/trib/gIBSCBS/cClassTrib)
''' - CST fixo 000 e cClassTrib fixo 000001 (conforme sua regra)
'''
''' Observação importante:
''' - A tag <Assinatura> do RPS (Paulistana) é uma assinatura "de campos" (não é XMLDSig).
'''   A rotina abaixo implementa a assinatura conforme padrão comum do manual (RSA/SHA1)
'''   e foi deixada isolada para ajuste fino caso o seu manual tenha máscara/ordem diferente.
''' </summary>
Public Class NfsePaulistanaEmitter

    Private Const COD_MUN_SAO_PAULO As Integer = 3550308
    Private Const CST_IBSCBS As String = "000"
    Private Const CCLASSTRIB As String = "000001"

    Public Function Envia_Lote_RPS_Paulistana(ByVal sEmpresa As String, ByVal lLote As Long, ByVal iFilialEmpresa As Integer) As Long
        Dim db As New SGEDadosDataContext()

        Dim itensLote = (From r In db.RPSWEBLotes
                         Where r.FilialEmpresa = iFilialEmpresa AndAlso r.Lote = CInt(lLote)
                         Order By r.NumIntNF
                         Select r).ToList()

        If itensLote Is Nothing OrElse itensLote.Count = 0 Then
            Throw New Exception("Lote não encontrado em RPSWEBLote.")
        End If

        Dim filial = (From f In db.FiliaisEmpresas Where f.FilialEmpresa = iFilialEmpresa Select f).FirstOrDefault()
        If filial Is Nothing Then Throw New Exception("FilialEmpresa não encontrada.")

        Dim endFilial = (From e In db.Enderecos Where e.Codigo = filial.Endereco Select e).FirstOrDefault()
        If endFilial Is Nothing Then Throw New Exception("Endereço da filial não encontrado.")

        'Dim cidadeFilial = (From c In db.Cidades Where c.NumIntCidade = endFilial.Cidade Select c).FirstOrDefault()
        'If cidadeFilial Is Nothing Then Throw New Exception("Cidade da filial não encontrada.")
        'If cidadeFilial.CodMunIBGE <> COD_MUN_SAO_PAULO Then
        '    Throw New Exception("Rota Paulistana acionada, mas a filial não é São Paulo (3550308).")
        'End If

        'Configurações mínimas
        Dim cnpjPrestador As String = LimparNumeros(filial.CGC)
        Dim imPrestador As String = LimparNumeros(filial.InscricaoMunicipal)
        Dim serieRps As String = GetAppSetting("NFSE_SP_SERIE", "1")
        Dim ambiente As String = GetAppSetting("NFSE_AMBIENTE", "HOMOLOG")

        'Endpoints (você ajusta no app.config)
        Dim url As String = If(ambiente.ToUpperInvariant().Contains("PROD"),
                               GetAppSetting("NFSE_SP_URL_PRODUCAO", ""),
                               GetAppSetting("NFSE_SP_URL_HOMOLOGACAO", ""))
        If String.IsNullOrWhiteSpace(url) Then
            Throw New Exception("Configurar NFSE_SP_URL_HOMOLOGACAO e/ou NFSE_SP_URL_PRODUCAO no app.config")
        End If

        Dim certNome As String = GetAppSetting("NFSE_CERT_NOME", "")
        Dim cert As X509Certificate2 = (New Certificado()).BuscaNome(certNome)
        If cert Is Nothing OrElse Not cert.HasPrivateKey Then
            Throw New Exception("Certificado não encontrado ou sem chave privada. Ajuste NFSE_CERT_NOME.")
        End If

        Dim xmlPedido As String = MontarPedidoEnvioLoteRps(db, itensLote, cnpjPrestador, imPrestador, serieRps, cert)
        Dim xmlResp As String = EnviarSoap(url, "EnvioLoteRPS", xmlPedido)

        GravarLogLote(db, iFilialEmpresa, lLote, 0, "ENVIO_SP_XML", xmlPedido)
        GravarLogLote(db, iFilialEmpresa, lLote, 0, "RET_SP_XML", xmlResp)

        ProcessarRetorno(db, iFilialEmpresa, lLote, xmlResp)

        Return lLote
    End Function

    Private Function MontarPedidoEnvioLoteRps(ByVal db As SGEDadosDataContext,
                                             ByVal itensLote As List(Of RPSWEBLote),
                                             ByVal cnpjPrestador As String,
                                             ByVal imPrestador As String,
                                             ByVal serieRps As String,
                                             ByVal cert As X509Certificate2) As String

        Dim doc As New XmlDocument()
        doc.PreserveWhitespace = False

        Dim ns As String = "http://www.prefeitura.sp.gov.br/nfe"
        Dim root = doc.CreateElement("PedidoEnvioLoteRPS", ns)
        root.SetAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
        root.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
        doc.AppendChild(root)

        'Cabecalho (conforme exemplo)
        Dim cab = doc.CreateElement("Cabecalho")
        cab.SetAttribute("Versao", "2")

        Dim cpfCnpjRem = doc.CreateElement("CPFCNPJRemetente")
        Dim cnpjNode = doc.CreateElement("CNPJ")
        cnpjNode.InnerText = cnpjPrestador
        cpfCnpjRem.AppendChild(cnpjNode)
        cab.AppendChild(cpfCnpjRem)

        Dim trans = doc.CreateElement("transacao")
        trans.InnerText = "false"
        cab.AppendChild(trans)

        'dtInicio/dtFim - usa range do próprio lote
        Dim datas = (From r In itensLote
                     Let nf = (From n In db.NFeNFiscals Where n.NumIntDoc = r.NumIntNF Select n).FirstOrDefault()
                     Where nf IsNot Nothing AndAlso nf.DataEmissao.HasValue
                     Select nf.DataEmissao.Value).ToList()

        Dim dtIni As Date = If(datas.Count > 0, datas.Min(), Date.Today)
        Dim dtFim As Date = If(datas.Count > 0, datas.Max(), Date.Today)

        Dim nDtIni = doc.CreateElement("dtInicio")
        nDtIni.InnerText = dtIni.ToString("yyyy-MM-dd")
        cab.AppendChild(nDtIni)

        Dim nDtFim = doc.CreateElement("dtFim")
        nDtFim.InnerText = dtFim.ToString("yyyy-MM-dd")
        cab.AppendChild(nDtFim)

        Dim qtd = doc.CreateElement("QtdRPS")
        qtd.InnerText = itensLote.Count.ToString()
        cab.AppendChild(qtd)

        root.AppendChild(cab)

        'RPSs
        For Each item In itensLote
            Dim nf = (From n In db.NFeNFiscals Where n.NumIntDoc = item.NumIntNF Select n).FirstOrDefault()
            If nf Is Nothing Then Continue For

            Dim tom As FiliaisCliente = Nothing
            If nf.Cliente.HasValue AndAlso nf.FilialCli.HasValue Then
                tom = (From t In db.FiliaisClientes Where t.CodCliente = nf.Cliente.Value AndAlso t.CodFilial = nf.FilialCli.Value Select t).FirstOrDefault()
            End If

            Dim rps = doc.CreateElement("RPS")

            'Assinatura (Paulistana)
            Dim assinaturaCampos As String = MontarStringAssinatura(imPrestador, serieRps, nf.NumNotaFiscal.GetValueOrDefault(0), nf.DataEmissao, nf.ValorTotal)
            Dim assinatura As String = AssinarCamposBase64(assinaturaCampos, cert)
            Dim nAss = doc.CreateElement("Assinatura")
            nAss.InnerText = assinatura
            rps.AppendChild(nAss)

            Dim chave = doc.CreateElement("ChaveRPS")
            Dim nIM = doc.CreateElement("InscricaoPrestador")
            nIM.InnerText = imPrestador
            chave.AppendChild(nIM)

            Dim nSerie = doc.CreateElement("SerieRPS")
            nSerie.InnerText = serieRps
            chave.AppendChild(nSerie)

            Dim nNum = doc.CreateElement("NumeroRPS")
            nNum.InnerText = nf.NumNotaFiscal.GetValueOrDefault(item.NumIntNF).ToString()
            chave.AppendChild(nNum)
            rps.AppendChild(chave)

            Dim tipoRps = doc.CreateElement("TipoRPS")
            tipoRps.InnerText = "RPS"
            rps.AppendChild(tipoRps)

            Dim dataEmi = doc.CreateElement("DataEmissao")
            dataEmi.InnerText = If(nf.DataEmissao.HasValue, nf.DataEmissao.Value.ToString("yyyy-MM-dd"), Date.Today.ToString("yyyy-MM-dd"))
            rps.AppendChild(dataEmi)

            Dim status = doc.CreateElement("StatusRPS")
            status.InnerText = "N"
            rps.AppendChild(status)

            Dim tribRps = doc.CreateElement("TributacaoRPS")
            tribRps.InnerText = "T"
            rps.AppendChild(tribRps)

            'Valores básicos (mantém o que já existe no ABRASF)
            AddValor(doc, rps, "ValorDeducoes", 0)
            AddValor(doc, rps, "ValorPIS", 0)
            AddValor(doc, rps, "ValorCOFINS", 0)
            AddValor(doc, rps, "ValorINSS", 0)
            AddValor(doc, rps, "ValorIR", 0)
            AddValor(doc, rps, "ValorCSLL", 0)

            'Serviço (tenta aproveitar CodTribMun/ProdutoCNAE; se não achar, mantém 0)
            Dim codigoServico As String = "0"
            Dim aliquota As Decimal = 0D
            'Try
            '    Dim trib = (From t In db.TributacaoDocs Where t.NumIntDoc = nf.NumIntTrib Select t).FirstOrDefault()
            '    If trib IsNot Nothing Then
            '        aliquota = CDec(If(trib.ISSQNPercentual.HasValue, trib.ISSQNPercentual.Value / 100.0, 0))
            '    End If
            'Catch
            'End Try

            Dim nCodServ = doc.CreateElement("CodigoServico")
            nCodServ.InnerText = codigoServico
            rps.AppendChild(nCodServ)

            Dim nAliq = doc.CreateElement("AliquotaServicos")
            nAliq.InnerText = aliquota.ToString("0.####", Globalization.CultureInfo.InvariantCulture)
            rps.AppendChild(nAliq)

            Dim issRet = doc.CreateElement("ISSRetido")
            issRet.InnerText = "false"
            rps.AppendChild(issRet)

            'Tomador
            If tom IsNot Nothing AndAlso Not String.IsNullOrEmpty(tom.CGC) Then
                Dim cpfcnpj = doc.CreateElement("CPFCNPJTomador")
                Dim docTom As String = LimparNumeros(tom.CGC)
                If docTom.Length <= 11 Then
                    Dim ncpf = doc.CreateElement("CPF")
                    ncpf.InnerText = docTom.PadLeft(11, "0"c)
                    cpfcnpj.AppendChild(ncpf)
                Else
                    Dim ncnpj = doc.CreateElement("CNPJ")
                    ncnpj.InnerText = docTom.PadLeft(14, "0"c)
                    cpfcnpj.AppendChild(ncnpj)
                End If
                rps.AppendChild(cpfcnpj)

                Dim raz = doc.CreateElement("RazaoSocialTomador")
                raz.InnerText = tom.Nome
                rps.AppendChild(raz)

                If tom.Endereco.HasValue Then
                    Dim endTom = (From e In db.Enderecos Where e.Codigo = tom.Endereco.Value Select e).FirstOrDefault()
                    If endTom IsNot Nothing Then
                        Dim cidTom = (From c In db.Cidades Where c.Descricao = endTom.Cidade Select c).FirstOrDefault()
                        Dim endNode = doc.CreateElement("EnderecoTomador")
                        AddText(doc, endNode, "TipoLogradouro", "RUA")
                        AddText(doc, endNode, "Logradouro", endTom.Logradouro)
                        AddText(doc, endNode, "NumeroEndereco", If(String.IsNullOrEmpty(endTom.Numero), "0", endTom.Numero))
                        AddText(doc, endNode, "ComplementoEndereco", endTom.Complemento)
                        AddText(doc, endNode, "Bairro", endTom.Bairro)
                        AddText(doc, endNode, "Cidade", If(cidTom IsNot Nothing, cidTom.CodIBGE.ToString(), "0"))
                        AddText(doc, endNode, "UF", endTom.SiglaEstado)
                        AddText(doc, endNode, "CEP", LimparNumeros(endTom.CEP))
                        rps.AppendChild(endNode)
                    End If
                End If
            End If

            Dim discr = doc.CreateElement("Discriminacao")
            discr.InnerText = If(String.IsNullOrEmpty(nf.MensagemNota), "", nf.MensagemNota)
            rps.AppendChild(discr)

            AddValor(doc, rps, "ValorFinalCobrado", If(nf.ValorTotal.HasValue, nf.ValorTotal.Value, 0))
            AddValor(doc, rps, "ValorIPI", 0)
            AddValor(doc, rps, "ExigibilidadeSuspensa", 0)
            AddValor(doc, rps, "PagamentoParceladoAntecipado", 0)

            'IBS/CBS (Reforma) - mínimo conforme sua regra
            Dim ibs = doc.CreateElement("IBSCBS")
            AddText(doc, ibs, "finNFSe", "0")
            AddText(doc, ibs, "indFinal", "0")
            AddText(doc, ibs, "cIndOp", "020101")
            AddText(doc, ibs, "indDest", "1")

            Dim valores = doc.CreateElement("valores")
            Dim trib = doc.CreateElement("trib")
            Dim g = doc.CreateElement("gIBSCBS")
            'Pelo exemplo, cst não aparece (mas o XSD v02.2 aceita/define em tipo). Mantemos cst + cClassTrib.
            AddText(doc, g, "cst", CST_IBSCBS)
            AddText(doc, g, "cClassTrib", CCLASSTRIB)
            trib.AppendChild(g)
            valores.AppendChild(trib)
            ibs.AppendChild(valores)

            rps.AppendChild(ibs)

            root.AppendChild(rps)
        Next

        Using sw As New StringWriter()
            Dim xw As XmlWriter = XmlWriter.Create(sw, New XmlWriterSettings() With {.Encoding = Encoding.UTF8, .OmitXmlDeclaration = True, .Indent = True})
            doc.Save(xw)
            xw.Flush()
            Return sw.ToString()
        End Using
    End Function

    Private Sub ProcessarRetorno(ByVal db As SGEDadosDataContext, ByVal iFilialEmpresa As Integer, ByVal lLote As Long, ByVal xmlResp As String)
        Dim doc As New XmlDocument()
        doc.PreserveWhitespace = False
        doc.LoadXml(xmlResp)

        Dim nsmgr As New XmlNamespaceManager(doc.NameTable)
        nsmgr.AddNamespace("nfe", "http://www.prefeitura.sp.gov.br/nfe")

        Dim sucesso = doc.SelectSingleNode("//nfe:RetornoEnvioLoteRPS/nfe:Cabecalho/nfe:Sucesso", nsmgr)
        If sucesso IsNot Nothing AndAlso sucesso.InnerText.Trim().ToLowerInvariant() <> "true" Then
            Throw New Exception("Retorno SP: Sucesso=false")
        End If

        Dim numLoteNode = doc.SelectSingleNode("//nfe:InformacoesLote/nfe:NumeroLote", nsmgr)
        Dim numeroLote As String = If(numLoteNode IsNot Nothing, numLoteNode.InnerText, lLote.ToString())

        Dim chaves = doc.SelectNodes("//nfe:ChaveNFeRPS", nsmgr)
        If chaves Is Nothing Then Return

        For Each chave As XmlNode In chaves
            Dim numRps = chave.SelectSingleNode("nfe:ChaveRPS/nfe:NumeroRPS", nsmgr)
            Dim numNfe = chave.SelectSingleNode("nfe:ChaveNFe/nfe:NumeroNFe", nsmgr)
            Dim codVer = chave.SelectSingleNode("nfe:ChaveNFe/nfe:CodigoVerificacao", nsmgr)

            If numRps Is Nothing OrElse numNfe Is Nothing Then Continue For

            'Grava em RPSWEBProt (mínimo): lote, numero, código verificação
            Dim prot As New RPSWEBProt()
            prot.Ambiente = 2 'homologação por padrão; ajuste via config se desejar
            prot.Data = Date.Today
            prot.Hora = CDbl(Date.Now.ToString("HHmmss"))
            prot.Lote = numeroLote
            prot.FilialEmpresa = CShort(iFilialEmpresa)
            prot.TipoRPS = 1
            prot.SerieRPS = GetAppSetting("NFSE_SP_SERIE", "1")
            prot.NumeroRPS = numRps.InnerText
            prot.Numero = numNfe.InnerText
            prot.CodigoVerificacao = If(codVer IsNot Nothing, codVer.InnerText, "")
            prot.Versao = "3.3.4"
            prot.DataEmissao = Date.Today
            prot.DataEmissaoRPS = Date.Today

            db.RPSWEBProts.InsertOnSubmit(prot)
            db.SubmitChanges()
        Next
    End Sub

    Private Sub GravarLogLote(ByVal db As SGEDadosDataContext, ByVal iFilialEmpresa As Integer, ByVal lLote As Long, ByVal numIntNF As Integer, ByVal status As String, ByVal texto As String)
        Try
            Dim log As New RPSWEBLoteLog()
            log.FilialEmpresa = CShort(iFilialEmpresa)
            log.Lote = CInt(lLote)
            log.Data = Date.Today
            log.Hora = CDbl(Date.Now.ToString("HHmmss"))
            log.status = status
            log.NumIntNF = numIntNF
            log.NumIntDoc = 0
            log.Ambiente = 2
            db.RPSWEBLoteLogs.InsertOnSubmit(log)
            db.SubmitChanges()
        Catch
        End Try
    End Sub

    Private Function EnviarSoap(ByVal url As String, ByVal metodo As String, ByVal xmlConteudo As String) As String
        'Envelope SOAP 1.1 simples
        Dim soap As String = "<?xml version=""1.0"" encoding=""utf-8""?>" &
            "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &
            "<soap:Body>" &
            "<" & metodo & " xmlns=""http://www.prefeitura.sp.gov.br/nfe"">" &
            "<PedidoEnvioLoteRPS><![CDATA[" & xmlConteudo & "]]></PedidoEnvioLoteRPS>" &
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
                Dim soapResp As String = sr.ReadToEnd()
                'extrai o conteúdo XML do retorno, se vier encapsulado
                Return ExtrairPrimeiroXml(soapResp)
            End Using
        End Using
    End Function

    Private Function ExtrairPrimeiroXml(ByVal soapResp As String) As String
        'Tenta localizar <RetornoEnvioLoteRPS ...> dentro da resposta
        Dim i As Integer = soapResp.IndexOf("<RetornoEnvioLoteRPS", StringComparison.OrdinalIgnoreCase)
        If i >= 0 Then
            Dim j As Integer = soapResp.IndexOf("</RetornoEnvioLoteRPS>", StringComparison.OrdinalIgnoreCase)
            If j > i Then
                j += "</RetornoEnvioLoteRPS>".Length
                Return soapResp.Substring(i, j - i)
            End If
        End If
        'Fallback: retorna tudo
        Return soapResp
    End Function

    Private Sub AddText(ByVal doc As XmlDocument, ByVal parent As XmlElement, ByVal name As String, ByVal value As String)
        Dim n = doc.CreateElement(name)
        n.InnerText = If(value, "")
        parent.AppendChild(n)
    End Sub

    Private Sub AddValor(ByVal doc As XmlDocument, ByVal parent As XmlElement, ByVal name As String, ByVal value As Double)
        Dim n = doc.CreateElement(name)
        n.InnerText = value.ToString("0.##", Globalization.CultureInfo.InvariantCulture)
        parent.AppendChild(n)
    End Sub

    Private Function LimparNumeros(ByVal s As String) As String
        If String.IsNullOrEmpty(s) Then Return ""
        Dim sb As New StringBuilder(s.Length)
        For Each ch As Char In s
            If Char.IsDigit(ch) Then sb.Append(ch)
        Next
        Return sb.ToString()
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

    '==================== Assinatura Paulistana (campos) ====================

    Private Function MontarStringAssinatura(ByVal imPrestador As String,
                                           ByVal serie As String,
                                           ByVal numeroRps As Integer,
                                           ByVal dataEmissao As Nullable(Of Date),
                                           ByVal valorTotal As Nullable(Of Double)) As String
        'Implementação conservadora (ajuste conforme seu manual):
        'IM(8) + Serie(5) + NumeroRPS(12) + Data(8) + Valor(15)
        Dim im = (If(imPrestador, "")).PadLeft(8, "0"c)
        Dim se = (If(serie, "")).PadRight(5, " "c).Substring(0, 5)
        Dim num = numeroRps.ToString().PadLeft(12, "0"c)
        Dim dt = If(dataEmissao.HasValue, dataEmissao.Value.ToString("yyyyMMdd"), Date.Today.ToString("yyyyMMdd"))
        Dim vl = CLng(Math.Round(If(valorTotal.HasValue, valorTotal.Value, 0) * 100, MidpointRounding.AwayFromZero)).ToString().PadLeft(15, "0"c)
        Return im & se & num & dt & vl
    End Function

    Private Function AssinarCamposBase64(ByVal dados As String, ByVal cert As X509Certificate2) As String
        Dim rsa As RSA = cert.GetRSAPrivateKey()
        Dim bytes = Encoding.UTF8.GetBytes(dados)
        Dim sig = rsa.SignData(bytes, HashAlgorithmName.SHA1, RSASignaturePadding.Pkcs1)
        Return Convert.ToBase64String(sig)
    End Function

End Class
