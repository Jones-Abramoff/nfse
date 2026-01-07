Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Net
Imports System.Net.Http
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Text.Json
Imports System.Xml

Public Class NfseNacionalEmissor

    Private Const NS_DPS As String = "http://www.nfse.gov.br/dps"
    Private Const URL_HOMO As String = "http://sefin.producaorestrita.nfse.gov.br/SefinNacional/nfse"

    ' Ajuste conforme sua política (verProc etc.)
    Private Const VER_PROC As String = "SGE-NFSE-NACIONAL"

    Public Function EmitirNfse(empresa As String, numIntDoc As Long, filialEmpresa As Integer) As NfseEmissaoResult
        Dim repo As New DbRepository(empresa)

        Dim filial = repo.GetFilial(filialEmpresa)
        Dim endFilial = repo.GetEndereco(filial.EnderecoCod)
        Dim nf = repo.GetNFiscal(numIntDoc)
        Dim itens = repo.GetItens(numIntDoc)
        Dim tomador = repo.GetTomador(numIntDoc)

        ' Certificado (A1 no repositório do Windows) - você já tem uma rotina BuscaNome, troque aqui se quiser.
        Dim cert = CertUtil.BuscaPorSubjectOuThumbprint(filial.CertificadoA1A3)
        If cert Is Nothing Then Throw New Exception("Certificado não encontrado: " & filial.CertificadoA1A3)

        ' Monta XML DPS (na ORDEM do XSD)
        Dim idInf As String = "DPS" & endFilial.CidadeCodMun.ToString() & DateTime.Now.ToString("yyyyMMddHHmmssfff")
        Dim xmlDps = MontaDpsXml(idInf, filial, endFilial, nf, itens, tomador, filial.RPSAmbiente)

        ' Assina infDPS
        Dim xmlAssinado = XmlUtil.AssinarDps(xmlDps, idInf, cert)

        ' GZip+Base64
        Dim payloadObj = New With {.dpsXmlGZipB64 = XmlUtil.GZipToBase64(xmlAssinado)}
        Dim payload = JsonSerializer.Serialize(payloadObj)

        ' POST
        Dim httpRes = PostJson(URL_HOMO, payload)

        Dim result As New NfseEmissaoResult()
        If httpRes.StatusCode <> HttpStatusCode.OK Then
            result.Sucesso = False
            result.Erro = $"HTTP {(CInt(httpRes.StatusCode))} - {httpRes.Body}"
            Return result
        End If

        ' parse JSON resposta
        Dim doc = JsonDocument.Parse(httpRes.Body)
        Dim root = doc.RootElement

        result.TipoAmbiente = root.GetProperty("tipoAmbiente").GetInt32()
        result.VersaoAplicativo = root.GetProperty("versaoAplicativo").GetString()
        result.DataHoraProcessamento = DateTime.Parse(root.GetProperty("dataHoraProcessamento").GetString(), Nothing, DateTimeStyles.RoundtripKind)
        result.IdDps = root.GetProperty("idDps").GetString()
        result.ChaveAcesso = root.GetProperty("chaveAcesso").GetString()

        Dim nfseB64 = root.GetProperty("nfseXmlGZipB64").GetString()
        If Not String.IsNullOrWhiteSpace(nfseB64) Then
            result.NfseXml = XmlUtil.Base64GunzipToString(nfseB64)
        End If

        If root.TryGetProperty("alertas", Nothing) Then
            For Each a In root.GetProperty("alertas").EnumerateArray()
                result.Alertas.Add(New NfseAlerta With {
                    .mensagem = a.GetProperty("mensagem").GetString(),
                    .codigo = a.GetProperty("codigo").GetString(),
                    .descricao = a.GetProperty("descricao").GetString(),
                    .complemento = a.GetProperty("complemento").GetString()
                })
            Next
        End If

        result.Sucesso = True

        ' Grava no BD: (ajuste Protocolo/CódigoVerif conforme retorno real do XML)
        repo.GravaEmissaoOk(numIntDoc, filialEmpresa, filial.RPSAmbiente, result.IdDps, result.ChaveAcesso, result.NfseXml, protocolo:="", codigoVerificacao:="")

        Return result
    End Function

    Private Function MontaDpsXml(idInfDps As String,
                                filial As FilialEmpresaRow,
                                endFilial As EnderecoRow,
                                nf As NFiscalRow,
                                itens As List(Of ItensNFiscalRow),
                                tomador As TomadorRow,
                                tpAmb As Integer) As String

        Dim settings As New XmlWriterSettings() With {
            .Encoding = Encoding.UTF8,
            .Indent = True,
            .OmitXmlDeclaration = False
        }

        Using sw As New System.IO.StringWriter()
            Using xw = XmlWriter.Create(sw, settings)
                xw.WriteStartDocument()
                xw.WriteStartElement("DPS", NS_DPS)
                xw.WriteAttributeString("versao", "1.00")

                xw.WriteStartElement("infDPS")
                xw.WriteAttributeString("Id", idInfDps)

                ' -------- ORDEM XSD (TCInfDPS) --------
                xw.WriteElementString("tpAmb", tpAmb.ToString()) ' 2=Homologação
                xw.WriteElementString("dhEmi", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz"))
                xw.WriteElementString("verProc", VER_PROC)

                ' dCompet: use a competência (mês) - aqui por DataEmissao
                Dim comp = If(nf.DataEmissao.HasValue, nf.DataEmissao.Value, DateTime.Today)
                xw.WriteElementString("dCompet", New DateTime(comp.Year, comp.Month, 1).ToString("yyyy-MM-dd"))

                xw.WriteElementString("tpEmit", "1") ' 1=prestador (ajuste se XSD diferir)
                xw.WriteElementString("cLocEmi", endFilial.CidadeCodMun.ToString())

                ' subst (opcional) - não preenche
                ' xw.WriteStartElement("subst") ...

                ' prest (obrigatório)
                WritePrestador(xw, filial, endFilial)

                ' toma (obrigatório no seu fluxo; no XSD é obrigatório)
                WriteTomador(xw, tomador)

                ' serv (obrigatório)
                WriteServico(xw, itens)

                ' valores (obrigatório)
                WriteValores(xw, nf)

                xw.WriteEndElement() ' infDPS
                xw.WriteEndElement() ' DPS
                xw.WriteEndDocument()
            End Using
            Return sw.ToString()
        End Using
    End Function

    Private Sub WritePrestador(xw As XmlWriter, filial As FilialEmpresaRow, endFilial As EnderecoRow)
        xw.WriteStartElement("prest")
        ' choice: CNPJ/CPF/NIF/cNaoNIF
        Dim cgc = SoNumeros(filial.CGC)
        If cgc.Length = 14 Then
            xw.WriteElementString("CNPJ", cgc)
        ElseIf cgc.Length = 11 Then
            xw.WriteElementString("CPF", cgc)
        Else
            ' último recurso (não recomendado). Ajuste conforme regra
            xw.WriteElementString("cNaoNIF", "1")
        End If

        xw.WriteElementString("IM", SoNumeros(filial.InscricaoMunicipal))
        xw.WriteElementString("xNome", Trunc(filial.RazaoSocial, 115))

        ' endereço (TCEndereco: endNac|endExt, xLgr, nro, xCpl, xBairro)
        xw.WriteStartElement("end")
        xw.WriteStartElement("endNac")
        xw.WriteElementString("cMun", endFilial.CidadeCodMun.ToString())
        xw.WriteElementString("CEP", PadLeftZeros(SoNumeros(endFilial.CEP), 8))
        xw.WriteEndElement() ' endNac
        xw.WriteElementString("xLgr", Trunc(endFilial.Logradouro, 125))
        xw.WriteElementString("nro", Trunc(endFilial.Numero, 10))
        If Not String.IsNullOrWhiteSpace(endFilial.Complemento) Then xw.WriteElementString("xCpl", Trunc(endFilial.Complemento, 60))
        xw.WriteElementString("xBairro", Trunc(endFilial.Bairro, 60))
        xw.WriteEndElement() ' end

        If Not String.IsNullOrWhiteSpace(filial.Telefone) Then xw.WriteElementString("fone", PadLeftZeros(SoNumeros(filial.Telefone), 11))
        If Not String.IsNullOrWhiteSpace(filial.Email) Then xw.WriteElementString("email", Trunc(filial.Email, 80))

        ' regTrib (opcional) - não preenche aqui (Simples, MEI, etc.)
        xw.WriteEndElement() ' prest
    End Sub

    Private Sub WriteTomador(xw As XmlWriter, t As TomadorRow)
        xw.WriteStartElement("toma")

        Dim doc = SoNumeros(t.CpfCnpj)
        If doc.Length = 14 Then
            xw.WriteElementString("CNPJ", doc)
        ElseIf doc.Length = 11 Then
            xw.WriteElementString("CPF", doc)
        Else
            ' sem CPF/CNPJ => usa cNaoNIF (XSD exige um item do choice)
            xw.WriteElementString("cNaoNIF", t.CNaoNIF.ToString())
        End If

        xw.WriteElementString("xNome", Trunc(t.RazaoSocial, 115))

        xw.WriteStartElement("end")
        xw.WriteStartElement("endNac")
        xw.WriteElementString("cMun", t.Endereco.CidadeCodMun.ToString())
        xw.WriteElementString("CEP", PadLeftZeros(SoNumeros(t.Endereco.CEP), 8))
        xw.WriteEndElement() ' endNac
        xw.WriteElementString("xLgr", Trunc(t.Endereco.Logradouro, 125))
        xw.WriteElementString("nro", Trunc(t.Endereco.Numero, 10))
        If Not String.IsNullOrWhiteSpace(t.Endereco.Complemento) Then xw.WriteElementString("xCpl", Trunc(t.Endereco.Complemento, 60))
        xw.WriteElementString("xBairro", Trunc(t.Endereco.Bairro, 60))
        xw.WriteEndElement() ' end

        If Not String.IsNullOrWhiteSpace(t.Telefone) Then xw.WriteElementString("fone", PadLeftZeros(SoNumeros(t.Telefone), 11))
        If Not String.IsNullOrWhiteSpace(t.Email) Then xw.WriteElementString("email", Trunc(t.Email, 80))

        xw.WriteEndElement() ' toma
    End Sub

    Private Sub WriteServico(xw As XmlWriter, itens As List(Of ItensNFiscalRow))
        xw.WriteStartElement("serv")

        ' locPrest (opcional)
        ' cServ (obrigatório): cTribNac, xDescServ, cNBS (opcional), cTribMun (obrigatório)
        xw.WriteStartElement("cServ")
        ' aqui você deve derivar do seu cadastro ISSQN (lista LC116). Ex: "7.02" => "070200"
        Dim cTribNac As String = "010101" ' TODO: mapear corretamente!
        xw.WriteElementString("cTribNac", cTribNac)

        Dim discr As String = MontaDiscriminacaoAbrasfStyle(itens)
        xw.WriteElementString("xDescServ", Trunc(discr, 2000))

        ' cTribMun (3 dígitos). Se você tiver "CodServNFe" no cadastro, formate aqui:
        xw.WriteElementString("cTribMun", "000")

        xw.WriteEndElement() ' cServ

        ' info compl (opcional)
        ' trib (TCInfoTributacao) opcional - se seu cenário exige, inclua aqui.

        xw.WriteEndElement() ' serv
    End Sub

    Private Sub WriteValores(xw As XmlWriter, nf As NFiscalRow)
        xw.WriteStartElement("valores")
        ' TCInfoValores: vServPrest, vDescCondIncond, vDescIncond, vDedRed, trib? etc.
        Dim vServ As Double = If(nf.ValorServicos.HasValue, nf.ValorServicos.Value, 0R)
        Dim vDesc As Double = If(nf.ValorDesconto.HasValue, nf.ValorDesconto.Value, 0R)
        Dim vDed As Double = If(nf.ValorDeducoes.HasValue, nf.ValorDeducoes.Value, 0R)

        xw.WriteElementString("vServPrest", FormatValor(vServ))
        xw.WriteElementString("vDescCondIncond", FormatValor(vDesc))
        xw.WriteElementString("vDescIncond", FormatValor(vDesc))
        xw.WriteElementString("vDedRed", FormatValor(vDed))

        ' trib (opcional). Se você quiser, inclua TCTribMunicipal aqui.
        ' xw.WriteStartElement("trib") ...
        xw.WriteEndElement() ' valores
    End Sub

    Private Shared Function MontaDiscriminacaoAbrasfStyle(itens As List(Of ItensNFiscalRow)) As String
        Dim sb As New StringBuilder()
        For Each it In itens
            Dim linha = $"{it.DescricaoItem} Quant: {it.Quantidade.ToString("###.###,###", CultureInfo.GetCultureInfo("pt-BR"))} P.Unit: {it.PrecoUnitario.ToString("0.00", CultureInfo.GetCultureInfo("pt-BR"))} Total: {(it.Quantidade * it.PrecoUnitario - it.ValorDesconto).ToString("0.00", CultureInfo.GetCultureInfo("pt-BR"))}"
            If sb.Length = 0 Then
                sb.Append(linha)
            Else
                sb.Append("|").Append(linha)
            End If
        Next
        Return sb.ToString()
    End Function

    Private Shared Function PostJson(url As String, json As String) As (StatusCode As HttpStatusCode, Body As String)
        Using handler As New HttpClientHandler()
            ' Se precisar de proxy/TLS custom, ajuste aqui.
            Using client As New HttpClient(handler)
                client.Timeout = TimeSpan.FromSeconds(120)
                Using content As New StringContent(json, Encoding.UTF8, "application/json")
                    Dim resp = client.PostAsync(url, content).GetAwaiter().GetResult()
                    Dim body = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                    Return (resp.StatusCode, body)
                End Using
            End Using
        End Using
    End Function

    Private Shared Function SoNumeros(s As String) As String
        If String.IsNullOrEmpty(s) Then Return ""
        Dim sb As New StringBuilder()
        For Each ch In s
            If Char.IsDigit(ch) Then sb.Append(ch)
        Next
        Return sb.ToString()
    End Function

    Private Shared Function Trunc(s As String, maxLen As Integer) As String
        If s Is Nothing Then Return ""
        If s.Length <= maxLen Then Return s
        Return s.Substring(0, maxLen)
    End Function

    Private Shared Function PadLeftZeros(s As String, len As Integer) As String
        If s Is Nothing Then s = ""
        If s.Length >= len Then Return s
        Return New String("0"c, len - s.Length) & s
    End Function

    Private Shared Function FormatValor(v As Double) As String
        ' XSD geralmente exige decimal com ponto.
        Return v.ToString("0.00", CultureInfo.InvariantCulture)
    End Function
End Class
