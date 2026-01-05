
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
Imports System.Math



Public Class ClassConsultaLoteNFSE



    Public Const RPS_AMBIENTE_HOMOLOGACAO As Integer = 2
    Public Const RPS_AMBIENTE_PRODUCAO As Integer = 1

    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Long
    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

    Const TIPODOC_TRIB_NF = 0

    Public Sub Consulta_Lote_NFSE(ByVal sEmpresa As String, ByVal lLote As Long, ByVal iFilialEmpresa As Integer)

#If Not TATUI2 Then

#If Not RJ Then
        Dim a4 As cabecalho = New cabecalho
#End If

        Dim iIndice2 As Integer
        Dim iIndiceRPS As Integer
        Dim iAchou As Integer

        Dim XMLStreamRPS As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamRetRPS As MemoryStream = New MemoryStream(10000)
        Dim XMLStream1 As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamCabec As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamDados As MemoryStream = New MemoryStream(10000)

        Dim xm11 As Byte()
        Dim xm22 As Byte()

        Dim db1 As SGEDadosDataContext = New SGEDadosDataContext
        Dim db2 As SGEDadosDataContext = New SGEDadosDataContext

        Dim dic As SGEDicDataContext = New SGEDicDataContext

        Dim resFatConfig As IEnumerable(Of FATConfig)
        Dim resRPSWEBLote As IEnumerable(Of RPSWEBLote)
        Dim resRPSWEBRetEnvi As IEnumerable(Of RPSWEBRetEnvi)
        Dim resNFiscal As IEnumerable(Of NFiscal)
        Dim resRPSWEBProt As IEnumerable(Of RPSWEBProt)
        Dim resFilialEmpresa As IEnumerable(Of FiliaisEmpresa)
        Dim resEndereco As IEnumerable(Of Endereco)

        Dim objNFiscal As NFiscal
        Dim objRPSWEBRetEnvi As RPSWEBRetEnvi
        Dim objRPSWEBLote As RPSWEBLote
        Dim objFatConfig As FATConfig
        Dim objFilialEmpresa As FiliaisEmpresa


        Dim colNFiscal As Collection = New Collection

        Dim XMLString1 As String
        Dim XMLString2 As String
        Dim XMLStringCabec As String
        Dim iResult As Integer

        Dim cert As X509Certificate2 = New X509Certificate2
        Dim certificado As Certificado = New Certificado

        Dim homologacao As New NotaCariocaHomologacao.NfseSoapClient
        Dim producao As New NotaCariocaProducao.NfseSoapClient

        Dim homosalvadorconslote As New br.gov.ba.salvador.sefaz.nfsehml3.ConsultaLoteRPS
        Dim prodsalvadorconslote As New br.gov.ba.salvador.sefaz.nfse3.ConsultaLoteRPS

        Dim homosalvadorconsRPS As New br.gov.ba.salvador.sefaz.nfsehml2.ConsultaNfseRPS
        Dim prodsalvadorconsRPS As New br.gov.ba.salvador.sefaz.nfse2.ConsultaNfseRPS

        Dim lNumIntNF As Long

        Dim objValidaXML As ClassValidaXML = New ClassValidaXML

        Dim odbc As OdbcConnection = New OdbcConnection
        Dim sString As String

        Dim iInicio As Integer
        Dim iFim As Integer



        Dim iTamanho As Integer
        Dim sRetorno As String
        Dim sFile As String
        Dim sDir As String
        Dim sDir1 As String
        Dim iPos As Integer

        Dim sSerie As String

        Dim sCertificado As String
        Dim iRPSAmbiente As Integer



        Dim xRet As Byte()
        Dim xRetRPS As Byte()

        Dim lErro As Long



        Dim objListaMsgRetorno As ListaMensagemRetorno
#If ABRASF Then
        Dim objListaMsgRetornoLote As ListaMensagemRetornoLote
        Dim objMsgRetornoLote As New tcMensagemRetornoLote
#End If
        Dim objMsgRetorno As New tcMensagemRetorno
        Dim objConsultarLoteRpsRespostaListaNFse As ConsultarLoteRpsRespostaListaNfse
        Dim objCompNfse As tcCompNfse
        Dim sProtocolo As String
        Dim sLote As String

        Dim lRPSWEBConsLoteNumIntDoc As Long
        Dim lRPSWEBProtNumIntDoc As Long

        Dim colNumIntNFiscal As Collection = New Collection
        Dim lNumIntNFiscal As Long
        Dim lRPSWEBLoteLogNumIntDoc As Long

        Dim sEmailPrestador As String
        Dim sTelefonePrestador As String
        Dim sEmailTomador As String
        Dim sTelefoneTomador As String
        Dim sAux As String
        Dim sAuxRPS As String
        Dim objEndFilial As Endereco

        Dim objConsNfseRpsEnvio As ConsultarNfseRpsEnvio = New ConsultarNfseRpsEnvio

        Dim homologacaobh = New GetWebRequest_bhHomologacao
        Dim producaobh = New GetWebRequest_bhProducao

        Dim homologacaoginfes = New br.com.ginfes.homologacao.ServiceGinfesImplService
        Dim producaoginfes = New br.com.ginfes.producao.ServiceGinfesImplService

        Dim homologacaosistema4rconsulta = New br.com.sistemas4r.abrasf.nfsehml2.ConsultarNfsePorRps
        Dim producaosisstema4rconsulta = New br.com.sistemas4r.abrasf.nfse2.ConsultarNfsePorRps


        Dim sCNPJCPFTomador As String
        Dim sInscricaoMunTomador As String

        Dim sRazaoSocialTomador As String
        Dim sCodigoVerificacao As String
        Dim dtDataEmissao As Date
        Dim dtDataEmissaoRPS As Date
        Dim sCompetencia As String
        Dim sbTipo As SByte
        Dim sNumeroRPS As String
        Dim sNumero As String
        Dim sbNaturezaOperacao As SByte
        Dim sbOptanteSimplesNacional As SByte
        Dim sUFGerador As String
        Dim sCNPJPrestador As String
        Dim sRazaoSocialPrestador As String
        Dim sDiscriminacao As String
        Dim sItemListaServico As String
        Dim sInscricaoMunicipalPrestador As String
        Dim sBairroPrestador As String
        Dim lCepPrestador As Long
        Dim lCodMunPrestador As Long
        Dim sComplementoPrestador As String
        Dim sEnderecoPrestador As String
        Dim sNumeroPrestador As String
        Dim sUFPrestador As String
        Dim sNomeFantasiaPrestador As String
        Dim sbRegimeEspecialTributacao As SByte
        Dim lCodigoCnae As Long
        Dim lCodigoMunicipioServico As Long
        Dim sCodigoTributacaoMunicipio As String
        Dim sBairroTomador As String
        Dim lCepTomador As Long
        Dim lCodigoMunicipioTomador As Long
        Dim sComplementoTomador As String
        Dim sEnderecoTomador As String
        Dim sNumeroTomador As String
        Dim sUFTomador As String
        Dim lCodigoMunicipioGerador As Long

        Dim dAliquota As Double
        Dim dBaseCalculo As Double
        Dim dDescontoCondicionado As Double
        Dim dDescontoIncondicionado As Double
        Dim dISSRetido As Double
        Dim dOutrasRetencoes As Double
        Dim dValorCofins As Double
        Dim dValorCsll As Double
        Dim dValorDeducoes As Double
        Dim dValorInss As Double
        Dim dValorIr As Double
        Dim dValorIss As Double
        Dim dValorIssRetido As Double
        Dim dValorLiquidoNfse As Double
        Dim dValorPis As Double
        Dim dValorServicos As Double
        Dim dValorCredito As Double
        Dim XMLStringConsultarLoteRPSResposta As String
        Dim XMLStringConsultarNfseRpsResposta As String



        Dim sErro As String
        Dim sMsg1 As String
        Dim iDebug As Integer

        sErro = ""
        sMsg1 = ""

        sSerie = ""


        Try

            iTamanho = 255
            sRetorno = StrDup(iTamanho, Chr(0))

            iDebug = 0

            Call GetPrivateProfileString("Geral", "Debug", 0, sRetorno, iTamanho, "Adm100.ini")

            If IsNumeric(sRetorno) Then
                iDebug = CInt(sRetorno)
            End If


            lNumIntNF = 0

            '********** pega o diretorio do log para colocar os arquivos xml *************
            iTamanho = 255
            sRetorno = StrDup(iTamanho, Chr(0))

            Call GetPrivateProfileString("Geral", "ArqLog", -1, sRetorno, iTamanho, "Adm100.ini")

            iPos = InStr(sRetorno, Chr(0))

            sRetorno = Mid(sRetorno, 1, iPos - 1)

            sFile = Dir(sRetorno)

            iPos = InStr(sRetorno, sFile)

            sDir = Mid(sRetorno, 1, iPos - 1)


            '********** pega o diretorio dos executaveis para ler os arquivos xsd *************
            iTamanho = 255
            sRetorno = StrDup(iTamanho, Chr(0))

            Call GetPrivateProfileString("Forprint", "DirBin", -1, sRetorno, iTamanho, "ADM100.INI")

            iPos = InStr(sRetorno, Chr(0))

            sDir1 = Mid(sRetorno, 1, iPos - 1)

            sErro = "101"
            sMsg1 = "vai abrir SGEDados"

            If iDebug = 1 Then MsgBox("101")

            '***** coloca a string de conexao apontando para o SGEDados em questao *****
            odbc.ConnectionString = "DSN=SGEDados" & sEmpresa & ";UID=sa;PWD=SAPWD"
            '            odbc.ConnectionString = "DSN=SGEDados" & sEmpresa & ";UID=sa;PWD=SAPWD"

            odbc.Open()

            sString = db1.Connection.ConnectionString

            iInicio = InStr(sString, "Data Source=")

            iFim = InStr(iInicio, sString, ";")

            sString = Mid(sString, 1, iInicio + 11) & odbc.DataSource & Mid(sString, iFim)

            iInicio = InStr(sString, "Initial Catalog=")

            iFim = InStr(iInicio, sString, ";")

            sString = Mid(sString, 1, iInicio + 15) & odbc.Database & Mid(sString, iFim)
            sString = sString & ";TIMEOUT=30000"

            db1.Connection.ConnectionString = sString

            db2.Connection.ConnectionString = sString

            odbc.Close()

            sErro = "102"
            sMsg1 = "vai abrir SGEDic"

            If iDebug = 1 Then MsgBox("102")

            '***** coloca a string de conexao apontando para o SGEDic *****
            odbc.ConnectionString = "DSN=SGEDic" & ";UID=admin;PWD=cacareco"

            odbc.Open()

            sString = dic.Connection.ConnectionString

            iInicio = InStr(sString, "Data Source=")

            iFim = InStr(iInicio, sString, ";")

            sString = Mid(sString, 1, iInicio + 11) & odbc.DataSource & Mid(sString, iFim)

            iInicio = InStr(sString, "Initial Catalog=")

            iFim = InStr(iInicio, sString, ";")

            sString = Mid(sString, 1, iInicio + 15) & odbc.Database & Mid(sString, iFim)
            sString = sString & ";TIMEOUT=30000"

            dic.Connection.ConnectionString = sString

            odbc.Close()

            db1.Connection.Open()
            db2.Connection.Open()

            dic.Connection.Open()

            db2.Transaction = db2.Connection.BeginTransaction()
            db1.Transaction = db1.Connection.BeginTransaction()


            '  seleciona certificado do repositório MY do windows
            '

            sErro = "103"
            sMsg1 = "vai ler FiliaisEmpresa"


            resFilialEmpresa = db1.ExecuteQuery(Of FiliaisEmpresa) _
            ("SELECT * FROM FiliaisEmpresa WHERE FilialEmpresa = {0} ", iFilialEmpresa)

            For Each objFilialEmpresa In resFilialEmpresa
                sCertificado = objFilialEmpresa.CertificadoA1A3
                iRPSAmbiente = objFilialEmpresa.RPSAmbiente
                Exit For
            Next

            sErro = "104"
            sMsg1 = "vai ler o certificado"

            cert = certificado.BuscaNome(sCertificado)

            sErro = "105"
            sMsg1 = "vai ler o Endereco"

            If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

            resEndereco = db1.ExecuteQuery(Of Endereco) _
            ("SELECT * FROM Enderecos WHERE Codigo = {0}", objFilialEmpresa.Endereco)

            objEndFilial = resEndereco(0)

#If Not RJ Then
            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then

                a4.versao = "1.00"
                a4.versaoDados = "1.00"

            End If

            If UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then
                '                If UCase(objEndFilial.Cidade) = "TATUÍ" Or SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then

                a4.versao = "3"
                a4.versaoDados = "3"

            End If

            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Or UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then
                '                If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Or UCase(objEndFilial.Cidade) = "TATUÍ" Or UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then
                Dim mySerializercabec As New XmlSerializer(GetType(cabecalho))

                XMLStreamCabec = New MemoryStream(10000)

                mySerializercabec.Serialize(XMLStreamCabec, a4)

                Dim doccabec As XmlDocument = New XmlDocument
                XMLStreamCabec.Position = 0
                'doccabec.Load(XMLStreamCabec)

                'doccabec.Save(sDir & "XmlRPSCabec.xml")

                'If iDebug = 1 Then MsgBox("14")

                'lErro = objValidaXML.validaXML(sDir & "XmlRPScabec.xml", sDir1 & "nfse1.xsd", lLote, lNumIntNF, db2, iFilialEmpresa)
                'If lErro = 1 Then


                '    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6} , {7})", _
                '    lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "ERRO - Encerrado o envio deste lote", 0)

                '    lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1


                '    Form1.Msg.Items.Add("ERRO - o envio do lote " & CStr(lLote) & " foi encerrado por erro.")

                '    Application.DoEvents()

                '    Exit Try
                'End If

                Dim xmcabec As Byte()

                xmcabec = XMLStreamCabec.ToArray

                XMLStringCabec = System.Text.Encoding.UTF8.GetString(xmcabec)

                XMLStringCabec = Mid(XMLStringCabec, 1, 19) & " encoding=""utf-8"" " & Mid(XMLStringCabec, 20)

            End If
#End If

            sErro = "106"
            sMsg1 = "vai ler FatConfig - NUM_INT_PROX_RPSWEBLOTELOG"

            resFatConfig = db2.ExecuteQuery(Of FATConfig) _
            ("SELECT * FROM FatConfig WHERE Codigo = {0} ", "NUM_INT_PROX_RPSWEBLOTELOG")

            objFatConfig = resFatConfig(0)

            lRPSWEBLoteLogNumIntDoc = CLng(objFatConfig.Conteudo)

            sErro = "107"
            sMsg1 = "vai ler FatConfig - NUM_INT_PROX_RPSWEBCONSLOTE"

            If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

            resFatConfig = db2.ExecuteQuery(Of FATConfig) _
            ("SELECT * FROM FatConfig WHERE Codigo = {0} ", "NUM_INT_PROX_RPSWEBCONSLOTE")

            objFatConfig = resFatConfig(0)

            lRPSWEBConsLoteNumIntDoc = CLng(objFatConfig.Conteudo)

            sErro = "108"
            sMsg1 = "vai ler FatConfig - NUM_INT_PROX_RPSWEBPROT"

            resFatConfig = db2.ExecuteQuery(Of FATConfig) _
            ("SELECT * FROM FatConfig WHERE Codigo = {0} ", "NUM_INT_PROX_RPSWEBPROT")

            objFatConfig = resFatConfig(0)

            lRPSWEBProtNumIntDoc = CLng(objFatConfig.Conteudo)

            sErro = "109"
            sMsg1 = "vai ler RPSWEBRetEnvi"

            If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

            '******** pega o ultimo retorno de envio do lote em questao *************
            resRPSWEBRetEnvi = db2.ExecuteQuery(Of RPSWEBRetEnvi) _
            ("SELECT * FROM RPSWEBRetEnvi WHERE Lote = {0} AND Ambiente = {1} AND Len(Protocolo) > 0 ORDER BY data DESC, hora DESC ", lLote, iRPSAmbiente)

            If resRPSWEBRetEnvi.Count = 0 Then

                sErro = "110"
                sMsg1 = "vai gravar RPSWEBLoteLog"

                If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
                lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "Não foi encontrado protocolo de envio do lote", 0)

                lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                Form1.Msg.Items.Add("Não foi encontrado protocolo de envio do lote.")

                If Form1.Msg.Items.Count - 15 < 1 Then
                    Form1.Msg.TopIndex = 1
                Else
                    Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                End If


                Application.DoEvents()

                Exit Try

            End If

            sErro = "111"
            sMsg1 = "vai ler RPSWEBRetEnvi"

            If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

            '******** pega o ultimo retorno de envio do lote em questao *************
            resRPSWEBRetEnvi = db2.ExecuteQuery(Of RPSWEBRetEnvi) _
            ("SELECT * FROM RPSWEBRetEnvi WHERE Lote = {0} AND Ambiente = {1} AND Len(Protocolo) > 0 ORDER BY data DESC, hora DESC ", lLote, iRPSAmbiente)

            objRPSWEBRetEnvi = resRPSWEBRetEnvi(0)

            Dim objConsLoteRpsEnvio As ConsultarLoteRpsEnvio = New ConsultarLoteRpsEnvio

            objConsLoteRpsEnvio.Prestador = New tcIdentificacaoPrestador

#If ABRASF2 Then
            If Len(objFilialEmpresa.CGC) = 14 Then
                objConsLoteRpsEnvio.Prestador.CpfCnpj = New tcCpfCnpj
                objConsLoteRpsEnvio.Prestador.CpfCnpj.ItemElementName = ItemChoiceType.Cnpj
                objConsLoteRpsEnvio.Prestador.CpfCnpj.Item = objFilialEmpresa.CGC
            ElseIf Len(objFilialEmpresa.CGC) = 11 Then
                objConsLoteRpsEnvio.Prestador.CpfCnpj = New tcCpfCnpj
                objConsLoteRpsEnvio.Prestador.CpfCnpj.ItemElementName = ItemChoiceType.Cpf
                objConsLoteRpsEnvio.Prestador.CpfCnpj.Item = objFilialEmpresa.CGC
            End If

#Else
            objConsLoteRpsEnvio.Prestador.Cnpj = objFilialEmpresa.CGC
#End If


            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then
                Call ADM.Formata_String_AlfaNumerico(objFilialEmpresa.InscricaoMunicipal, objConsLoteRpsEnvio.Prestador.InscricaoMunicipal)
            Else
                Call ADM.Formata_String_Numero(objFilialEmpresa.InscricaoMunicipal, objConsLoteRpsEnvio.Prestador.InscricaoMunicipal)
            End If


            '            Call Formata_String_Numero(objFilialEmpresa.InscricaoMunicipal, objConsLoteRpsEnvio.Prestador.InscricaoMunicipal)

            objConsLoteRpsEnvio.Protocolo = objRPSWEBRetEnvi.Protocolo

            Dim mySerializerx2 As New XmlSerializer(GetType(ConsultarLoteRpsEnvio))

            XMLStream1 = New MemoryStream(10000)
            mySerializerx2.Serialize(XMLStream1, objConsLoteRpsEnvio)

            xm11 = XMLStream1.ToArray

            XMLString1 = System.Text.Encoding.UTF8.GetString(xm11)

            XMLString1 = Mid(XMLString1, 1, 19) & " encoding=""utf-8"" " & Mid(XMLString1, 20)

            sErro = "112"
            sMsg1 = "vai gravar RPSWEBLoteLog"

            If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "Iniciando a consulta do lote", 0)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            sLote = lLote
            sProtocolo = objConsLoteRpsEnvio.Protocolo

            Form1.Msg.Items.Add("Iniciando a consulta do lote - Aguarde")

            If Form1.Msg.Items.Count - 15 < 1 Then
                Form1.Msg.TopIndex = 1
            Else
                Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
            End If

            Application.DoEvents()



            '                If UCase(objEndFilial.Cidade) = "TATUÍ" Or UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then
            If UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then

                sErro = "113"
                sMsg1 = "vai assinar o xml de consulta"


                Dim AD2 As AssinaturaDigital = New AssinaturaDigital

                AD2.Assinar(XMLString1, "Protocolo", cert, objEndFilial.Cidade)

                Dim xMlD1 As XmlDocument

                xMlD1 = AD2.XMLDocAssinado()

                XMLString1 = AD2.XMLStringAssinado

            End If


            sErro = "114"
            sMsg1 = "vai consultar o lote"

            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then

                XMLStringCabec = Replace(XMLStringCabec, "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd", "http://www.abrasf.org.br/nfse.xsd")
                XMLString1 = Replace(XMLString1, "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd", "http://www.abrasf.org.br/nfse.xsd")




                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homologacaobh.ClientCertificates.Add(cert)
                    XMLStringConsultarLoteRPSResposta = homologacaobh.ConsultarLoteRps(XMLStringCabec, XMLString1)
                Else
                    producaobh.ClientCertificates.Add(cert)
                    XMLStringConsultarLoteRPSResposta = producaobh.ConsultarLoteRps(XMLStringCabec, XMLString1)
                End If

                XMLStringConsultarLoteRPSResposta = Replace(XMLStringConsultarLoteRPSResposta, "http://www.abrasf.org.br/nfse.xsd", "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd")


                'resRPSWEBLote = db2.ExecuteQuery(Of RPSWEBLote) _
                '    ("SELECT * FROM RPSWEBLote WHERE Lote = {0} ORDER BY NumIntNF", lLote)

                'For Each objRPSWEBLote In resRPSWEBLote

                '    resNFiscal = db1.ExecuteQuery(Of NFiscal) _
                '        ("SELECT * FROM NFiscal WHERE NumINtDoc = {0}", objRPSWEBLote.NumIntNF)

                '    objNFiscal = resNFiscal(0)

                '    objConsNfseRpsEnvio.Prestador = New tcIdentificacaoPrestador

                '    objConsNfseRpsEnvio.Prestador.Cnpj = objFilialEmpresa.CGC
                '    Call Formata_String_Numero(objFilialEmpresa.InscricaoMunicipal, objConsNfseRpsEnvio.Prestador.InscricaoMunicipal)

                '    objConsNfseRpsEnvio.IdentificacaoRps = New tcIdentificacaoRps
                '    objConsNfseRpsEnvio.IdentificacaoRps.Numero = objNFiscal.NumNotaFiscal
                '    objConsNfseRpsEnvio.IdentificacaoRps.Serie = ADM.Serie_Sem_E(objNFiscal.Serie)
                '    objConsNfseRpsEnvio.IdentificacaoRps.Tipo = 1

                '    Dim mySerializerx3 As New XmlSerializer(GetType(ConsultarNfseRpsEnvio))

                '    XMLStreamRPS = New MemoryStream(10000)
                '    mySerializerx3.Serialize(XMLStreamRPS, objConsNfseRpsEnvio)

                '    xm22 = XMLStreamRPS.ToArray

                '    XMLString2 = System.Text.Encoding.UTF8.GetString(xm22)

                '    XMLString2 = Mid(XMLString2, 1, 19) & " encoding=""utf-8"" " & Mid(XMLString2, 20)

                '    Dim XMLStringConsultarNfseRPSResposta As String

                '    If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then


                '        homologacaobh.ClientCertificates.Add(cert)
                '        '                            homologacaobh.ClientCertificates.Add(cert)

                '        '                            XMLStringConsultarNfseRPSResposta = homologacaobh.ConsultarNfsePorRps(XMLStringCabec, XMLString2)
                '        XMLStringConsultarNfseRPSResposta = homologacaobh.ConsultarNfsePorRps(XMLStringCabec, XMLString2)

                '    Else
                '        producaobh.ClientCertificates.Add(cert)
                '        XMLStringConsultarNfseRPSResposta = producaobh.ConsultarNfsePorRps(XMLStringCabec, XMLString2)
                '    End If

                '    xRetRPS = System.Text.Encoding.UTF8.GetBytes(XMLStringConsultarNfseRPSResposta)
                '    XMLStreamRetRPS = New MemoryStream(10000)
                '    XMLStreamRetRPS.Write(xRetRPS, 0, xRetRPS.Length)

                '    Dim mySerializerConsultarNfseRPSResposta As New XmlSerializer(GetType(ConsultarNfseRpsResposta))

                '    Dim objConsultarNfseRPSResposta = New ConsultarNfseRpsResposta

                '    XMLStreamRetRPS.Position = 0

                '    objConsultarNfseRPSResposta = mySerializerConsultarNfseRPSResposta.Deserialize(XMLStreamRetRPS)

                '    sAuxRPS = objConsultarNfseRPSResposta.Item.GetType.ToString()
                '    If InStr(sAuxRPS, ".") <> 0 Then
                '        sAuxRPS = Mid(sAuxRPS, InStr(sAuxRPS, ".") + 1)
                '    End If

                '    Select Case sAuxRPS

                '        Case "ListaMensagemRetorno"

                '            objListaMsgRetorno = objConsultarNfseRPSResposta.Item

                '            For iIndiceRPS = 0 To objListaMsgRetorno.MensagemRetorno.Count - 1

                '                objMsgRetorno = objListaMsgRetorno.MensagemRetorno(iIndiceRPS)

                '                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Correcao, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9})", _
                '                lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objMsgRetorno.Codigo, Left(objMsgRetorno.Mensagem, 200), Left(objMsgRetorno.Correcao, 200), Now.Date, TimeOfDay.ToOADate)

                '                lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                '                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                '                lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Retorno da consulta de RPS - Serie = & " & objConsNfseRpsEnvio.IdentificacaoRps.Serie & " NFiscal = " & objConsNfseRpsEnvio.IdentificacaoRps.Numero & " - " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao, 255), 0)

                '                lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                '                Form1.Msg.Items.Add("Retorno da consulta de RPS - Serie = " & objConsNfseRpsEnvio.IdentificacaoRps.Serie & " NFiscal = " & objConsNfseRpsEnvio.IdentificacaoRps.Numero & " - " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao)

                '                If Form1.Msg.Items.Count - 15 < 1 Then
                '                    Form1.Msg.TopIndex = 1
                '                Else
                '                    Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                '                End If

                '                Application.DoEvents()

                '            Next

                '        Case "tcCompNfse"

                '            objCompNfse = objConsultarNfseRPSResposta.Item


                '            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, tipo, serie, numero, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11})", _
                '            lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, 0, "Nota Fiscal de Servico gerada - " & objCompNfse.Nfse.InfNfse.Numero, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, Now.Date, TimeOfDay.ToOADate)

                '            lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                '            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                '            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, 255), 0)

                '            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                '            Form1.Msg.Items.Add("Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero)

                '            If Form1.Msg.Items.Count - 15 < 1 Then

                '                Form1.Msg.TopIndex = 1
                '            Else
                '                Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                '            End If


                '            Application.DoEvents()

                '            lNumIntNFiscal = objRPSWEBLote.NumIntNF

                '            '******** pega o ultimo retorno de envio do lote em questao *************
                '            resRPSWEBProt = db2.ExecuteQuery(Of RPSWEBProt) _
                '            ("SELECT * FROM RPSWEBProt WHERE FilialEmpresa = {0} AND NumIntNF = {1} AND Ambiente = {2}", iFilialEmpresa, lNumIntNFiscal, iRPSAmbiente)

                '            If resRPSWEBProt.Count = 0 Then

                '                If objCompNfse.Nfse.InfNfse.PrestadorServico.Contato Is Nothing Then
                '                    sEmailPrestador = ""
                '                    sTelefonePrestador = ""
                '                Else
                '                    sEmailPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Contato.Email
                '                    If sEmailPrestador = Nothing Then sEmailPrestador = ""
                '                    sTelefonePrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Contato.Telefone
                '                    If sTelefonePrestador = Nothing Then sTelefonePrestador = ""
                '                End If

                '                If objCompNfse.Nfse.InfNfse.TomadorServico.Contato Is Nothing Then
                '                    sEmailTomador = ""
                '                    sTelefoneTomador = ""
                '                Else
                '                    sEmailTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Contato.Email
                '                    If sEmailTomador = Nothing Then sEmailTomador = ""
                '                    sTelefoneTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Contato.Telefone
                '                    If sTelefoneTomador = Nothing Then sTelefoneTomador = ""
                '                End If

                '                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBProt ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, data, hora, Versao, CodigoVerificacao, Competencia, DataEmissao, DataEmissaoRPS, Id,  TipoRPS, SerieRPS, NumeroRPS, NaturezaOperacao, Numero, OptanteSimplesNacional, OrgaoGeradorCodMun, OrgaoGeradorUF, " & _
                '                "OutrasInformacoes, ContatoEmail, ContatoTelefone, PrestadorBairro, PrestadorCEP, PrestadorCodMun, PrestadorComplemento, PrestadorEndereco, PrestadorNumero, PrestadorUF, PrestadorCNPJ, PrestadorInscMun, PrestadorNomeFantasia, PrestadorRazaoSocial, " & _
                '                "RegimeEspecialTributaco, ServicoCodigoCNAE,ServicoCodMun, ServicoCodTrib, ServicoDiscriminacao, ServicoItemLista, Aliquota, BaseCalculo, DescontoCondicionado, DescontoIncondicionado, ISSRetido, OutrasRetencoes, " & _
                '                "ValorCofins, ValorCsll, ValorDeducoes, ValorINSS, ValorIR, ValorISS, ValorISSRetido, ValorLiquidoNfse, ValorPIS, ValorServicos,  TomadorEmail, TomadorTelefone, TomadorBairro, TomadorCep, TomadorCodMun, " & _
                '                "TomadorComplemento, TomandorEndereco, TomadorNumero, TomadorUF, TomadorCPFCNPJ, TomadorInscMun, TomadorRazaoSocial, ValorCredito, NumIntNF)  VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, " & _
                '                "{20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, {59}, {60}, {61}, {62}, {63}, {64}, {65}, {66}, {67}, {68}, {69}, {70})", _
                '                lRPSWEBProtNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, Now.Date, TimeOfDay.ToOADate, "", objCompNfse.Nfse.InfNfse.CodigoVerificacao, objCompNfse.Nfse.InfNfse.Competencia, objCompNfse.Nfse.InfNfse.DataEmissao, objCompNfse.Nfse.InfNfse.DataEmissaoRps, IIf(objCompNfse.Nfse.InfNfse.Id = Nothing, "", objCompNfse.Nfse.InfNfse.Id), objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, _
                '                objCompNfse.Nfse.InfNfse.NaturezaOperacao, objCompNfse.Nfse.InfNfse.Numero, objCompNfse.Nfse.InfNfse.OptanteSimplesNacional, objCompNfse.Nfse.InfNfse.OrgaoGerador.CodigoMunicipio, objCompNfse.Nfse.InfNfse.OrgaoGerador.Uf, IIf(objCompNfse.Nfse.InfNfse.OutrasInformacoes = Nothing, "", objCompNfse.Nfse.InfNfse.OutrasInformacoes), _
                '                sEmailPrestador, sTelefonePrestador, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Bairro, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Cep, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.CodigoMunicipio, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Complemento, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Endereco, _
                '                objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Numero, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Uf, objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.Cnpj, objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.InscricaoMunicipal, IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia = Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia), objCompNfse.Nfse.InfNfse.PrestadorServico.RazaoSocial, _
                '                objCompNfse.Nfse.InfNfse.RegimeEspecialTributacao, objCompNfse.Nfse.InfNfse.Servico.CodigoCnae, objCompNfse.Nfse.InfNfse.Servico.CodigoMunicipio, objCompNfse.Nfse.InfNfse.Servico.CodigoTributacaoMunicipio, Left(objCompNfse.Nfse.InfNfse.Servico.Discriminacao, 255), _
                '                objCompNfse.Nfse.InfNfse.Servico.ItemListaServico, objCompNfse.Nfse.InfNfse.Servico.Valores.Aliquota, objCompNfse.Nfse.InfNfse.Servico.Valores.BaseCalculo, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoCondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoIncondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.IssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.OutrasRetencoes, _
                '                objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCofins, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCsll, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorDeducoes, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorInss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIr, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorLiquidoNfse, _
                '                objCompNfse.Nfse.InfNfse.Servico.Valores.ValorPis, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorServicos, sEmailTomador, sTelefoneTomador, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Bairro, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Cep, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.CodigoMunicipio, _
                '                IIf(objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento = Nothing, "", objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento), objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Endereco, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Numero, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Uf, objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.CpfCnpj.Item, IIf(objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal = Nothing, "", objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal), _
                '                objCompNfse.Nfse.InfNfse.TomadorServico.RazaoSocial, objCompNfse.Nfse.InfNfse.ValorCredito, lNumIntNFiscal)

                '                lRPSWEBProtNumIntDoc = lRPSWEBProtNumIntDoc + 1


                '            End If


                '    End Select

                'Next


            ElseIf UCase(objEndFilial.Cidade) = "SALVADOR" Then

                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homosalvadorconslote.ClientCertificates.Add(cert)
                    XMLStringConsultarLoteRPSResposta = homosalvadorconslote.ConsultarLoteRPS(XMLString1)
                Else
                    prodsalvadorconslote.ClientCertificates.Add(cert)
                    XMLStringConsultarLoteRPSResposta = prodsalvadorconslote.ConsultarLoteRPS(XMLString1)
                End If

            ElseIf UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then

                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homologacaoginfes.ClientCertificates.Add(cert)
                    XMLStringConsultarLoteRPSResposta = homologacaoginfes.ConsultarLoteRpsV3(XMLStringCabec, XMLString1)
                Else
                    producaoginfes.ClientCertificates.Add(cert)
                    XMLStringConsultarLoteRPSResposta = producaoginfes.ConsultarLoteRpsV3(XMLStringCabec, XMLString1)
                End If

                XMLStringConsultarLoteRPSResposta = Replace(XMLStringConsultarLoteRPSResposta, "ListaMensagemRetorno", "ns4:ListaMensagemRetorno")

                Dim iPos1 As Integer
                Dim sData As String

                iPos = InStr(XMLStringConsultarLoteRPSResposta, "<ns4:RegimeEspecialTributacao>")
                iPos1 = InStr(XMLStringConsultarLoteRPSResposta, "</ns4:RegimeEspecialTributacao>")

                If iPos <> 0 Then
                    XMLStringConsultarLoteRPSResposta = Mid(XMLStringConsultarLoteRPSResposta, 1, iPos - 1) & Mid(XMLStringConsultarLoteRPSResposta, iPos1 + 31)
                End If

                iPos = InStr(XMLStringConsultarLoteRPSResposta, "<ns4:DataEmissaoRps>")
                iPos1 = InStr(XMLStringConsultarLoteRPSResposta, "</ns4:DataEmissaoRps>")

                If iPos <> 0 Then
                    sData = Mid(XMLStringConsultarLoteRPSResposta, iPos + 20, iPos1 - (iPos + 20))

                    iPos = InStr(XMLStringConsultarLoteRPSResposta, "<ns4:Competencia>")
                    iPos1 = InStr(XMLStringConsultarLoteRPSResposta, "</ns4:Competencia>")


                    If iPos <> 0 Then
                        XMLStringConsultarLoteRPSResposta = Mid(XMLStringConsultarLoteRPSResposta, 1, iPos + 16) & sData & Mid(XMLStringConsultarLoteRPSResposta, iPos1)
                    End If


                End If

            ElseIf UCase(objEndFilial.Cidade) = "TATUÍ" Then


                'If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                '    homologacaosistema4rconsulta.ClientCertificates.Add(cert)
                '    XMLStringConsultarNfseRpsResposta = homologacaosistema4rconsulta.Execute(XMLString1)
                'Else
                '    producaosisstema4rconsulta.ClientCertificates.Add(cert)
                '    XMLStringConsultarNfseRpsResposta = producaosisstema4rconsulta.Execute(XMLString1)
                'End If

                lErro = ConsultaRPS_Aux(objConsNfseRpsEnvio, sCertificado, iRPSAmbiente, objEndFilial, XMLStringCabec, db2, lRPSWEBConsLoteNumIntDoc, lRPSWEBLoteLogNumIntDoc, iFilialEmpresa, db1, sLote, sProtocolo, lRPSWEBProtNumIntDoc, lLote, iDebug, objFilialEmpresa)

                If lErro = 1 Then
                    Throw New System.Exception("Erro na execução da rotina ConsultaRPS_Aux")
                End If


                GoTo Label_Fim

                'se nao for bh
            Else




                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homologacao.ClientCredentials.ClientCertificate.Certificate = cert
                    XMLStringConsultarLoteRPSResposta = homologacao.ConsultarLoteRps(XMLString1)
                Else
                    producao.ClientCredentials.ClientCertificate.Certificate = cert
                    XMLStringConsultarLoteRPSResposta = producao.ConsultarLoteRps(XMLString1)
                End If

            End If

            If UCase(objEndFilial.Cidade) = "SALVADOR" Then
                XMLStringConsultarLoteRPSResposta = Replace(XMLStringConsultarLoteRPSResposta, "id=", "Id=")
            End If







            '************* valida dados antes do envio **********************
            'Dim xDados As Byte()
            'Dim sArquivo As String

            ' ''            xDados = System.Text.Encoding.UTF8.GetBytes(xString)
            'xDados = System.Text.Encoding.UTF8.GetBytes(XMLStringConsultarLoteRPSResposta)

            'XMLStreamDados = New MemoryStream(10000)

            'XMLStreamDados.Write(xDados, 0, xDados.Length)

            'Dim DocDados As XmlDocument = New XmlDocument
            'XMLStreamDados.Position = 0
            'DocDados.Load(XMLStreamDados)
            'sArquivo = sDir & "teste.xml"
            'DocDados.Save(sArquivo)


            ''            Dim lErro As Long

            ''            lErro = objValidaXML.validaXML(sArquivo, sDir1 & "\nfse.xsd", lLote, lNumIntNF, db2, iFilialEmpresa)
            'lErro = objValidaXML.validaXML(sArquivo, "c:\nfeservico\tatui\servico_consultar_lote_rps_resposta_v03.xsd", lLote, lNumIntNF, db2, iFilialEmpresa)
            'If lErro = 1 Then

            '    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5})", _
            '    iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "ERRO - Encerrado o envio deste lote", 0)

            '    Form1.Msg.Items.Add("ERRO - o envio do lote " & CStr(lLote) & " foi encerrado por erro.")

            '    Application.DoEvents()

            '    Exit Try
            'End If







            sErro = "115"
            sMsg1 = "vai deserializar a resposta da consulta do lote"

            If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

            If iDebug = 1 Then MsgBox(XMLStringConsultarLoteRPSResposta)

            If iDebug = 1 Then MsgBox(XMLStringConsultarNfseRPSResposta)

            Dim objConsultarLoteRPSResposta = New ConsultarLoteRpsResposta

            xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringConsultarLoteRPSResposta)

            XMLStreamRet = New MemoryStream(10000)
            XMLStreamRet.Write(xRet, 0, xRet.Length)

            Dim mySerializerConsultarLoteRPSResposta As New XmlSerializer(GetType(ConsultarLoteRpsResposta))

            XMLStreamRet.Position = 0

            objConsultarLoteRPSResposta = mySerializerConsultarLoteRPSResposta.Deserialize(XMLStreamRet)

            sAux = objConsultarLoteRPSResposta.Item.GetType.ToString()


            If InStr(sAux, ".") <> 0 Then
                sAux = Mid(sAux, InStr(sAux, ".") + 1)
            End If

            Select Case sAux

                Case "ListaMensagemRetorno"

                    sErro = "116"
                    sMsg1 = "vai trata ListaMensagemRetorno"

                    If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

                    objListaMsgRetorno = objConsultarLoteRPSResposta.Item

                    For iIndice1 = 0 To objListaMsgRetorno.MensagemRetorno.Count - 1

                        objMsgRetorno = objListaMsgRetorno.MensagemRetorno(iIndice1)

                        If objMsgRetorno.Codigo = "E10" And UCase(objEndFilial.Cidade) <> "SALVADOR" Then


                            lErro = ConsultaRPS_Aux(objConsNfseRpsEnvio, sCertificado, iRPSAmbiente, objEndFilial, XMLStringCabec, db2, lRPSWEBConsLoteNumIntDoc, lRPSWEBLoteLogNumIntDoc, iFilialEmpresa, db1, sLote, sProtocolo, lRPSWEBProtNumIntDoc, lLote, iDebug, objFilialEmpresa)

                            If lErro = 1 Then
                                Throw New System.Exception("Erro na execução da rotina ConsultaRPS_Aux")
                            End If


                            '                            objConsNfseRpsEnvio.Prestador = New tcIdentificacaoPrestador

                            '#If ABRASF2 Then
                            '                            If Len(objFilialEmpresa.CGC) = 14 Then
                            '                                objConsNfseRpsEnvio.Prestador.CpfCnpj = New tcCpfCnpj
                            '                                objConsNfseRpsEnvio.Prestador.CpfCnpj.ItemElementName = ItemChoiceType.Cnpj
                            '                                objConsNfseRpsEnvio.Prestador.CpfCnpj.Item = objFilialEmpresa.CGC
                            '                            ElseIf Len(objFilialEmpresa.CGC) = 11 Then
                            '                                objConsNfseRpsEnvio.Prestador.CpfCnpj = New tcCpfCnpj
                            '                                objConsNfseRpsEnvio.Prestador.CpfCnpj.ItemElementName = ItemChoiceType.Cpf
                            '                                objConsNfseRpsEnvio.Prestador.CpfCnpj.Item = objFilialEmpresa.CGC
                            '                            End If

                            '#Else
                            '                            objConsNfseRpsEnvio.Prestador.Cnpj = objFilialEmpresa.CGC

                            '#End If


                            '                            Call Formata_String_Numero(objFilialEmpresa.InscricaoMunicipal, objConsNfseRpsEnvio.Prestador.InscricaoMunicipal)

                            '                            objConsNfseRpsEnvio.IdentificacaoRps = New tcIdentificacaoRps


                            '                            resRPSWEBLote = db2.ExecuteQuery(Of RPSWEBLote) _
                            '                                ("SELECT * FROM RPSWEBLote WHERE Lote = {0} ORDER BY NumIntNF", lLote)

                            '                            For Each objRPSWEBLote In resRPSWEBLote

                            '                                resNFiscal = db1.ExecuteQuery(Of NFiscal) _
                            '                                    ("SELECT * FROM NFiscal WHERE NumINtDoc = {0}", objRPSWEBLote.NumIntNF)

                            '                                objNFiscal = resNFiscal(0)

                            '                                objConsNfseRpsEnvio.IdentificacaoRps.Numero = objNFiscal.NumNotaFiscal
                            '                                objConsNfseRpsEnvio.IdentificacaoRps.Serie = objNFiscal.Serie
                            '                                objConsNfseRpsEnvio.IdentificacaoRps.Tipo = "1"

                            '                                sErro = "117"
                            '                                sMsg1 = "vai iniciar ConsultaRPS"

                            '                                If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

                            '                                lErro = ConsultaRPS(objConsNfseRpsEnvio, sCertificado, iRPSAmbiente, objEndFilial, XMLStringCabec, db2, lRPSWEBConsLoteNumIntDoc, lRPSWEBLoteLogNumIntDoc, iFilialEmpresa, db1, sLote, sProtocolo, lRPSWEBProtNumIntDoc, lLote)


                            '                                If lErro = 1 Then
                            '                                    Throw New System.Exception("Erro na execução da rotina ConsultaRPS")
                            '                                End If

                            '                            Next

                        Else

Label_118:
                            sErro = "118"
                            sMsg1 = "vai gravar RPSWEBConsLote"
                            If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)


                            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Correcao, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9})", _
                                lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objMsgRetorno.Codigo, Left(objMsgRetorno.Mensagem, 200), Left(objMsgRetorno.Correcao, 200), Now.Date, TimeOfDay.ToOADate)

                            lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                            sErro = "119"
                            sMsg1 = "vai gravar RPSWEBLoteLog"

                            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Retorno da consulta do lote - " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao, 255), 0)

                            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                            Form1.Msg.Items.Add("Retorno da consulta do lote - " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao)

                            If Form1.Msg.Items.Count - 15 < 1 Then
                                Form1.Msg.TopIndex = 1
                            Else
                                Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                            End If

                            Application.DoEvents()

                        End If

                    Next

                Case "ListaMensagemRetornoLote"

                    sErro = "120"
                    sMsg1 = "vai tratar ListaMensagemRetornoLote"

                    If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

#If ABRASF Then
                    objListaMsgRetornoLote = objConsultarLoteRPSResposta.Item

                    For iIndice1 = 0 To objListaMsgRetornoLote.MensagemRetorno.Count - 1

                        objMsgRetornoLote = objListaMsgRetornoLote.MensagemRetorno(iIndice1)

                        If objMsgRetornoLote.Codigo = "E10" Then

                            sErro = "120.1"
                            sMsg1 = "vai tratar o codigo E10"
                            If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)


                            objConsNfseRpsEnvio.Prestador = New tcIdentificacaoPrestador

                            objConsNfseRpsEnvio.Prestador.Cnpj = objFilialEmpresa.CGC

                            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then
                                Call ADM.Formata_String_AlfaNumerico(objFilialEmpresa.InscricaoMunicipal, objConsNfseRpsEnvio.Prestador.InscricaoMunicipal)
                            Else
                                Call ADM.Formata_String_Numero(objFilialEmpresa.InscricaoMunicipal, objConsNfseRpsEnvio.Prestador.InscricaoMunicipal)
                            End If

                            '                            Call Formata_String_Numero(objFilialEmpresa.InscricaoMunicipal, objConsNfseRpsEnvio.Prestador.InscricaoMunicipal)

                            objConsNfseRpsEnvio.IdentificacaoRps = New tcIdentificacaoRps
                            objConsNfseRpsEnvio.IdentificacaoRps.Numero = objMsgRetornoLote.IdentificacaoRps.Numero
                            objConsNfseRpsEnvio.IdentificacaoRps.Serie = objMsgRetornoLote.IdentificacaoRps.Serie
                            objConsNfseRpsEnvio.IdentificacaoRps.Tipo = objMsgRetornoLote.IdentificacaoRps.Tipo

                            sErro = "121"
                            sMsg1 = "vai executar ConsultaRPS"

                            If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

                            lErro = ConsultaRPS(objConsNfseRpsEnvio, sCertificado, iRPSAmbiente, objEndFilial, XMLStringCabec, db2, lRPSWEBConsLoteNumIntDoc, lRPSWEBLoteLogNumIntDoc, iFilialEmpresa, db1, sLote, sProtocolo, lRPSWEBProtNumIntDoc, lLote)


                            If lErro = 1 Then
                                Throw New System.Exception("Erro na execução da rotina ConsultaRPS")
                            End If

                            'Dim mySerializerx3 As New XmlSerializer(GetType(ConsultarNfseRpsEnvio))

                            'XMLStreamRPS = New MemoryStream(10000)
                            'mySerializerx3.Serialize(XMLStreamRPS, objConsNfseRpsEnvio)

                            'xm22 = XMLStreamRPS.ToArray

                            'XMLString2 = System.Text.Encoding.UTF8.GetString(xm22)

                            'XMLString2 = Mid(XMLString2, 1, 19) & " encoding=""utf-8"" " & Mid(XMLString2, 20)

                            'Dim XMLStringConsultarNfseRPSResposta As String

                            'cert = certificado.BuscaNome(sCertificado)

                            'If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then

                            '    XMLString2 = Replace(XMLString2, "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd", "http://www.abrasf.org.br/nfse.xsd")


                            '    If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                            '        homologacaobh.ClientCertificates.Add(cert)
                            '        XMLStringConsultarNfseRPSResposta = homologacaobh.ConsultarNfsePorRps(XMLStringCabec, XMLString2)
                            '    Else
                            '        producaobh.ClientCertificates.Add(cert)
                            '        XMLStringConsultarNfseRPSResposta = producaobh.ConsultarNfsePorRps(XMLStringCabec, XMLString2)
                            '    End If

                            '    XMLStringConsultarNfseRPSResposta = Replace(XMLStringConsultarNfseRPSResposta, "http://www.abrasf.org.br/nfse.xsd", "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd")


                            'Else

                            '    If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                            '        homologacao.ClientCredentials.ClientCertificate.Certificate = cert
                            '        XMLStringConsultarNfseRPSResposta = homologacao.ConsultarNfsePorRps(XMLString2)
                            '    Else
                            '        producaobh.ClientCertificates.Add(cert)
                            '        XMLStringConsultarNfseRPSResposta = producao.ConsultarNfsePorRps(XMLString2)
                            '    End If

                            'End If

                            'xRetRPS = System.Text.Encoding.UTF8.GetBytes(XMLStringConsultarNfseRPSResposta)

                            'XMLStreamRetRPS = New MemoryStream(10000)
                            'XMLStreamRetRPS.Write(xRetRPS, 0, xRetRPS.Length)

                            'Dim mySerializerConsultarNfseRPSResposta As New XmlSerializer(GetType(ConsultarNfseRpsResposta))

                            'Dim objConsultarNfseRPSResposta = New ConsultarNfseRpsResposta

                            'XMLStreamRetRPS.Position = 0

                            'objConsultarNfseRPSResposta = mySerializerConsultarNfseRPSResposta.Deserialize(XMLStreamRetRPS)

                            'sAuxRPS = objConsultarNfseRPSResposta.Item.GetType.ToString()
                            'If InStr(sAuxRPS, ".") <> 0 Then
                            '    sAuxRPS = Mid(sAuxRPS, InStr(sAuxRPS, ".") + 1)
                            'End If

                            'Select Case sAuxRPS

                            '    Case "ListaMensagemRetorno"

                            '        objListaMsgRetorno = objConsultarNfseRPSResposta.Item

                            '        For iIndiceRPS = 0 To objListaMsgRetorno.MensagemRetorno.Count - 1

                            '            objMsgRetorno = objListaMsgRetorno.MensagemRetorno(iIndiceRPS)

                            '            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Correcao, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9})", _
                            '            lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objMsgRetorno.Codigo, Left(objMsgRetorno.Mensagem, 200), Left(objMsgRetorno.Correcao, 200), Now.Date, TimeOfDay.ToOADate)

                            '            lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                            '            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                            '            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Retorno da consulta de RPS - Serie = & " & objConsNfseRpsEnvio.IdentificacaoRps.Serie & " NFiscal = " & objConsNfseRpsEnvio.IdentificacaoRps.Numero & " - " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao, 255), 0)

                            '            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                            '            Form1.Msg.Items.Add("Retorno da consulta de RPS - Serie = & " & objConsNfseRpsEnvio.IdentificacaoRps.Serie & " NFiscal = " & objConsNfseRpsEnvio.IdentificacaoRps.Numero & " - " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao)

                            '            If Form1.Msg.Items.Count - 15 < 1 Then
                            '                Form1.Msg.TopIndex = 1
                            '            Else
                            '                Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                            '            End If

                            '            Application.DoEvents()

                            '        Next

                            '    Case "tcCompNfse"

                            '        objCompNfse = objConsultarNfseRPSResposta.Item

                            '        resRPSWEBLote = db2.ExecuteQuery(Of RPSWEBLote) _
                            '        ("SELECT * FROM RPSWEBLote WHERE Lote = {0} ORDER BY NumIntNF", lLote)

                            '        colNumIntNFiscal = New Collection

                            '        For Each objRPSWEBLote In resRPSWEBLote
                            '            colNumIntNFiscal.Add(objRPSWEBLote.NumIntNF)
                            '        Next


                            '        iAchou = 0

                            '        For iIndice2 = 1 To colNumIntNFiscal.Count
                            '            lNumIntNFiscal = colNumIntNFiscal(iIndice2)

                            '            resNFiscal = db1.ExecuteQuery(Of NFiscal) _
                            '            ("SELECT * FROM NFiscal WHERE NumINtDoc = {0}", lNumIntNFiscal)

                            '            objNFiscal = resNFiscal(0)

                            '            If Format(CInt(ADM.Serie_Sem_E(objNFiscal.Serie)), "000") = Format(CInt(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie), "000") And _
                            '                   Format(objNFiscal.NumNotaFiscal, "000000000") = Format(CLng(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero), "000000000") Then
                            '                iAchou = 1
                            '                Exit For
                            '            End If

                            '        Next

                            '        If iAchou = 0 Then
                            '            Throw New System.Exception("A nota consultada nao corresponde a nenhuma das enviadas. Serie Consulta = " & Format(CInt(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie), "000") & " Numero Consulta = " & Format(CLng(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero), "000000000"))
                            '        End If

                            '        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, tipo, serie, numero, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11})", _
                            '        lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, 0, "Nota Fiscal de Servico gerada - " & objCompNfse.Nfse.InfNfse.Numero, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, Now.Date, TimeOfDay.ToOADate)

                            '        lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                            '        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                            '        lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, 255), 0)

                            '        lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                            '        Form1.Msg.Items.Add("Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero)

                            '        If Form1.Msg.Items.Count - 15 < 1 Then
                            '            Form1.Msg.TopIndex = 1
                            '        Else
                            '            Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                            '        End If

                            '        Application.DoEvents()

                            '        '******** pega o ultimo retorno de envio do lote em questao *************
                            '        resRPSWEBProt = db2.ExecuteQuery(Of RPSWEBProt) _
                            '        ("SELECT * FROM RPSWEBProt WHERE FilialEmpresa = {0} AND NumIntNF = {1} AND Ambiente = {2}", iFilialEmpresa, lNumIntNFiscal, iRPSAmbiente)

                            '        If resRPSWEBProt.Count = 0 Then

                            '            If objCompNfse.Nfse.InfNfse.PrestadorServico.Contato Is Nothing Then
                            '                sEmailPrestador = ""
                            '                sTelefonePrestador = ""
                            '            Else
                            '                sEmailPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Contato.Email
                            '                If sEmailPrestador = Nothing Then sEmailPrestador = ""
                            '                sTelefonePrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Contato.Telefone
                            '                If sTelefonePrestador = Nothing Then sTelefonePrestador = ""
                            '            End If

                            '            If objCompNfse.Nfse.InfNfse.TomadorServico.Contato Is Nothing Then
                            '                sEmailTomador = ""
                            '                sTelefoneTomador = ""
                            '            Else
                            '                sEmailTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Contato.Email
                            '                If sEmailTomador = Nothing Then sEmailTomador = ""
                            '                sTelefoneTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Contato.Telefone
                            '                If sTelefoneTomador = Nothing Then sTelefoneTomador = ""
                            '            End If

                            '            'iResult = db2.ExecuteCommand("INSERT INTO RPSWEBProt ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, data, hora, Versao, CodigoVerificacao, Competencia, DataEmissao, DataEmissaoRPS, Id,  TipoRPS, SerieRPS, NumeroRPS, NaturezaOperacao, Numero, OptanteSimplesNacional, OrgaoGeradorCodMun, OrgaoGeradorUF, " & _
                            '            '"OutrasInformacoes, ContatoEmail, ContatoTelefone, PrestadorBairro, PrestadorCEP, PrestadorCodMun, PrestadorComplemento, PrestadorEndereco, PrestadorNumero, PrestadorUF, PrestadorCNPJ, PrestadorInscMun, PrestadorNomeFantasia, PrestadorRazaoSocial, " & _
                            '            '"RegimeEspecialTributaco, ServicoCodigoCNAE,ServicoCodMun, ServicoCodTrib, ServicoDiscriminacao, ServicoItemLista, Aliquota, BaseCalculo, DescontoCondicionado, DescontoIncondicionado, ISSRetido, OutrasRetencoes, " & _
                            '            '"ValorCofins, ValorCsll, ValorDeducoes, ValorINSS, ValorIR, ValorISS, ValorISSRetido, ValorLiquidoNfse, ValorPIS, ValorServicos,  TomadorEmail, TomadorTelefone, TomadorBairro, TomadorCep, TomadorCodMun, " & _
                            '            '"TomadorComplemento, TomandorEndereco, TomadorNumero, TomadorUF, TomadorCPFCNPJ, TomadorInscMun, TomadorRazaoSocial, ValorCredito, NumIntNF)  VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, " & _
                            '            '"{20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, {59}, {60}, {61}, {62}, {63}, {64}, {65}, {66}, {67}, {68}, {69}, {70})", _
                            '            'lRPSWEBProtNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, Now.Date, TimeOfDay.ToOADate, "", objCompNfse.Nfse.InfNfse.CodigoVerificacao, objCompNfse.Nfse.InfNfse.Competencia, objCompNfse.Nfse.InfNfse.DataEmissao, objCompNfse.Nfse.InfNfse.DataEmissaoRps, IIf(objCompNfse.Nfse.InfNfse.Id = Nothing, "", objCompNfse.Nfse.InfNfse.Id), objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, _
                            '            'objCompNfse.Nfse.InfNfse.NaturezaOperacao, objCompNfse.Nfse.InfNfse.Numero, objCompNfse.Nfse.InfNfse.OptanteSimplesNacional, objCompNfse.Nfse.InfNfse.OrgaoGerador.CodigoMunicipio, objCompNfse.Nfse.InfNfse.OrgaoGerador.Uf, IIf(objCompNfse.Nfse.InfNfse.OutrasInformacoes = Nothing, "", objCompNfse.Nfse.InfNfse.OutrasInformacoes), _
                            '            'sEmailPrestador, sTelefonePrestador, IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Bairro), IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Cep), IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.CodigoMunicipio), IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Complemento), IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Endereco), _
                            '            'IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Numero), IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Uf), objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.Cnpj, objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.InscricaoMunicipal, IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia = Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia), objCompNfse.Nfse.InfNfse.PrestadorServico.RazaoSocial, _
                            '            'objCompNfse.Nfse.InfNfse.RegimeEspecialTributacao, objCompNfse.Nfse.InfNfse.Servico.CodigoCnae, objCompNfse.Nfse.InfNfse.Servico.CodigoMunicipio, objCompNfse.Nfse.InfNfse.Servico.CodigoTributacaoMunicipio, Left(objCompNfse.Nfse.InfNfse.Servico.Discriminacao, 255), _
                            '            'objCompNfse.Nfse.InfNfse.Servico.ItemListaServico, objCompNfse.Nfse.InfNfse.Servico.Valores.Aliquota, objCompNfse.Nfse.InfNfse.Servico.Valores.BaseCalculo, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoCondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoIncondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.IssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.OutrasRetencoes, _
                            '            'objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCofins, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCsll, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorDeducoes, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorInss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIr, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorLiquidoNfse, _
                            '            'objCompNfse.Nfse.InfNfse.Servico.Valores.ValorPis, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorServicos, sEmailTomador, sTelefoneTomador, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Bairro, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Cep, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.CodigoMunicipio, _
                            '            'IIf(objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento = Nothing, "", objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento), objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Endereco, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Numero, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Uf, objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.CpfCnpj.Item, IIf(objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal = Nothing, "", objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal), _
                            '            'objCompNfse.Nfse.InfNfse.TomadorServico.RazaoSocial, objCompNfse.Nfse.InfNfse.ValorCredito, lNumIntNFiscal)


                            '            If objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing Then

                            '                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBProt ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, data, hora, Versao, CodigoVerificacao, Competencia, DataEmissao, DataEmissaoRPS, Id,  TipoRPS, SerieRPS, NumeroRPS, NaturezaOperacao, Numero, OptanteSimplesNacional, OrgaoGeradorCodMun, OrgaoGeradorUF, " & _
                            '                "OutrasInformacoes, ContatoEmail, ContatoTelefone, PrestadorBairro, PrestadorCEP, PrestadorCodMun, PrestadorComplemento, PrestadorEndereco, PrestadorNumero, PrestadorUF, PrestadorCNPJ, PrestadorInscMun, PrestadorNomeFantasia, PrestadorRazaoSocial, " & _
                            '                "RegimeEspecialTributaco, ServicoCodigoCNAE,ServicoCodMun, ServicoCodTrib, ServicoDiscriminacao, ServicoItemLista, Aliquota, BaseCalculo, DescontoCondicionado, DescontoIncondicionado, ISSRetido, OutrasRetencoes, " & _
                            '                "ValorCofins, ValorCsll, ValorDeducoes, ValorINSS, ValorIR, ValorISS, ValorISSRetido, ValorLiquidoNfse, ValorPIS, ValorServicos,  TomadorEmail, TomadorTelefone, TomadorBairro, TomadorCep, TomadorCodMun, " & _
                            '                "TomadorComplemento, TomandorEndereco, TomadorNumero, TomadorUF, TomadorCPFCNPJ, TomadorInscMun, TomadorRazaoSocial, ValorCredito, NumIntNF)  VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, " & _
                            '                "{20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, {59}, {60}, {61}, {62}, {63}, {64}, {65}, {66}, {67}, {68}, {69}, {70})", _
                            '                lRPSWEBProtNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, Now.Date, TimeOfDay.ToOADate, "", objCompNfse.Nfse.InfNfse.CodigoVerificacao, objCompNfse.Nfse.InfNfse.Competencia, objCompNfse.Nfse.InfNfse.DataEmissao, objCompNfse.Nfse.InfNfse.DataEmissaoRps, IIf(objCompNfse.Nfse.InfNfse.Id = Nothing, "", objCompNfse.Nfse.InfNfse.Id), objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, _
                            '                objCompNfse.Nfse.InfNfse.NaturezaOperacao, objCompNfse.Nfse.InfNfse.Numero, objCompNfse.Nfse.InfNfse.OptanteSimplesNacional, objCompNfse.Nfse.InfNfse.OrgaoGerador.CodigoMunicipio, objCompNfse.Nfse.InfNfse.OrgaoGerador.Uf, IIf(objCompNfse.Nfse.InfNfse.OutrasInformacoes = Nothing, "", objCompNfse.Nfse.InfNfse.OutrasInformacoes), _
                            '                sEmailPrestador, sTelefonePrestador, "", "", "", "", "", _
                            '                "", "", objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.Cnpj, objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.InscricaoMunicipal, IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia = Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia), objCompNfse.Nfse.InfNfse.PrestadorServico.RazaoSocial, _
                            '                objCompNfse.Nfse.InfNfse.RegimeEspecialTributacao, objCompNfse.Nfse.InfNfse.Servico.CodigoCnae, objCompNfse.Nfse.InfNfse.Servico.CodigoMunicipio, objCompNfse.Nfse.InfNfse.Servico.CodigoTributacaoMunicipio, Left(objCompNfse.Nfse.InfNfse.Servico.Discriminacao, 255), _
                            '                objCompNfse.Nfse.InfNfse.Servico.ItemListaServico, objCompNfse.Nfse.InfNfse.Servico.Valores.Aliquota, objCompNfse.Nfse.InfNfse.Servico.Valores.BaseCalculo, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoCondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoIncondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.IssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.OutrasRetencoes, _
                            '                objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCofins, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCsll, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorDeducoes, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorInss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIr, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorLiquidoNfse, _
                            '                objCompNfse.Nfse.InfNfse.Servico.Valores.ValorPis, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorServicos, sEmailTomador, sTelefoneTomador, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Bairro, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Cep, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.CodigoMunicipio, _
                            '                IIf(objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento = Nothing, "", objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento), objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Endereco, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Numero, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Uf, objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.CpfCnpj.Item, IIf(objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal = Nothing, "", objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal), _
                            '                objCompNfse.Nfse.InfNfse.TomadorServico.RazaoSocial, objCompNfse.Nfse.InfNfse.ValorCredito, lNumIntNFiscal)


                            '            Else

                            '                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBProt ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, data, hora, Versao, CodigoVerificacao, Competencia, DataEmissao, DataEmissaoRPS, Id,  TipoRPS, SerieRPS, NumeroRPS, NaturezaOperacao, Numero, OptanteSimplesNacional, OrgaoGeradorCodMun, OrgaoGeradorUF, " & _
                            '                "OutrasInformacoes, ContatoEmail, ContatoTelefone, PrestadorBairro, PrestadorCEP, PrestadorCodMun, PrestadorComplemento, PrestadorEndereco, PrestadorNumero, PrestadorUF, PrestadorCNPJ, PrestadorInscMun, PrestadorNomeFantasia, PrestadorRazaoSocial, " & _
                            '                "RegimeEspecialTributaco, ServicoCodigoCNAE,ServicoCodMun, ServicoCodTrib, ServicoDiscriminacao, ServicoItemLista, Aliquota, BaseCalculo, DescontoCondicionado, DescontoIncondicionado, ISSRetido, OutrasRetencoes, " & _
                            '                "ValorCofins, ValorCsll, ValorDeducoes, ValorINSS, ValorIR, ValorISS, ValorISSRetido, ValorLiquidoNfse, ValorPIS, ValorServicos,  TomadorEmail, TomadorTelefone, TomadorBairro, TomadorCep, TomadorCodMun, " & _
                            '                "TomadorComplemento, TomandorEndereco, TomadorNumero, TomadorUF, TomadorCPFCNPJ, TomadorInscMun, TomadorRazaoSocial, ValorCredito, NumIntNF)  VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, " & _
                            '                "{20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, {59}, {60}, {61}, {62}, {63}, {64}, {65}, {66}, {67}, {68}, {69}, {70})", _
                            '                lRPSWEBProtNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, Now.Date, TimeOfDay.ToOADate, "", objCompNfse.Nfse.InfNfse.CodigoVerificacao, objCompNfse.Nfse.InfNfse.Competencia, objCompNfse.Nfse.InfNfse.DataEmissao, objCompNfse.Nfse.InfNfse.DataEmissaoRps, IIf(objCompNfse.Nfse.InfNfse.Id = Nothing, "", objCompNfse.Nfse.InfNfse.Id), objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, _
                            '                objCompNfse.Nfse.InfNfse.NaturezaOperacao, objCompNfse.Nfse.InfNfse.Numero, objCompNfse.Nfse.InfNfse.OptanteSimplesNacional, objCompNfse.Nfse.InfNfse.OrgaoGerador.CodigoMunicipio, objCompNfse.Nfse.InfNfse.OrgaoGerador.Uf, IIf(objCompNfse.Nfse.InfNfse.OutrasInformacoes = Nothing, "", objCompNfse.Nfse.InfNfse.OutrasInformacoes), _
                            '                sEmailPrestador, sTelefonePrestador, IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Bairro), IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Cep), IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.CodigoMunicipio), IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Complemento), IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Endereco), _
                            '                IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Numero), IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Uf), objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.Cnpj, objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.InscricaoMunicipal, IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia = Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia), objCompNfse.Nfse.InfNfse.PrestadorServico.RazaoSocial, _
                            '                objCompNfse.Nfse.InfNfse.RegimeEspecialTributacao, objCompNfse.Nfse.InfNfse.Servico.CodigoCnae, objCompNfse.Nfse.InfNfse.Servico.CodigoMunicipio, objCompNfse.Nfse.InfNfse.Servico.CodigoTributacaoMunicipio, Left(objCompNfse.Nfse.InfNfse.Servico.Discriminacao, 255), _
                            '                objCompNfse.Nfse.InfNfse.Servico.ItemListaServico, objCompNfse.Nfse.InfNfse.Servico.Valores.Aliquota, objCompNfse.Nfse.InfNfse.Servico.Valores.BaseCalculo, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoCondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoIncondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.IssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.OutrasRetencoes, _
                            '                objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCofins, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCsll, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorDeducoes, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorInss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIr, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorLiquidoNfse, _
                            '                objCompNfse.Nfse.InfNfse.Servico.Valores.ValorPis, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorServicos, sEmailTomador, sTelefoneTomador, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Bairro, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Cep, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.CodigoMunicipio, _
                            '                IIf(objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento = Nothing, "", objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento), objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Endereco, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Numero, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Uf, objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.CpfCnpj.Item, IIf(objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal = Nothing, "", objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal), _
                            '                objCompNfse.Nfse.InfNfse.TomadorServico.RazaoSocial, objCompNfse.Nfse.InfNfse.ValorCredito, lNumIntNFiscal)

                            '            End If


                            '            lRPSWEBProtNumIntDoc = lRPSWEBProtNumIntDoc + 1


                            '        End If


                            'End Select

                        Else

                            sErro = "122"
                            sMsg1 = "vai gravar RPSWEBConsLote"

                            If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

                            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Tipo, Serie, Numero, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11})", _
                            lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objMsgRetornoLote.Codigo, objMsgRetornoLote.Mensagem, objMsgRetornoLote.IdentificacaoRps.Tipo, objMsgRetornoLote.IdentificacaoRps.Serie, objMsgRetornoLote.IdentificacaoRps.Numero, Now.Date, TimeOfDay.ToOADate)

                            lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                            sErro = "122"
                            sMsg1 = "vai gravar RPSWEBLoteLog"

                            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "Retorno da consulta do lote - " & objMsgRetornoLote.Codigo & " - " & objMsgRetornoLote.Mensagem & " Série = " & objMsgRetornoLote.IdentificacaoRps.Serie & " NFiscal = " & objMsgRetornoLote.IdentificacaoRps.Numero, 0)

                            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                            Form1.Msg.Items.Add("Retorno da consulta do lote - " & objMsgRetornoLote.Codigo & " - " & objMsgRetornoLote.Mensagem & " Série = " & objMsgRetornoLote.IdentificacaoRps.Serie & " NFiscal = " & objMsgRetornoLote.IdentificacaoRps.Numero)

                            If Form1.Msg.Items.Count - 15 < 1 Then
                                Form1.Msg.TopIndex = 1
                            Else
                                Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                            End If


                            Application.DoEvents()

                        End If

                    Next
#End If

                Case "ConsultarLoteRpsRespostaListaNfse"

                    sErro = "123"
                    sMsg1 = "vai tratar ConsultarLoteRpsRespostaListaNfse"
                    If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

                    objConsultarLoteRpsRespostaListaNFse = objConsultarLoteRPSResposta.Item

                    resRPSWEBLote = db2.ExecuteQuery(Of RPSWEBLote) _
                    ("SELECT * FROM RPSWEBLote WHERE Lote = {0} ORDER BY NumIntNF", lLote)

                    For Each objRPSWEBLote In resRPSWEBLote
                        colNumIntNFiscal.Add(objRPSWEBLote.NumIntNF)
                    Next

                    For iIndice1 = 0 To objConsultarLoteRpsRespostaListaNFse.CompNfse.Count - 1


                        objCompNfse = objConsultarLoteRpsRespostaListaNFse.CompNfse(iIndice1)

                        iAchou = 0

                        For iIndice2 = 1 To colNumIntNFiscal.Count
                            lNumIntNFiscal = colNumIntNFiscal(iIndice2)

                            resNFiscal = db1.ExecuteQuery(Of NFiscal) _
                            ("SELECT * FROM NFiscal WHERE NumINtDoc = {0}", lNumIntNFiscal)

                            objNFiscal = resNFiscal(0)
#If ABRASF2 Then
                            If ADM.Serie_Sem_E(objNFiscal.Serie) = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Serie And _
                                   Format(objNFiscal.NumNotaFiscal, "000000000") = Format(CLng(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Numero), "000000000") Then
#Else
                            If ADM.Serie_Sem_E(objNFiscal.Serie) = objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie And _
                                   Format(objNFiscal.NumNotaFiscal, "000000000") = Format(CLng(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero), "000000000") Then
#End If
                                iAchou = 1
                                Exit For
                            End If

                        Next

                        If iAchou = 0 Then

#If ABRASF2 Then

                            Throw New System.Exception("A nota consultada nao corresponde a nenhuma das enviadas. Serie Consulta = " & objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Serie & " Numero Consulta = " & Format(CLng(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Numero), "000000000"))
#Else
                            Throw New System.Exception("A nota consultada nao corresponde a nenhuma das enviadas. Serie Consulta = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie & " Numero Consulta = " & Format(CLng(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero), "000000000"))

#End If
                        End If

                        sErro = "124"
                        sMsg1 = "vai gravar RPSWEBConsLote"

#If ABRASF2 Then
                        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, tipo, serie, numero, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11})", _
                        lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, 0, "Nota Fiscal de Servico gerada - " & objCompNfse.Nfse.InfNfse.Numero, objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Tipo, objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Serie, objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Numero, Now.Date, TimeOfDay.ToOADate)
#Else
                        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, tipo, serie, numero, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11})", _
                        lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, 0, "Nota Fiscal de Servico gerada - " & objCompNfse.Nfse.InfNfse.Numero, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, Now.Date, TimeOfDay.ToOADate)
#End If


                        lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                        sErro = "125"
                        sMsg1 = "vai gravar RPSWEBLoteLog"

                        If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

#If ABRASF2 Then
                        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                        lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Numero, 255), 0)
#Else
                        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                        lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, 255), 0)
#End If

                        lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

#If ABRASF2 Then
                        Form1.Msg.Items.Add("Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Numero)
#Else
                        Form1.Msg.Items.Add("Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero)
#End If
                        If Form1.Msg.Items.Count - 15 < 1 Then
                            Form1.Msg.TopIndex = 1
                        Else
                            Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                        End If

                        Application.DoEvents()

                        sErro = "125"
                        sMsg1 = "vai consultar RPSWEBProt"

                        '******** pega o ultimo retorno de envio do lote em questao *************
                        resRPSWEBProt = db2.ExecuteQuery(Of RPSWEBProt) _
                        ("SELECT * FROM RPSWEBProt WHERE FilialEmpresa = {0} AND NumIntNF = {1} AND Ambiente = {2}", iFilialEmpresa, lNumIntNFiscal, iRPSAmbiente)

                        If resRPSWEBProt.Count = 0 Then

                            sErro = "126"
                            sMsg1 = "conulta lote nfse"
                            If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)


#If ABRASF2 Then
                            sEmailPrestador = ""
                            sTelefonePrestador = ""
#Else
                            If objCompNfse.Nfse.InfNfse.PrestadorServico.Contato Is Nothing Then
                                sEmailPrestador = ""
                                sTelefonePrestador = ""
                            Else
                                sEmailPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Contato.Email
                                If sEmailPrestador = Nothing Then sEmailPrestador = ""
                                sTelefonePrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Contato.Telefone
                                If sTelefonePrestador = Nothing Then sTelefonePrestador = ""
                            End If
#End If
                            sErro = "127"
                            sMsg1 = "conulta lote nfse"


#If ABRASF2 Then
                            If objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Contato Is Nothing Then
                                sEmailTomador = ""
                                sTelefoneTomador = ""
                            Else
                                sEmailTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Contato.Email
                                If sEmailTomador = Nothing Then sEmailTomador = ""
                                sTelefoneTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Contato.Telefone
                                If sTelefoneTomador = Nothing Then sTelefoneTomador = ""
                            End If
#Else
                            If objCompNfse.Nfse.InfNfse.TomadorServico.Contato Is Nothing Then
                                sEmailTomador = ""
                                sTelefoneTomador = ""
                            Else
                                sEmailTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Contato.Email
                                If sEmailTomador = Nothing Then sEmailTomador = ""
                                sTelefoneTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Contato.Telefone
                                If sTelefoneTomador = Nothing Then sTelefoneTomador = ""
                            End If
#End If

                            sErro = "128"
                            sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                            sComplementoTomador = ""
                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco Is Nothing Then
                                If objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Complemento = Nothing Then
                                    objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Complemento = ""
                                End If
                            End If
#Else
                            If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco Is Nothing Then
                                If objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento = Nothing Then
                                    objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento = ""
                                End If
                            End If
#End If

                            sErro = "129"
                            sMsg1 = "conulta lote nfse"

                            If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

#If Not ABRASF2 Then

                            If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing Then
                                If objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Complemento = Nothing Then
                                    objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Complemento = ""
                                End If
                            End If
#End If

                            sErro = "130"
                            sMsg1 = "conulta lote nfse"

#If ABRASF2 Then

                            If objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.IdentificacaoTomador Is Nothing Then
                                sCNPJCPFTomador = ""
                                sInscricaoMunTomador = ""
                            Else
                                sCNPJCPFTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.IdentificacaoTomador.CpfCnpj.Item
                                If objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.IdentificacaoTomador.InscricaoMunicipal = Nothing Then
                                    sInscricaoMunTomador = ""
                                Else
                                    sInscricaoMunTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.IdentificacaoTomador.InscricaoMunicipal
                                End If
                            End If
#Else
                            If objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador Is Nothing Then
                                sCNPJCPFTomador = ""
                                sInscricaoMunTomador = ""
                            Else
                                sCNPJCPFTomador = objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.CpfCnpj.Item
                                If objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal = Nothing Then
                                    sInscricaoMunTomador = ""
                                Else
                                    sInscricaoMunTomador = objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal
                                End If
                            End If
#End If
                            sErro = "131"
                            sMsg1 = "conulta lote nfse"
#If ABRASF2 Then

                            If objCompNfse.Nfse.InfNfse.DataEmissao = #12:00:00 AM# Then
                                objCompNfse.Nfse.InfNfse.DataEmissao = Now.Date
                            End If

#Else
                            If objCompNfse.Nfse.InfNfse.DataEmissaoRps = #12:00:00 AM# Then
                                objCompNfse.Nfse.InfNfse.DataEmissaoRps = Now.Date
                            End If

#End If
                            sCodigoVerificacao = ""
                            sCompetencia = ""
                            dtDataEmissao = Now.Date
                            dtDataEmissaoRPS = Now.Date
                            sbTipo = 0
                            sSerie = ""
                            sNumeroRPS = ""
                            sbNaturezaOperacao = 0
                            sNumero = ""
                            sbOptanteSimplesNacional = 0
                            sUFGerador = ""
                            sCNPJPrestador = ""
                            sRazaoSocialPrestador = ""
                            sDiscriminacao = ""
                            sItemListaServico = ""
                            sInscricaoMunicipalPrestador = ""

                            sErro = "133"
                            sMsg1 = "conulta lote nfse"

                            If Not objCompNfse.Nfse.InfNfse.CodigoVerificacao = Nothing Then
                                sCodigoVerificacao = objCompNfse.Nfse.InfNfse.CodigoVerificacao
                            End If

                            sErro = "134"
                            sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Competencia = Nothing Then
                                sCompetencia = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Competencia
                            End If
#Else
                            If Not objCompNfse.Nfse.InfNfse.Competencia = Nothing Then
                                sCompetencia = objCompNfse.Nfse.InfNfse.Competencia
                            End If
#End If

                            sErro = "135"
                            sMsg1 = "conulta lote nfse"
                            If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

                            If Not objCompNfse.Nfse.InfNfse.DataEmissao = Nothing Then
                                dtDataEmissao = objCompNfse.Nfse.InfNfse.DataEmissao
                            End If

                            sErro = "136"
                            sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.DataEmissao = Nothing Then
                                dtDataEmissaoRPS = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.DataEmissao
                            End If
#Else
                            If Not objCompNfse.Nfse.InfNfse.DataEmissaoRps = Nothing Then
                                dtDataEmissaoRPS = objCompNfse.Nfse.InfNfse.DataEmissaoRps
                            End If
#End If
                            sErro = "137"
                            sMsg1 = "conulta lote nfse"

#If ABRASF2 Then

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Tipo = Nothing Then
                                sbTipo = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Tipo
                            End If
#Else
                            If Not objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo = Nothing Then
                                sbTipo = objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo
                            End If
#End If

                            sErro = "138"
                            sMsg1 = "conulta lote nfse"
#If ABRASF2 Then

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Serie = Nothing Then
                                sSerie = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Serie
                            End If
#Else
                            If Not objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie = Nothing Then
                                sSerie = objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie
                            End If
#End If

                            sErro = "139"
                            sMsg1 = "conulta lote nfse"

#If ABRASF2 Then

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Numero = Nothing Then
                                sNumeroRPS = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Numero
                            End If
#Else
                            If Not objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero = Nothing Then
                                sNumeroRPS = objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero
                            End If
#End If


                            sErro = "140"
                            sMsg1 = "conulta lote nfse"

#If Not ABRASF2 Then
                            If Not objCompNfse.Nfse.InfNfse.NaturezaOperacao = Nothing Then
                                sbNaturezaOperacao = objCompNfse.Nfse.InfNfse.NaturezaOperacao
                            End If
#End If

                            sErro = "141"
                            sMsg1 = "conulta lote nfse"

                            If Not objCompNfse.Nfse.InfNfse.Numero = Nothing Then
                                sNumero = objCompNfse.Nfse.InfNfse.Numero
                            End If

                            sErro = "142"
                            sMsg1 = "conulta lote nfse"


#If ABRASF2 Then

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.OptanteSimplesNacional = Nothing Then
                                sbOptanteSimplesNacional = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.OptanteSimplesNacional
                            End If

#Else
                            If Not objCompNfse.Nfse.InfNfse.OptanteSimplesNacional = Nothing Then
                                sbOptanteSimplesNacional = objCompNfse.Nfse.InfNfse.OptanteSimplesNacional
                            End If
#End If


                            sErro = "143"
                            sMsg1 = "conulta lote nfse"

                            If Not objCompNfse.Nfse.InfNfse.OrgaoGerador.Uf = Nothing Then
                                sUFGerador = objCompNfse.Nfse.InfNfse.OrgaoGerador.Uf
                            End If

                            sErro = "144"
                            sMsg1 = "conulta lote nfse"



#If ABRASF2 Then
                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Prestador.CpfCnpj Is Nothing Then
                                sCNPJPrestador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Prestador.CpfCnpj.Item
                            End If

#Else
                            If Not objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.Cnpj = Nothing Then
                                sCNPJPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.Cnpj
                            End If
#End If


                            sErro = "145"
                            sMsg1 = "conulta lote nfse"


#If Not ABRASF2 Then
                            If Not objCompNfse.Nfse.InfNfse.PrestadorServico.RazaoSocial = Nothing Then
                                sRazaoSocialPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.RazaoSocial
                            End If
#End If


                            sErro = "146"
                            sMsg1 = "conulta lote nfse"


#If ABRASF2 Then
                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Discriminacao = Nothing Then
                                sDiscriminacao = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Discriminacao
                            End If

#Else
                            If Not objCompNfse.Nfse.InfNfse.Servico.Discriminacao = Nothing Then
                                sDiscriminacao = objCompNfse.Nfse.InfNfse.Servico.Discriminacao
                            End If
#End If



                            sErro = "147"
                            sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.ItemListaServico = Nothing Then
                                sItemListaServico = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.ItemListaServico
                            End If
#Else
                            If Not objCompNfse.Nfse.InfNfse.Servico.Discriminacao = Nothing Then
                                sDiscriminacao = objCompNfse.Nfse.InfNfse.Servico.Discriminacao
                            End If
#End If


                            sErro = "148"
                            sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Prestador.InscricaoMunicipal = Nothing Then
                                sInscricaoMunicipalPrestador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Prestador.InscricaoMunicipal
                            End If
#Else
                            If Not objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.InscricaoMunicipal = Nothing Then
                                sInscricaoMunicipalPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.InscricaoMunicipal
                            End If
#End If




                            sBairroPrestador = ""
                            lCepPrestador = 0
                            lCodMunPrestador = 0
                            sComplementoPrestador = ""
                            sEnderecoPrestador = ""
                            sNumeroPrestador = ""
                            sUFPrestador = ""
                            sNomeFantasiaPrestador = ""
                            sbRegimeEspecialTributacao = 0
                            lCodigoCnae = 0
                            lCodigoMunicipioServico = 0
                            sCodigoTributacaoMunicipio = ""
                            sBairroTomador = ""
                            lCepTomador = 0
                            lCodigoMunicipioTomador = 0
                            sComplementoTomador = ""
                            sEnderecoTomador = ""
                            sNumeroTomador = ""
                            sUFTomador = ""
                            sRazaoSocialTomador = ""
                            lCodigoMunicipioGerador = 0

                            sErro = "149"
                            sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                            If Not objCompNfse.Nfse.InfNfse Is Nothing Then
#Else
                            If Not objCompNfse.Nfse.InfNfse.PrestadorServico Is Nothing Then
#End If

                                sErro = "150"
                                sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                                If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico Is Nothing Then
#Else
                                If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing Then
#End If

                                    sErro = "151"
                                    sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Bairro = Nothing Then
                                        sBairroPrestador = objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Bairro
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Bairro = Nothing Then
                                        sBairroPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Bairro
                                    End If
#End If

                                    sErro = "152"
                                    sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Cep = Nothing Then
                                        lCepPrestador = objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Cep
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Cep = Nothing Then
                                        lCepPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Cep
                                    End If
#End If


                                    sErro = "153"
                                    sMsg1 = "conulta lote nfse"


#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.CodigoMunicipio = Nothing Then
                                        lCodMunPrestador = objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.CodigoMunicipio
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.CodigoMunicipio = Nothing Then
                                        lCodMunPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.CodigoMunicipio
                                    End If
#End If


                                    sErro = "154"
                                    sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Complemento = Nothing Then
                                        sComplementoPrestador = objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Complemento
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Complemento = Nothing Then
                                        sComplementoPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Complemento
                                    End If
#End If

                                    sErro = "155"
                                    sMsg1 = "conulta lote nfse"


#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Endereco = Nothing Then
                                        sEnderecoPrestador = objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Endereco
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Endereco = Nothing Then
                                        sEnderecoPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Endereco
                                    End If
#End If


                                    sErro = "156"
                                    sMsg1 = "conulta lote nfse"


#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Numero = Nothing Then
                                        sNumeroPrestador = objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Numero
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Numero = Nothing Then
                                        sNumeroPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Numero
                                    End If
#End If

                                    sErro = "157"
                                    sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Uf = Nothing Then
                                        sUFPrestador = objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Uf
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Uf = Nothing Then
                                        sUFPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Uf
                                    End If
#End If

                                End If

                                sErro = "158"
                                sMsg1 = "conulta lote nfse"

#If Not ABRASF2 Then
                                If Not objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia = Nothing Then
                                    sNomeFantasiaPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia
                                End If
#End If


                            End If

                            sErro = "159"
                            sMsg1 = "conulta lote nfse"


#If ABRASF2 Then
                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.RegimeEspecialTributacao = Nothing Then
                                sbRegimeEspecialTributacao = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.RegimeEspecialTributacao
                            End If
#Else
                            If Not objCompNfse.Nfse.InfNfse.RegimeEspecialTributacao = Nothing Then
                                sbRegimeEspecialTributacao = objCompNfse.Nfse.InfNfse.RegimeEspecialTributacao
                            End If
#End If

                            sErro = "160"
                            sMsg1 = "conulta lote nfse"

                            If Not objCompNfse.Nfse.InfNfse.OrgaoGerador Is Nothing Then

                                sErro = "161"
                                sMsg1 = "conulta lote nfse"

                                If Not objCompNfse.Nfse.InfNfse.OrgaoGerador.CodigoMunicipio = Nothing Then
                                    lCodigoMunicipioGerador = objCompNfse.Nfse.InfNfse.OrgaoGerador.CodigoMunicipio
                                End If
                            End If

                            sErro = "162"
                            sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico Is Nothing Then
#Else
                            If Not objCompNfse.Nfse.InfNfse.Servico Is Nothing Then
#End If

                                sErro = "163"
                                sMsg1 = "conulta lote nfse"


#If ABRASF2 Then
                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.CodigoCnae = Nothing Then
                                    lCodigoCnae = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.CodigoCnae
                                End If
#Else
                                If Not objCompNfse.Nfse.InfNfse.Servico.CodigoCnae = Nothing Then
                                    lCodigoCnae = objCompNfse.Nfse.InfNfse.Servico.CodigoCnae
                                End If
#End If


                                sErro = "164"
                                sMsg1 = "conulta lote nfse"


#If ABRASF2 Then
                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.CodigoMunicipio = Nothing Then
                                    lCodigoMunicipioServico = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.CodigoMunicipio
                                End If
#Else
                                If Not objCompNfse.Nfse.InfNfse.Servico.CodigoMunicipio = Nothing Then
                                    lCodigoMunicipioServico = objCompNfse.Nfse.InfNfse.Servico.CodigoMunicipio
                                End If
#End If

                                sErro = "165"
                                sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.CodigoTributacaoMunicipio = Nothing Then
                                    sCodigoTributacaoMunicipio = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.CodigoTributacaoMunicipio
                                End If
#Else
                                If Not objCompNfse.Nfse.InfNfse.Servico.CodigoTributacaoMunicipio = Nothing Then
                                    sCodigoTributacaoMunicipio = objCompNfse.Nfse.InfNfse.Servico.CodigoTributacaoMunicipio
                                End If
#End If

                            End If

                            sErro = "166"
                            sMsg1 = "conulta lote nfse"


#If ABRASF2 Then
                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador Is Nothing Then
#Else
                            If Not objCompNfse.Nfse.InfNfse.TomadorServico Is Nothing Then
#End If

                                sErro = "167"
                                sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco Is Nothing Then
#Else
                                If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco Is Nothing Then
#End If

                                    sErro = "168"
                                    sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Bairro = Nothing Then
                                        sBairroTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Bairro
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Bairro = Nothing Then
                                        sBairroTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Bairro
                                    End If
#End If

                                    sErro = "169"
                                    sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Cep = Nothing Then
                                        lCepTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Cep
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Cep = Nothing Then
                                        lCepTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Cep
                                    End If
#End If

                                    sErro = "170"
                                    sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.CodigoMunicipio = Nothing Then
                                        lCodigoMunicipioTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.CodigoMunicipio
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.CodigoMunicipio = Nothing Then
                                        lCodigoMunicipioTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.CodigoMunicipio
                                    End If
#End If

                                    sErro = "171"
                                    sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Complemento = Nothing Then
                                        sComplementoTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Complemento
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento = Nothing Then
                                        sComplementoTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento
                                    End If
#End If

                                    sErro = "172"
                                    sMsg1 = "conulta lote nfse"


#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Endereco = Nothing Then
                                        sEnderecoTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Endereco
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Endereco = Nothing Then
                                        sEnderecoTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Endereco
                                    End If
#End If
                                    sErro = "173"
                                    sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Numero = Nothing Then
                                        sNumeroTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Numero
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Numero = Nothing Then
                                        sNumeroTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Numero
                                    End If
#End If

                                    sErro = "174"
                                    sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Uf = Nothing Then
                                        sUFTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Uf
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Uf = Nothing Then
                                        sUFTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Uf
                                    End If
#End If

                                    sErro = "175"
                                    sMsg1 = "conulta lote nfse"

#If ABRASF2 Then
                                    If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.RazaoSocial = Nothing Then
                                        sRazaoSocialTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.RazaoSocial
                                    End If
#Else
                                    If Not objCompNfse.Nfse.InfNfse.TomadorServico.RazaoSocial = Nothing Then
                                        sRazaoSocialTomador = objCompNfse.Nfse.InfNfse.TomadorServico.RazaoSocial
                                    End If
#End If


                                End If
                            End If

                            dAliquota = 0
                            dBaseCalculo = 0
                            dDescontoCondicionado = 0
                            dDescontoIncondicionado = 0
                            dISSRetido = 0
                            dOutrasRetencoes = 0
                            dValorCofins = 0
                            dValorCsll = 0
                            dValorDeducoes = 0
                            dValorInss = 0
                            dValorIr = 0
                            dValorIss = 0
                            dValorIssRetido = 0
                            dValorLiquidoNfse = 0
                            dValorPis = 0
                            dValorServicos = 0
                            dValorCredito = 0

#If ABRASF2 Then
                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.Aliquota = Nothing Then
                                dAliquota = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.Aliquota, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.ValoresNfse.BaseCalculo = Nothing Then
                                dBaseCalculo = Replace(objCompNfse.Nfse.InfNfse.ValoresNfse.BaseCalculo, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.DescontoCondicionado = Nothing Then
                                dDescontoCondicionado = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.DescontoCondicionado, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.DescontoIncondicionado = Nothing Then
                                dDescontoIncondicionado = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.DescontoIncondicionado, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.IssRetido = Nothing Then
                                dISSRetido = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.IssRetido, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.OutrasRetencoes = Nothing Then
                                dOutrasRetencoes = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.OutrasRetencoes, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorCofins = Nothing Then
                                dValorCofins = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorCofins, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorCsll = Nothing Then
                                dValorCsll = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorCsll, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorDeducoes = Nothing Then
                                dValorDeducoes = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorDeducoes, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorInss = Nothing Then
                                dValorInss = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorInss, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorIr = Nothing Then
                                dValorIr = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorIr, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorIss = Nothing Then
                                dValorIss = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorIss, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.ValoresNfse.ValorLiquidoNfse = Nothing Then
                                dValorLiquidoNfse = Replace(objCompNfse.Nfse.InfNfse.ValoresNfse.ValorLiquidoNfse, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorPis = Nothing Then
                                dValorPis = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorPis, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorServicos = Nothing Then
                                dValorServicos = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorServicos, ".", ",")
                            End If

#Else
                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.Aliquota = Nothing Then
                                dAliquota = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.Aliquota, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.BaseCalculo = Nothing Then
                                dBaseCalculo = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.BaseCalculo, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoCondicionado = Nothing Then
                                dDescontoCondicionado = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoCondicionado, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoIncondicionado = Nothing Then
                                dDescontoIncondicionado = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoIncondicionado, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.IssRetido = Nothing Then
                                dISSRetido = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.IssRetido, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.OutrasRetencoes = Nothing Then
                                dOutrasRetencoes = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.OutrasRetencoes, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCofins = Nothing Then
                                dValorCofins = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCofins, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCsll = Nothing Then
                                dValorCsll = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCsll, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorDeducoes = Nothing Then
                                dValorDeducoes = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorDeducoes, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorInss = Nothing Then
                                dValorInss = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorInss, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIr = Nothing Then
                                dValorIr = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIr, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIss = Nothing Then
                                dValorIss = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIss, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido = Nothing Then
                                dValorIssRetido = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorLiquidoNfse = Nothing Then
                                dValorLiquidoNfse = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorLiquidoNfse, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorPis = Nothing Then
                                dValorPis = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorPis, ".", ",")
                            End If

                            If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorServicos = Nothing Then
                                dValorServicos = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorServicos, ".", ",")
                            End If

#End If
                            If Not objCompNfse.Nfse.InfNfse.ValorCredito = Nothing Then
                                dValorCredito = Replace(objCompNfse.Nfse.InfNfse.ValorCredito, ".", ",")
                            End If

                            'If objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing Then


                            '    sErro = "149"
                            '    sMsg1 = "vai gravar RPSWEBProt"

                            '    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBProt ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, data, hora, Versao, CodigoVerificacao, Competencia, DataEmissao, DataEmissaoRPS, Id,  TipoRPS, SerieRPS, NumeroRPS, NaturezaOperacao, Numero, OptanteSimplesNacional, OrgaoGeradorCodMun, OrgaoGeradorUF, " & _
                            '    "OutrasInformacoes, ContatoEmail, ContatoTelefone, PrestadorBairro, PrestadorCEP, PrestadorCodMun, PrestadorComplemento, PrestadorEndereco, PrestadorNumero, PrestadorUF, PrestadorCNPJ, PrestadorInscMun, PrestadorNomeFantasia, PrestadorRazaoSocial, " & _
                            '    "RegimeEspecialTributaco, ServicoCodigoCNAE,ServicoCodMun, ServicoCodTrib, ServicoDiscriminacao, ServicoItemLista, Aliquota, BaseCalculo, DescontoCondicionado, DescontoIncondicionado, ISSRetido, OutrasRetencoes, " & _
                            '    "ValorCofins, ValorCsll, ValorDeducoes, ValorINSS, ValorIR, ValorISS, ValorISSRetido, ValorLiquidoNfse, ValorPIS, ValorServicos,  TomadorEmail, TomadorTelefone, TomadorBairro, TomadorCep, TomadorCodMun, " & _
                            '    "TomadorComplemento, TomandorEndereco, TomadorNumero, TomadorUF, TomadorCPFCNPJ, TomadorInscMun, TomadorRazaoSocial, ValorCredito, NumIntNF)  VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, " & _
                            '    "{20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, {59}, {60}, {61}, {62}, {63}, {64}, {65}, {66}, {67}, {68}, {69}, {70})", _
                            '    lRPSWEBProtNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, Now.Date, TimeOfDay.ToOADate, "", sCodigoVerificacao, dtCompetencia, dtDataEmissao, dtDataEmissaoRPS, IIf(objCompNfse.Nfse.InfNfse.Id = Nothing, "", objCompNfse.Nfse.InfNfse.Id), sbTipo, sSerie, sNumeroRPS, _
                            '    sbNaturezaOperacao, sNumero, sbOptanteSimplesNacional, objCompNfse.Nfse.InfNfse.OrgaoGerador.CodigoMunicipio, sUFGerador, IIf(objCompNfse.Nfse.InfNfse.OutrasInformacoes = Nothing, "", objCompNfse.Nfse.InfNfse.OutrasInformacoes), _
                            '    sEmailPrestador, sTelefonePrestador, "", 0, 0, "", "", _
                            '    "", "", sCNPJPrestador, sInscricaoMunicipalPrestador, sNomeFantasiaPrestador, sRazaoSocialPrestador, _
                            '    IIf(objCompNfse.Nfse.InfNfse.RegimeEspecialTributacao = Nothing, "", objCompNfse.Nfse.InfNfse.RegimeEspecialTributacao), IIf(objCompNfse.Nfse.InfNfse.Servico.CodigoCnae = Nothing, "", objCompNfse.Nfse.InfNfse.Servico.CodigoCnae), objCompNfse.Nfse.InfNfse.Servico.CodigoMunicipio, IIf(objCompNfse.Nfse.InfNfse.Servico.CodigoTributacaoMunicipio Is Nothing, "", objCompNfse.Nfse.InfNfse.Servico.CodigoTributacaoMunicipio), sDiscriminacao, _
                            '    sItemListaServico, objCompNfse.Nfse.InfNfse.Servico.Valores.Aliquota, objCompNfse.Nfse.InfNfse.Servico.Valores.BaseCalculo, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoCondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoIncondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.IssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.OutrasRetencoes, _
                            '    objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCofins, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCsll, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorDeducoes, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorInss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIr, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorLiquidoNfse, _
                            '    objCompNfse.Nfse.InfNfse.Servico.Valores.ValorPis, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorServicos, sEmailTomador, sTelefoneTomador, IIf(objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Bairro = Nothing, "", objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Bairro), objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Cep, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.CodigoMunicipio, _
                            '    IIf(objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento = Nothing, "", objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento), objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Endereco, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Numero, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Uf, sCNPJCPFTomador, sInscricaoMunTomador, _
                            '    objCompNfse.Nfse.InfNfse.TomadorServico.RazaoSocial, objCompNfse.Nfse.InfNfse.ValorCredito, lNumIntNFiscal)


                            'Else


                            sErro = "176"
                            sMsg1 = "vai gravar RPSWEBProt"
                            If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

                            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBProt ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, data, hora, Versao, CodigoVerificacao, Competencia, DataEmissao, DataEmissaoRPS, Id,  TipoRPS, SerieRPS, NumeroRPS, NaturezaOperacao, Numero, OptanteSimplesNacional, OrgaoGeradorCodMun, OrgaoGeradorUF, " & _
                            "OutrasInformacoes, ContatoEmail, ContatoTelefone, PrestadorBairro, PrestadorCEP, PrestadorCodMun, PrestadorComplemento, PrestadorEndereco, PrestadorNumero, PrestadorUF, PrestadorCNPJ, PrestadorInscMun, PrestadorNomeFantasia, PrestadorRazaoSocial, " & _
                            "RegimeEspecialTributaco, ServicoCodigoCNAE,ServicoCodMun, ServicoCodTrib, ServicoDiscriminacao, ServicoItemLista, Aliquota, BaseCalculo, DescontoCondicionado, DescontoIncondicionado, ISSRetido, OutrasRetencoes, " & _
                            "ValorCofins, ValorCsll, ValorDeducoes, ValorINSS, ValorIR, ValorISS, ValorISSRetido, ValorLiquidoNfse, ValorPIS, ValorServicos,  TomadorEmail, TomadorTelefone, TomadorBairro, TomadorCep, TomadorCodMun, " & _
                            "TomadorComplemento, TomandorEndereco, TomadorNumero, TomadorUF, TomadorCPFCNPJ, TomadorInscMun, TomadorRazaoSocial, ValorCredito, NumIntNF)  VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, " & _
                            "{20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, {59}, {60}, {61}, {62}, {63}, {64}, {65}, {66}, {67}, {68}, {69}, {70})", _
                            lRPSWEBProtNumIntDoc, iFilialEmpresa, iRPSAmbiente, Left(sLote, 15), Left(sProtocolo, 50), Now.Date, TimeOfDay.ToOADate, "", Left(sCodigoVerificacao, 9), Left(sCompetencia, 20), dtDataEmissao, dtDataEmissaoRPS, IIf(objCompNfse.Nfse.InfNfse.Id = Nothing, "", Left(objCompNfse.Nfse.InfNfse.Id, 255)), sbTipo, Left(sSerie, 5), Left(sNumeroRPS, 15), _
                            sbNaturezaOperacao, Left(sNumero, 15), sbOptanteSimplesNacional, lCodigoMunicipioGerador, Left(sUFGerador, 2), IIf(objCompNfse.Nfse.InfNfse.OutrasInformacoes = Nothing, "", Left(objCompNfse.Nfse.InfNfse.OutrasInformacoes, 255)), _
                            Left(sEmailPrestador, 80), Left(sTelefonePrestador, 11), Left(sBairroPrestador, 60), lCepPrestador, lCodMunPrestador, Left(sComplementoPrestador, 60), Left(sEnderecoPrestador, 125), _
                            Left(sNumeroPrestador, 10), Left(sUFPrestador, 2), Left(sCNPJPrestador, 14), Left(sInscricaoMunicipalPrestador, 15), Left(sNomeFantasiaPrestador, 60), Left(sRazaoSocialPrestador, 115), _
                            sbRegimeEspecialTributacao, lCodigoCnae, lCodigoMunicipioServico, Left(sCodigoTributacaoMunicipio, 20), Left(sDiscriminacao, 2000), _
                            Left(sItemListaServico, 5), Round(dAliquota, 2), Round(dBaseCalculo, 2), Round(dDescontoCondicionado, 2), Round(dDescontoIncondicionado, 2), Round(dISSRetido, 2), Round(dOutrasRetencoes, 2), _
                            Round(dValorCofins, 2), Round(dValorCsll, 2), Round(dValorDeducoes, 2), Round(dValorInss, 2), Round(dValorIr, 2), Round(dValorIss, 2), Round(dValorIssRetido, 2), Round(dValorLiquidoNfse, 2), _
                            Round(dValorPis, 2), Round(dValorServicos, 2), Left(sEmailTomador, 80), Left(sTelefoneTomador, 11), Left(sBairroTomador, 60), lCepTomador, lCodigoMunicipioTomador, _
                            Left(sComplementoTomador, 60), Left(sEnderecoTomador, 125), Left(sNumeroTomador, 10), Left(sUFTomador, 2), Left(sCNPJCPFTomador, 14), Left(sInscricaoMunTomador, 15), _
                            Left(sRazaoSocialTomador, 115), Round(dValorCredito, 2), lNumIntNFiscal)



                            'End If


                            lRPSWEBProtNumIntDoc = lRPSWEBProtNumIntDoc + 1

                        End If

                    Next

            End Select

Label_Fim:

            db1.Transaction.Commit()

        Catch ex As Exception When objRPSWEBRetEnvi Is Nothing

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7} )", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "o Lote provavelmente não chegou  a ser enviado", 0)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            Form1.Msg.Items.Add("ERRO - " & sErro & " - " & sMsg1)
            Form1.Msg.Items.Add("ERRO - o Lote provavelmente não chegou a ser enviado. Lote = " & CStr(lLote))

            If Form1.Msg.Items.Count - 15 < 1 Then
                Form1.Msg.TopIndex = 1
            Else
                Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
            End If

            Application.DoEvents()

            db1.Transaction.Rollback()

        Catch ex As Exception



            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Replace(Left(ex.Message, 255), "'", "*"), lNumIntNFiscal)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "ERRO - Encerrado a consulta deste lote", 0)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            Form1.Msg.Items.Add("ERRO - " & sErro & " - " & sMsg1)
            Form1.Msg.Items.Add("ERRO - a consulta do lote " & CStr(lLote) & " foi encerrado por erro. Erro = " & ex.Message)

            If Form1.Msg.Items.Count - 15 < 1 Then
                Form1.Msg.TopIndex = 1
            Else
                Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
            End If

            Application.DoEvents()

            db1.Transaction.Rollback()

        Finally
            Application.DoEvents()

            If lRPSWEBConsLoteNumIntDoc <> 0 Then
                iResult = db2.ExecuteCommand("UPDATE FatConfig Set Conteudo = {0} WHERE Codigo = {1}", lRPSWEBConsLoteNumIntDoc, "NUM_INT_PROX_RPSWEBCONSLOTE")
            End If

            If lRPSWEBProtNumIntDoc <> 0 Then
                iResult = db2.ExecuteCommand("UPDATE FatConfig Set Conteudo = {0} WHERE Codigo = {1}", lRPSWEBProtNumIntDoc, "NUM_INT_PROX_RPSWEBPROT")
            End If

            If lRPSWEBLoteLogNumIntDoc <> 0 Then
                iResult = db2.ExecuteCommand("UPDATE FatConfig Set Conteudo = {0} WHERE Codigo = {1}", lRPSWEBLoteLogNumIntDoc, "NUM_INT_PROX_RPSWEBLOTELOG")
            End If

            db2.Transaction.Commit()

            Form1.Msg.Items.Add("Transação Finalizada")

            If Form1.Msg.Items.Count - 15 < 1 Then
                Form1.Msg.TopIndex = 1
            Else
                Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
            End If

            db1.Connection.Close()
            db2.Connection.Close()
            dic.Connection.Close()


            db1.Dispose()
            dic.Dispose()
            db2.Dispose()

        End Try

#End If

    End Sub


    'Sub Formata_String_Numero(ByVal sStringRecebe As String, ByRef sStringRetorna As String)

    '    Dim iTamanho As Integer
    '    Dim sCaracter As String
    '    Dim iIndice As Integer

    '    iTamanho = Len(Trim(sStringRecebe))

    '    sStringRetorna = ""

    '    For iIndice = 1 To iTamanho

    '        sCaracter = Mid(sStringRecebe, iIndice, 1)

    '        If IsNumeric(sCaracter) Then
    '            sStringRetorna = sStringRetorna & sCaracter
    '        End If

    '    Next

    'End Sub

    Sub Formata_Sem_Espaco(ByVal sStringRecebe As String, ByRef sStringRetorna As String)

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

    Function DesacentuaTexto(ByVal sTexto As String) As String

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
#If Not TATUI2 Then
    Public Function ConsultaRPS(ByVal objConsNfseRpsEnvio As ConsultarNfseRpsEnvio, ByVal sCertificado As String, ByVal iRPSAmbiente As Integer, ByVal objEndFilial As Endereco, ByVal XMLStringCabec As String, ByVal db2 As SGEDadosDataContext, ByRef lRPSWEBConsLoteNumIntDoc As Long, ByRef lRPSWEBLoteLogNumIntDoc As Long, ByVal iFilialEmpresa As Integer, ByVal db1 As SGEDadosDataContext, ByVal sLote As String, ByVal sProtocolo As String, ByRef lRPSWEBProtNumIntDoc As Long, ByVal lLote As Long) As Long

        Dim XMLStreamRPS As MemoryStream = New MemoryStream(10000)
        Dim xm22 As Byte()
        Dim XMLString2 As String
        Dim certificado As Certificado = New Certificado
        Dim cert As X509Certificate2 = New X509Certificate2

        Dim homologacao As New NotaCariocaHomologacao.NfseSoapClient
        Dim producao As New NotaCariocaProducao.NfseSoapClient

        Dim homologacaobh = New GetWebRequest_bhHomologacao
        Dim producaobh = New GetWebRequest_bhProducao

        Dim homosalvadorconsrps As New br.gov.ba.salvador.sefaz.nfsehml2.ConsultaNfseRPS
        Dim prodsalvadorconsrps As New br.gov.ba.salvador.sefaz.nfse2.ConsultaNfseRPS

        Dim homologacaoginfes = New br.com.ginfes.homologacao.ServiceGinfesImplService
        Dim producaoginfes = New br.com.ginfes.producao.ServiceGinfesImplService

        Dim homsistema4rConsRPS = New br.com.sistemas4r.abrasf.nfsehml2.ConsultarNfsePorRps

        Dim prodsistema4rConsRPS As New br.com.sistemas4r.abrasf.nfse2.ConsultarNfsePorRps


        Dim xRetRPS As Byte()
        Dim XMLStreamRetRPS As MemoryStream = New MemoryStream(10000)
        Dim sAuxRPS As String

        Dim objListaMsgRetorno As ListaMensagemRetorno
        Dim objMsgRetorno As New tcMensagemRetorno

        Dim objCompNfse As tcCompNfse

        Dim iResult As Integer
        Dim iAchou As Integer
        Dim lNumIntNFiscal As Long

        Dim resRPSWEBProt As IEnumerable(Of RPSWEBProt)
        Dim resNFiscal As IEnumerable(Of NFiscal)

        Dim objNFiscal As NFiscal

        Dim resRPSWEBLote As IEnumerable(Of RPSWEBLote)

        Dim colNumIntNFiscal As Collection

        Dim mySerializerx3 As New XmlSerializer(GetType(ConsultarNfseRpsEnvio))

        Dim sEmailPrestador As String
        Dim sTelefonePrestador As String
        Dim sEmailTomador As String
        Dim sTelefoneTomador As String

        Dim sCPFCNPJTomador As String
        Dim sInscricaoMunicipalTomador As String
        Dim sRazaoSocialTomador As String
        Dim sCodigoVerificacao As String
        Dim dtDataEmissao As Date
        Dim sCompetencia As String
        Dim sbTipo As SByte
        Dim sSerie As String
        Dim sNumeroRPS As String
        Dim sNumero As String
        Dim sbNaturezaOperacao As SByte
        Dim sbOptanteSimplesNacional As SByte
        Dim sUFGerador As String
        Dim sCNPJPrestador As String
        Dim sRazaoSocialPrestador As String
        Dim sDiscriminacao As String
        Dim sItemListaServico As String
        Dim sInscricaoMunicipalPrestador As String

        Dim sBairroPrestador As String
        Dim lCepPrestador As Long
        Dim lCodMunPrestador As Long
        Dim sComplementoPrestador As String
        Dim sEnderecoPrestador As String
        Dim sNumeroPrestador As String
        Dim sUFPrestador As String
        Dim sNomeFantasiaPrestador As String
        Dim sbRegimeEspecialTributacao As SByte
        Dim lCodigoCnae As Long
        Dim lCodigoMunicipioServico As Long
        Dim sCodigoTributacaoMunicipio As String
        Dim sBairroTomador As String
        Dim lCepTomador As Long
        Dim lCodigoMunicipioTomador As Long
        Dim sComplementoTomador As String
        Dim sEnderecoTomador As String
        Dim sNumeroTomador As String
        Dim sUFTomador As String
        Dim lCodigoMunicipioGerador As Long

        Dim dAliquota As Double
        Dim dBaseCalculo As Double
        Dim dDescontoCondicionado As Double
        Dim dDescontoIncondicionado As Double
        Dim dISSRetido As Double
        Dim dOutrasRetencoes As Double
        Dim dValorCofins As Double
        Dim dValorCsll As Double
        Dim dValorDeducoes As Double
        Dim dValorInss As Double
        Dim dValorIr As Double
        Dim dValorIss As Double
        Dim dValorIssRetido As Double
        Dim dValorLiquidoNfse As Double
        Dim dValorPis As Double
        Dim dValorServicos As Double
        Dim dValorCredito As Double


        Dim sErro As String
        Dim sMsg1 As String
        Dim iDebug As Integer
        Dim sRetorno As String
        Dim iTamanho As Integer
        Dim dtDataEmissaoRPS As Date

        sErro = ""
        sMsg1 = ""

        Try



            iTamanho = 255
            sRetorno = StrDup(iTamanho, Chr(0))

            iDebug = 0

            Call GetPrivateProfileString("Geral", "Debug", 0, sRetorno, iTamanho, "Adm100.ini")

            If IsNumeric(sRetorno) Then
                iDebug = CInt(sRetorno)
            End If

            XMLStreamRPS = New MemoryStream(10000)
            mySerializerx3.Serialize(XMLStreamRPS, objConsNfseRpsEnvio)

            xm22 = XMLStreamRPS.ToArray

            XMLString2 = System.Text.Encoding.UTF8.GetString(xm22)

            XMLString2 = Mid(XMLString2, 1, 19) & " encoding=""utf-8"" " & Mid(XMLString2, 20)

            Dim XMLStringConsultarNfseRPSResposta As String


            sErro = "201"
            sMsg1 = "vai buscar o certificado"


            cert = certificado.BuscaNome(sCertificado)

            sErro = "202"
            sMsg1 = "vai conultar nfse por RPS"

            If iDebug = 1 Then MsgBox("202")

            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then

                XMLString2 = Replace(XMLString2, "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd", "http://www.abrasf.org.br/nfse.xsd")


                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homologacaobh.ClientCertificates.Add(cert)
                    XMLStringConsultarNfseRPSResposta = homologacaobh.ConsultarNfsePorRps(XMLStringCabec, XMLString2)
                Else
                    producaobh.ClientCertificates.Add(cert)
                    XMLStringConsultarNfseRPSResposta = producaobh.ConsultarNfsePorRps(XMLStringCabec, XMLString2)
                End If

                XMLStringConsultarNfseRPSResposta = Replace(XMLStringConsultarNfseRPSResposta, "http://www.abrasf.org.br/nfse.xsd", "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd")

            ElseIf UCase(objEndFilial.Cidade) = "SALVADOR" Then

                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homosalvadorconsrps.ClientCertificates.Add(cert)
                    XMLStringConsultarNfseRPSResposta = homosalvadorconsrps.ConsultarNfseRPS(XMLString2)
                Else
                    prodsalvadorconsrps.ClientCertificates.Add(cert)
                    XMLStringConsultarNfseRPSResposta = prodsalvadorconsrps.ConsultarNfseRPS(XMLString2)
                End If

            ElseIf UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then


                Dim AD2 As AssinaturaDigital = New AssinaturaDigital

                AD2.Assinar(XMLString2, "IdentificacaoRps", cert, objEndFilial.Cidade)

                Dim xMlD1 As XmlDocument

                xMlD1 = AD2.XMLDocAssinado()

                XMLString2 = AD2.XMLStringAssinado


                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homologacaoginfes.ClientCertificates.Add(cert)
                    XMLStringConsultarNfseRPSResposta = homologacaoginfes.ConsultarNfsePorRpsV3(XMLStringCabec, XMLString2)
                Else
                    producaoginfes.ClientCertificates.Add(cert)
                    XMLStringConsultarNfseRPSResposta = producaoginfes.ConsultarNfsePorRpsV3(XMLStringCabec, XMLString2)
                End If

            ElseIf UCase(objEndFilial.Cidade) = "TATUÍ" Then

                'Dim AD2 As AssinaturaDigital = New AssinaturaDigital

                'AD2.Assinar(XMLString2, "IdentificacaoRps", cert, objEndFilial.Cidade)

                'Dim xMlD1 As XmlDocument

                'xMlD1 = AD2.XMLDocAssinado()

                'XMLString2 = AD2.XMLStringAssinado


                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homsistema4rConsRPS.ClientCertificates.Add(cert)
                    XMLStringConsultarNfseRPSResposta = homsistema4rConsRPS.Execute(XMLString2)
                Else
                    prodsistema4rConsRPS.ClientCertificates.Add(cert)
                    XMLStringConsultarNfseRPSResposta = prodsistema4rConsRPS.Execute(XMLString2)
                End If



            Else

                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homologacao.ClientCredentials.ClientCertificate.Certificate = cert
                    XMLStringConsultarNfseRPSResposta = homologacao.ConsultarNfsePorRps(XMLString2)
                Else
                    producao.ClientCredentials.ClientCertificate.Certificate = cert
                    XMLStringConsultarNfseRPSResposta = producao.ConsultarNfsePorRps(XMLString2)
                End If

            End If

            If UCase(objEndFilial.Cidade) = "SALVADOR" Then
                XMLStringConsultarNfseRPSResposta = Replace(XMLStringConsultarNfseRPSResposta, "id=", "Id=")
            End If

            xRetRPS = System.Text.Encoding.UTF8.GetBytes(XMLStringConsultarNfseRPSResposta)

            XMLStreamRetRPS = New MemoryStream(10000)
            XMLStreamRetRPS.Write(xRetRPS, 0, xRetRPS.Length)

            Dim mySerializerConsultarNfseRPSResposta As New XmlSerializer(GetType(ConsultarNfseRpsResposta))

            Dim objConsultarNfseRPSResposta = New ConsultarNfseRpsResposta

            XMLStreamRetRPS.Position = 0

            sErro = "203"
            sMsg1 = "vai deserializar resposta conulta nfse por RPS"

            If iDebug = 1 Then MsgBox("203")

            If iDebug = 1 Then MsgBox(XMLStringConsultarNfseRPSResposta)


            objConsultarNfseRPSResposta = mySerializerConsultarNfseRPSResposta.Deserialize(XMLStreamRetRPS)

            sAuxRPS = objConsultarNfseRPSResposta.Item.GetType.ToString()
            If InStr(sAuxRPS, ".") <> 0 Then
                sAuxRPS = Mid(sAuxRPS, InStr(sAuxRPS, ".") + 1)
            End If

            Select Case sAuxRPS

                Case "ListaMensagemRetorno"

                    sErro = "204"
                    sMsg1 = "vai tratar ListaMensagemRetorno conulta nfse por RPS"

                    objListaMsgRetorno = objConsultarNfseRPSResposta.Item

                    For iIndiceRPS = 0 To objListaMsgRetorno.MensagemRetorno.Count - 1

                        objMsgRetorno = objListaMsgRetorno.MensagemRetorno(iIndiceRPS)

                        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Correcao, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9})", _
                        lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objMsgRetorno.Codigo, Left(objMsgRetorno.Mensagem, 200), Left(objMsgRetorno.Correcao, 200), Now.Date, TimeOfDay.ToOADate)

                        lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                        lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Retorno da consulta de RPS - Serie = & " & objConsNfseRpsEnvio.IdentificacaoRps.Serie & " NFiscal = " & objConsNfseRpsEnvio.IdentificacaoRps.Numero & " - " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao, 255), 0)

                        lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                        Form1.Msg.Items.Add("Retorno da consulta de RPS - Serie = & " & objConsNfseRpsEnvio.IdentificacaoRps.Serie & " NFiscal = " & objConsNfseRpsEnvio.IdentificacaoRps.Numero & " - " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao)

                        If Form1.Msg.Items.Count - 15 < 1 Then
                            Form1.Msg.TopIndex = 1
                        Else
                            Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                        End If

                        Application.DoEvents()

                    Next

                Case "tcCompNfse"

                    sErro = "205"
                    sMsg1 = "vai tratar tcCompNfse conulta nfse por RPS"

                    objCompNfse = objConsultarNfseRPSResposta.Item

                    resRPSWEBLote = db2.ExecuteQuery(Of RPSWEBLote) _
                    ("SELECT * FROM RPSWEBLote WHERE Lote = {0} ORDER BY NumIntNF", sLote)

                    colNumIntNFiscal = New Collection

                    For Each objRPSWEBLote In resRPSWEBLote
                        colNumIntNFiscal.Add(objRPSWEBLote.NumIntNF)
                    Next


                    iAchou = 0

                    sErro = "206"
                    sMsg1 = "vai tratar os dados da conulta nfse por RPS"

                    For iIndice2 = 1 To colNumIntNFiscal.Count
                        lNumIntNFiscal = colNumIntNFiscal(iIndice2)

                        resNFiscal = db1.ExecuteQuery(Of NFiscal) _
                        ("SELECT * FROM NFiscal WHERE NumINtDoc = {0}", lNumIntNFiscal)

                        objNFiscal = resNFiscal(0)

#If ABRASF2 Then
                        If ADM.Serie_Sem_E(objNFiscal.Serie) = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Serie And _
                               Format(objNFiscal.NumNotaFiscal, "000000000") = Format(CLng(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Numero), "000000000") Then
                            iAchou = 1
                            Exit For
                        End If

#Else
                        If ADM.Serie_Sem_E(objNFiscal.Serie) = objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie And _
                               Format(objNFiscal.NumNotaFiscal, "000000000") = Format(CLng(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero), "000000000") Then
                            iAchou = 1
                            Exit For
                        End If
#End If

                    Next

#If ABRASF2 Then
                    If iAchou = 0 Then
                        Throw New System.Exception("A nota consultada nao corresponde a nenhuma das enviadas. Serie Consulta = " & objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Serie & " Numero Consulta = " & Format(CLng(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Numero), "000000000"))
                    End If

                    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, tipo, serie, numero, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11})", _
                    lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, lLote, sProtocolo, 0, "Nota Fiscal de Servico gerada - " & objCompNfse.Nfse.InfNfse.Numero, objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Tipo, objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Serie, objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Numero, Now.Date, TimeOfDay.ToOADate)

                    lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                    lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Numero, 255), 0)

                    lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                    Form1.Msg.Items.Add("Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Numero)

#Else

                    If iAchou = 0 Then
                        Throw New System.Exception("A nota consultada nao corresponde a nenhuma das enviadas. Serie Consulta = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie & " Numero Consulta = " & Format(CLng(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero), "000000000"))
                    End If

                    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, tipo, serie, numero, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11})", _
                    lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, lLote, sProtocolo, 0, "Nota Fiscal de Servico gerada - " & objCompNfse.Nfse.InfNfse.Numero, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, Now.Date, TimeOfDay.ToOADate)

                    lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                    lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, 255), 0)

                    lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                    Form1.Msg.Items.Add("Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero)

#End If



                    If Form1.Msg.Items.Count - 15 < 1 Then
                        Form1.Msg.TopIndex = 1
                    Else
                        Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                    End If

                    Application.DoEvents()

                    sErro = "207"
                    sMsg1 = "vai ler a tabela RPSWEBProt conulta nfse por RPS"


                    '******** pega o ultimo retorno de envio do lote em questao *************
                    resRPSWEBProt = db2.ExecuteQuery(Of RPSWEBProt) _
                    ("SELECT * FROM RPSWEBProt WHERE FilialEmpresa = {0} AND NumIntNF = {1} AND Ambiente = {2}", iFilialEmpresa, lNumIntNFiscal, iRPSAmbiente)

                    If resRPSWEBProt.Count = 0 Then

                        sErro = "208"
                        sMsg1 = "conulta nfse por RPS"

#If ABRASF2 Then
                        sEmailPrestador = ""
                        sTelefonePrestador = ""

                        sErro = "209"
                        sMsg1 = "conulta nfse por RPS"
                        sEmailTomador = ""
                        sTelefoneTomador = ""

                        If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador Is Nothing Then
                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Contato Is Nothing Then
                                sEmailTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Contato.Email
                                If sEmailTomador = Nothing Then sEmailTomador = ""
                                sTelefoneTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Contato.Telefone
                                If sTelefoneTomador = Nothing Then sTelefoneTomador = ""
                            End If
                        End If
#Else
                        If objCompNfse.Nfse.InfNfse.PrestadorServico.Contato Is Nothing Then
                            sEmailPrestador = ""
                            sTelefonePrestador = ""
                        Else
                            sEmailPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Contato.Email
                            If sEmailPrestador = Nothing Then sEmailPrestador = ""
                            sTelefonePrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Contato.Telefone
                            If sTelefonePrestador = Nothing Then sTelefonePrestador = ""
                        End If

                        sErro = "209"
                        sMsg1 = "conulta nfse por RPS"

                        If objCompNfse.Nfse.InfNfse.TomadorServico.Contato Is Nothing Then
                            sEmailTomador = ""
                            sTelefoneTomador = ""
                        Else
                            sEmailTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Contato.Email
                            If sEmailTomador = Nothing Then sEmailTomador = ""
                            sTelefoneTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Contato.Telefone
                            If sTelefoneTomador = Nothing Then sTelefoneTomador = ""
                        End If

#End If


                            sCPFCNPJTomador = ""
                            sInscricaoMunicipalTomador = ""
                            sRazaoSocialTomador = ""
                            sCodigoVerificacao = ""
                            sCompetencia = ""
                            dtDataEmissao = Now.Date
                            sbTipo = 0
                            sSerie = ""
                            sNumeroRPS = ""
                            sbNaturezaOperacao = 0
                            sNumero = ""
                            sbOptanteSimplesNacional = 0
                            sUFGerador = ""
                            sCNPJPrestador = ""
                            sRazaoSocialPrestador = ""
                            sDiscriminacao = ""
                            sItemListaServico = ""
                            sInscricaoMunicipalPrestador = ""
                            dtDataEmissaoRPS = Now.Date

                            sErro = "218"
                            sMsg1 = "conulta nfse por RPS"

#If ABRASF2 Then

                        If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador Is Nothing Then

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.IdentificacaoTomador Is Nothing Then
                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.IdentificacaoTomador.CpfCnpj Is Nothing Then
                                    sCPFCNPJTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.IdentificacaoTomador.CpfCnpj.Item
                                End If

                                sErro = "219"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.IdentificacaoTomador.InscricaoMunicipal = Nothing Then
                                    sInscricaoMunicipalTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.IdentificacaoTomador.InscricaoMunicipal
                                End If
                            End If

                            sErro = "220"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.RazaoSocial = Nothing Then
                                sRazaoSocialTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.RazaoSocial
                            End If

                        End If
                        sErro = "221"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.CodigoVerificacao = Nothing Then
                            sCodigoVerificacao = objCompNfse.Nfse.InfNfse.CodigoVerificacao
                        End If

                        sErro = "222"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Competencia = Nothing Then
                            sCompetencia = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Competencia
                        End If

                        sErro = "223"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.DataEmissao = Nothing Then
                            dtDataEmissao = objCompNfse.Nfse.InfNfse.DataEmissao
                        End If

                        sErro = "224"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Tipo = Nothing Then
                            sbTipo = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Tipo
                        End If

                        sErro = "225"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Serie = Nothing Then
                            sSerie = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Serie
                        End If

                        sErro = "226"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Numero = Nothing Then
                            sNumeroRPS = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.IdentificacaoRps.Numero
                        End If

                        sErro = "227"
                        sMsg1 = "conulta nfse por RPS"

                        sErro = "228"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.Numero = Nothing Then
                            sNumero = objCompNfse.Nfse.InfNfse.Numero
                        End If

                        sErro = "229"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.OptanteSimplesNacional = Nothing Then
                            sbOptanteSimplesNacional = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.OptanteSimplesNacional
                        End If

                        sErro = "230"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.OrgaoGerador.Uf = Nothing Then
                            sUFGerador = objCompNfse.Nfse.InfNfse.OrgaoGerador.Uf
                        End If

                        sErro = "231"
                        sMsg1 = "conulta nfse por RPS"

                        sErro = "232"
                        sMsg1 = "conulta nfse por RPS"


                        sErro = "233"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico Is Nothing Then
                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Discriminacao = Nothing Then
                                sDiscriminacao = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Discriminacao
                            End If

                            sErro = "234"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.ItemListaServico = Nothing Then
                                sItemListaServico = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.ItemListaServico
                            End If

                        End If

                        If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Prestador Is Nothing Then
                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Prestador.CpfCnpj Is Nothing Then
                                sCNPJPrestador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Prestador.CpfCnpj.Item
                            End If


                            sErro = "235"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Prestador.InscricaoMunicipal = Nothing Then
                                sInscricaoMunicipalPrestador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Prestador.InscricaoMunicipal
                            End If

                        End If

                        If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.DataEmissao = Nothing Then
                            dtDataEmissaoRPS = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Rps.DataEmissao
                        End If

#Else

                        If Not objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador Is Nothing Then
                            If Not objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.CpfCnpj Is Nothing Then
                                sCPFCNPJTomador = objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.CpfCnpj.Item
                            End If

                            sErro = "219"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal = Nothing Then
                                sInscricaoMunicipalTomador = objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal
                            End If
                        End If

                        sErro = "220"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.TomadorServico.RazaoSocial = Nothing Then
                            sRazaoSocialTomador = objCompNfse.Nfse.InfNfse.TomadorServico.RazaoSocial
                        End If

                        sErro = "221"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.CodigoVerificacao = Nothing Then
                            sCodigoVerificacao = objCompNfse.Nfse.InfNfse.CodigoVerificacao
                        End If

                        sErro = "222"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.Competencia = Nothing Then
                            sCompetencia = objCompNfse.Nfse.InfNfse.Competencia
                        End If

                        sErro = "223"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.DataEmissao = Nothing Then
                            dtDataEmissao = objCompNfse.Nfse.InfNfse.DataEmissao
                        End If

                        sErro = "224"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo = Nothing Then
                            sbTipo = objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo
                        End If

                        sErro = "225"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie = Nothing Then
                            sSerie = objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie
                        End If

                        sErro = "226"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero = Nothing Then
                            sNumeroRPS = objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero
                        End If

                        sErro = "227"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.NaturezaOperacao = Nothing Then
                            sbNaturezaOperacao = objCompNfse.Nfse.InfNfse.NaturezaOperacao
                        End If

                        sErro = "228"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.Numero = Nothing Then
                            sNumero = objCompNfse.Nfse.InfNfse.Numero
                        End If

                        sErro = "229"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.OptanteSimplesNacional = Nothing Then
                            sbOptanteSimplesNacional = objCompNfse.Nfse.InfNfse.OptanteSimplesNacional
                        End If

                        sErro = "230"
                        sMsg1 = "conulta nfse por RPS"


                        If Not objCompNfse.Nfse.InfNfse.OrgaoGerador Is Nothing Then
                            If Not objCompNfse.Nfse.InfNfse.OrgaoGerador.Uf = Nothing Then
                                sUFGerador = objCompNfse.Nfse.InfNfse.OrgaoGerador.Uf
                            End If
                        End If

                        sErro = "231"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.Cnpj = Nothing Then
                            sCNPJPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.Cnpj
                        End If

                        sErro = "232"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.PrestadorServico.RazaoSocial = Nothing Then
                            sRazaoSocialPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.RazaoSocial
                        End If

                        sErro = "233"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.Servico.Discriminacao = Nothing Then
                            sDiscriminacao = objCompNfse.Nfse.InfNfse.Servico.Discriminacao
                        End If

                        sErro = "234"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.Servico.ItemListaServico = Nothing Then
                            sItemListaServico = objCompNfse.Nfse.InfNfse.Servico.ItemListaServico
                        End If

                        sErro = "235"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.InscricaoMunicipal = Nothing Then
                            sInscricaoMunicipalPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.InscricaoMunicipal
                        End If

                        If Not objCompNfse.Nfse.InfNfse.DataEmissaoRps = Nothing Then
                            dtDataEmissaoRPS = objCompNfse.Nfse.InfNfse.DataEmissaoRps
                        End If


#End If


                        sBairroPrestador = ""
                        lCepPrestador = 0
                        lCodMunPrestador = 0
                        sComplementoPrestador = ""
                        sEnderecoPrestador = ""
                        sNumeroPrestador = ""
                        sUFPrestador = ""
                        sNomeFantasiaPrestador = ""
                        sbRegimeEspecialTributacao = 0
                        lCodigoCnae = 0
                        lCodigoMunicipioServico = 0
                        sCodigoTributacaoMunicipio = ""
                        sBairroTomador = ""
                        lCepTomador = 0
                        lCodigoMunicipioTomador = 0
                        sComplementoTomador = ""
                        sEnderecoTomador = ""
                        sNumeroTomador = ""
                        sUFTomador = ""
                        sRazaoSocialTomador = ""
                        lCodigoMunicipioGerador = 0



                        sErro = "236"
                        sMsg1 = "conulta nfse por RPS"

#If ABRASF2 Then

                        If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico Is Nothing Then

                            sErro = "237"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Endereco Is Nothing Then

                                sErro = "238"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Bairro = Nothing Then
                                    sBairroPrestador = objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Bairro
                                End If

                                sErro = "239"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Cep = Nothing Then
                                    lCepPrestador = objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Cep
                                End If

                                sErro = "240"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.CodigoMunicipio = Nothing Then
                                    lCodMunPrestador = objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.CodigoMunicipio
                                End If

                                sErro = "241"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Complemento = Nothing Then
                                    sComplementoPrestador = objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Complemento
                                End If

                                sErro = "242"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Endereco = Nothing Then
                                    sEnderecoPrestador = objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Endereco
                                End If

                                sErro = "243"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Numero = Nothing Then
                                    sNumeroPrestador = objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Numero
                                End If

                                sErro = "244"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Uf = Nothing Then
                                    sUFPrestador = objCompNfse.Nfse.InfNfse.EnderecoPrestadorServico.Uf
                                End If
                            End If

                            sErro = "245"
                            sMsg1 = "conulta nfse por RPS"

                            'If Not objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia = Nothing Then
                            '    sNomeFantasiaPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia
                            'End If


                        End If

                        sErro = "246"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.RegimeEspecialTributacao = Nothing Then
                            sbRegimeEspecialTributacao = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.RegimeEspecialTributacao
                        End If

                        sErro = "247"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.OrgaoGerador Is Nothing Then

                            sErro = "248"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.OrgaoGerador.CodigoMunicipio = Nothing Then
                                lCodigoMunicipioGerador = objCompNfse.Nfse.InfNfse.OrgaoGerador.CodigoMunicipio
                            End If
                        End If

                        sErro = "249"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico Is Nothing Then

                            sErro = "250"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.CodigoCnae = Nothing Then
                                lCodigoCnae = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.CodigoCnae
                            End If

                            sErro = "251"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.CodigoMunicipio = Nothing Then
                                lCodigoMunicipioServico = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.CodigoMunicipio
                            End If

                            sErro = "252"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.CodigoTributacaoMunicipio = Nothing Then
                                sCodigoTributacaoMunicipio = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.CodigoTributacaoMunicipio
                            End If
                        End If

                        sErro = "253"
                        sMsg1 = "conulta nfse por RPS"


                        If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador Is Nothing Then

                            sErro = "254"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco Is Nothing Then

                                sErro = "255"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Bairro = Nothing Then
                                    sBairroTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Bairro
                                End If

                                sErro = "256"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Cep = Nothing Then
                                    lCepTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Cep
                                End If

                                sErro = "257"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.CodigoMunicipio = Nothing Then
                                    lCodigoMunicipioTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.CodigoMunicipio
                                End If

                                sErro = "258"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Complemento = Nothing Then
                                    sComplementoTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Complemento
                                End If

                                sErro = "259"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Endereco = Nothing Then
                                    sEnderecoTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Endereco
                                End If

                                sErro = "260"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Numero = Nothing Then
                                    sNumeroTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Numero
                                End If

                                sErro = "261"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Uf = Nothing Then
                                    sUFTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.Endereco.Uf
                                End If

                                sErro = "262"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.RazaoSocial = Nothing Then
                                    sRazaoSocialTomador = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Tomador.RazaoSocial
                                End If


                            End If
                        End If


#Else

                        If Not objCompNfse.Nfse.InfNfse.PrestadorServico Is Nothing Then

                            sErro = "237"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing Then

                                sErro = "238"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Bairro = Nothing Then
                                    sBairroPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Bairro
                                End If

                                sErro = "239"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Cep = Nothing Then
                                    lCepPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Cep
                                End If

                                sErro = "240"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.CodigoMunicipio = Nothing Then
                                    lCodMunPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.CodigoMunicipio
                                End If

                                sErro = "241"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Complemento = Nothing Then
                                    sComplementoPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Complemento
                                End If

                                sErro = "242"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Endereco = Nothing Then
                                    sEnderecoPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Endereco
                                End If

                                sErro = "243"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Numero = Nothing Then
                                    sNumeroPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Numero
                                End If

                                sErro = "244"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Uf = Nothing Then
                                    sUFPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Uf
                                End If
                            End If

                            sErro = "245"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia = Nothing Then
                                sNomeFantasiaPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia
                            End If


                        End If

                        sErro = "246"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.RegimeEspecialTributacao = Nothing Then
                            sbRegimeEspecialTributacao = objCompNfse.Nfse.InfNfse.RegimeEspecialTributacao
                        End If

                        sErro = "247"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.OrgaoGerador Is Nothing Then

                            sErro = "248"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.OrgaoGerador.CodigoMunicipio = Nothing Then
                                lCodigoMunicipioGerador = objCompNfse.Nfse.InfNfse.OrgaoGerador.CodigoMunicipio
                            End If
                        End If

                        sErro = "249"
                        sMsg1 = "conulta nfse por RPS"

                        If Not objCompNfse.Nfse.InfNfse.Servico Is Nothing Then

                            sErro = "250"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.Servico.CodigoCnae = Nothing Then
                                lCodigoCnae = objCompNfse.Nfse.InfNfse.Servico.CodigoCnae
                            End If

                            sErro = "251"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.Servico.CodigoMunicipio = Nothing Then
                                lCodigoMunicipioServico = objCompNfse.Nfse.InfNfse.Servico.CodigoMunicipio
                            End If

                            sErro = "252"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.Servico.CodigoTributacaoMunicipio = Nothing Then
                                sCodigoTributacaoMunicipio = objCompNfse.Nfse.InfNfse.Servico.CodigoTributacaoMunicipio
                            End If
                        End If

                        sErro = "253"
                        sMsg1 = "conulta nfse por RPS"


                        If Not objCompNfse.Nfse.InfNfse.TomadorServico Is Nothing Then

                            sErro = "254"
                            sMsg1 = "conulta nfse por RPS"

                            If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco Is Nothing Then

                                sErro = "255"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Bairro = Nothing Then
                                    sBairroTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Bairro
                                End If

                                sErro = "256"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Cep = Nothing Then
                                    lCepTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Cep
                                End If

                                sErro = "257"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.CodigoMunicipio = Nothing Then
                                    lCodigoMunicipioTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.CodigoMunicipio
                                End If

                                sErro = "258"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento = Nothing Then
                                    sComplementoTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento
                                End If

                                sErro = "259"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Endereco = Nothing Then
                                    sEnderecoTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Endereco
                                End If

                                sErro = "260"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Numero = Nothing Then
                                    sNumeroTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Numero
                                End If

                                sErro = "261"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Uf = Nothing Then
                                    sUFTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Uf
                                End If

                                sErro = "262"
                                sMsg1 = "conulta nfse por RPS"

                                If Not objCompNfse.Nfse.InfNfse.TomadorServico.RazaoSocial = Nothing Then
                                    sRazaoSocialTomador = objCompNfse.Nfse.InfNfse.TomadorServico.RazaoSocial
                                End If


                            End If
                        End If


#End If

                        dAliquota = 0
                        dBaseCalculo = 0
                        dDescontoCondicionado = 0
                        dDescontoIncondicionado = 0
                        dISSRetido = 0
                        dOutrasRetencoes = 0
                        dValorCofins = 0
                        dValorCsll = 0
                        dValorDeducoes = 0
                        dValorInss = 0
                        dValorIr = 0
                        dValorIss = 0
                        dValorIssRetido = 0
                        dValorLiquidoNfse = 0
                        dValorPis = 0
                        dValorServicos = 0
                        dValorCredito = 0
                        dtDataEmissaoRPS = Now.Date

#If ABRASF2 Then

                        If Not objCompNfse.Nfse.InfNfse.ValoresNfse.BaseCalculo = Nothing Then
                            dBaseCalculo = Replace(objCompNfse.Nfse.InfNfse.ValoresNfse.BaseCalculo, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.ValoresNfse.ValorLiquidoNfse = Nothing Then
                            dValorLiquidoNfse = Replace(objCompNfse.Nfse.InfNfse.ValoresNfse.ValorLiquidoNfse, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico Is Nothing Then

                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.IssRetido = Nothing Then
                                dISSRetido = objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.IssRetido
                            End If


                            If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores Is Nothing Then

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.Aliquota = Nothing Then
                                    dAliquota = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.Aliquota, ".", ",")
                                End If

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.DescontoCondicionado = Nothing Then
                                    dDescontoCondicionado = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.DescontoCondicionado, ".", ",")
                                End If

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.DescontoIncondicionado = Nothing Then
                                    dDescontoIncondicionado = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.DescontoIncondicionado, ".", ",")
                                End If

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.OutrasRetencoes = Nothing Then
                                    dOutrasRetencoes = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.OutrasRetencoes, ".", ",")
                                End If

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorCofins = Nothing Then
                                    dValorCofins = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorCofins, ".", ",")
                                End If

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorCsll = Nothing Then
                                    dValorCsll = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorCsll, ".", ",")
                                End If

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorDeducoes = Nothing Then
                                    dValorDeducoes = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorDeducoes, ".", ",")
                                End If

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorInss = Nothing Then
                                    dValorInss = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorInss, ".", ",")
                                End If

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorIr = Nothing Then
                                    dValorIr = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorIr, ".", ",")
                                End If

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorIss = Nothing Then
                                    dValorIss = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorIss, ".", ",")
                                End If

                                'If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido = Nothing Then
                                '    dValorIssRetido = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido, ".", ",")
                                'End If

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorPis = Nothing Then
                                    dValorPis = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorPis, ".", ",")
                                End If

                                If Not objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorServicos = Nothing Then
                                    dValorServicos = Replace(objCompNfse.Nfse.InfNfse.DeclaracaoPrestacaoServico.InfDeclaracaoPrestacaoServico.Servico.Valores.ValorServicos, ".", ",")
                                End If

                            End If
                        End If

#Else

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.Aliquota = Nothing Then
                            dAliquota = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.Aliquota, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.BaseCalculo = Nothing Then
                            dBaseCalculo = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.BaseCalculo, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoCondicionado = Nothing Then
                            dDescontoCondicionado = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoCondicionado, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoIncondicionado = Nothing Then
                            dDescontoIncondicionado = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoIncondicionado, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.IssRetido = Nothing Then
                            dISSRetido = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.IssRetido, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.OutrasRetencoes = Nothing Then
                            dOutrasRetencoes = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.OutrasRetencoes, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCofins = Nothing Then
                            dValorCofins = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCofins, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCsll = Nothing Then
                            dValorCsll = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCsll, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorDeducoes = Nothing Then
                            dValorDeducoes = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorDeducoes, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorInss = Nothing Then
                            dValorInss = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorInss, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIr = Nothing Then
                            dValorIr = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIr, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIss = Nothing Then
                            dValorIss = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIss, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido = Nothing Then
                            dValorIssRetido = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorLiquidoNfse = Nothing Then
                            dValorLiquidoNfse = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorLiquidoNfse, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorPis = Nothing Then
                            dValorPis = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorPis, ".", ",")
                        End If

                        If Not objCompNfse.Nfse.InfNfse.Servico.Valores.ValorServicos = Nothing Then
                            dValorServicos = Replace(objCompNfse.Nfse.InfNfse.Servico.Valores.ValorServicos, ".", ",")
                        End If


#End If

                        If Not objCompNfse.Nfse.InfNfse.ValorCredito = Nothing Then
                            dValorCredito = Replace(objCompNfse.Nfse.InfNfse.ValorCredito, ".", ",")
                        End If

                        'If objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco Is Nothing Then

                        '    sErro = "236"
                        '    sMsg1 = "conulta nfse por RPS"

                        '    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBProt ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, data, hora, Versao, CodigoVerificacao, Competencia, DataEmissao, DataEmissaoRPS, Id,  TipoRPS, SerieRPS, NumeroRPS, NaturezaOperacao, Numero, OptanteSimplesNacional, OrgaoGeradorCodMun, OrgaoGeradorUF, " & _
                        '    "OutrasInformacoes, ContatoEmail, ContatoTelefone, PrestadorBairro, PrestadorCEP, PrestadorCodMun, PrestadorComplemento, PrestadorEndereco, PrestadorNumero, PrestadorUF, PrestadorCNPJ, PrestadorInscMun, PrestadorNomeFantasia, PrestadorRazaoSocial, " & _
                        '    "RegimeEspecialTributaco, ServicoCodigoCNAE,ServicoCodMun, ServicoCodTrib, ServicoDiscriminacao, ServicoItemLista, Aliquota, BaseCalculo, DescontoCondicionado, DescontoIncondicionado, ISSRetido, OutrasRetencoes, " & _
                        '    "ValorCofins, ValorCsll, ValorDeducoes, ValorINSS, ValorIR, ValorISS, ValorISSRetido, ValorLiquidoNfse, ValorPIS, ValorServicos,  TomadorEmail, TomadorTelefone, TomadorBairro, TomadorCep, TomadorCodMun, " & _
                        '    "TomadorComplemento, TomandorEndereco, TomadorNumero, TomadorUF, TomadorCPFCNPJ, TomadorInscMun, TomadorRazaoSocial, ValorCredito, NumIntNF)  VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, " & _
                        '    "{20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, {59}, {60}, {61}, {62}, {63}, {64}, {65}, {66}, {67}, {68}, {69}, {70})", _
                        '    lRPSWEBProtNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, Now.Date, TimeOfDay.ToOADate, "", sCodigoVerificacao, dtCompetencia, dtDataEmissao, If(objCompNfse.Nfse.InfNfse.DataEmissaoRps = Nothing, Now.Date, objCompNfse.Nfse.InfNfse.DataEmissaoRps), IIf(objCompNfse.Nfse.InfNfse.Id = Nothing, "", objCompNfse.Nfse.InfNfse.Id), sbTipo, sSerie, sNumeroRPS, _
                        '    sbNaturezaOperacao, sNumero, sbOptanteSimplesNacional, objCompNfse.Nfse.InfNfse.OrgaoGerador.CodigoMunicipio, sUFGerador, IIf(objCompNfse.Nfse.InfNfse.OutrasInformacoes = Nothing, "", objCompNfse.Nfse.InfNfse.OutrasInformacoes), _
                        '    sEmailPrestador, sTelefonePrestador, "", "", "", "", "", _
                        '    "", "", sCNPJPrestador, sInscricaoMunicipalPrestador, IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia = Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia), sRazaoSocialPrestador, _
                        '    IIf(objCompNfse.Nfse.InfNfse.RegimeEspecialTributacao = Nothing, "", objCompNfse.Nfse.InfNfse.RegimeEspecialTributacao), IIf(objCompNfse.Nfse.InfNfse.Servico.CodigoCnae = Nothing, "", objCompNfse.Nfse.InfNfse.Servico.CodigoCnae), objCompNfse.Nfse.InfNfse.Servico.CodigoMunicipio, IIf(objCompNfse.Nfse.InfNfse.Servico.CodigoTributacaoMunicipio = Nothing, "", objCompNfse.Nfse.InfNfse.Servico.CodigoTributacaoMunicipio), sDiscriminacao, _
                        '    sItemListaServico, objCompNfse.Nfse.InfNfse.Servico.Valores.Aliquota, objCompNfse.Nfse.InfNfse.Servico.Valores.BaseCalculo, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoCondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoIncondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.IssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.OutrasRetencoes, _
                        '    objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCofins, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCsll, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorDeducoes, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorInss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIr, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorLiquidoNfse, _
                        '    objCompNfse.Nfse.InfNfse.Servico.Valores.ValorPis, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorServicos, sEmailTomador, sTelefoneTomador, sBairroTomador, sCEPTomador, sCodigoMunicipioTomador, _
                        '    sComplementoTomador, sEnderecoTomador, sNumeroTomador, sUFTomador, sCPFCNPJTomador, sInscricaoMunicipalTomador, _
                        '    sRazaoSocialTomador, objCompNfse.Nfse.InfNfse.ValorCredito, lNumIntNFiscal)


                        'Else

                        sErro = "263"
                        sMsg1 = "conulta nfse por RPS"

                        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBProt ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, data, hora, Versao, CodigoVerificacao, Competencia, DataEmissao, DataEmissaoRPS, Id,  TipoRPS, SerieRPS, NumeroRPS, NaturezaOperacao, Numero, OptanteSimplesNacional, OrgaoGeradorCodMun, OrgaoGeradorUF, " & _
                        "OutrasInformacoes, ContatoEmail, ContatoTelefone, PrestadorBairro, PrestadorCEP, PrestadorCodMun, PrestadorComplemento, PrestadorEndereco, PrestadorNumero, PrestadorUF, PrestadorCNPJ, PrestadorInscMun, PrestadorNomeFantasia, PrestadorRazaoSocial, " & _
                        "RegimeEspecialTributaco, ServicoCodigoCNAE,ServicoCodMun, ServicoCodTrib, ServicoDiscriminacao, ServicoItemLista, Aliquota, BaseCalculo, DescontoCondicionado, DescontoIncondicionado, ISSRetido, OutrasRetencoes, " & _
                        "ValorCofins, ValorCsll, ValorDeducoes, ValorINSS, ValorIR, ValorISS, ValorISSRetido, ValorLiquidoNfse, ValorPIS, ValorServicos,  TomadorEmail, TomadorTelefone, TomadorBairro, TomadorCep, TomadorCodMun, " & _
                        "TomadorComplemento, TomandorEndereco, TomadorNumero, TomadorUF, TomadorCPFCNPJ, TomadorInscMun, TomadorRazaoSocial, ValorCredito, NumIntNF)  VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, " & _
                        "{20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, {59}, {60}, {61}, {62}, {63}, {64}, {65}, {66}, {67}, {68}, {69}, {70})", _
                        lRPSWEBProtNumIntDoc, iFilialEmpresa, iRPSAmbiente, Left(sLote, 15), Left(sProtocolo, 50), Now.Date, TimeOfDay.ToOADate, "", Left(sCodigoVerificacao, 9), Left(sCompetencia, 20), dtDataEmissao, dtDataEmissaoRPS, IIf(objCompNfse.Nfse.InfNfse.Id = Nothing, "", Left(objCompNfse.Nfse.InfNfse.Id, 255)), sbTipo, Left(sSerie, 5), Left(sNumeroRPS, 15), _
                        sbNaturezaOperacao, Left(sNumero, 15), sbOptanteSimplesNacional, lCodigoMunicipioGerador, Left(sUFGerador, 2), IIf(objCompNfse.Nfse.InfNfse.OutrasInformacoes = Nothing, "", Left(objCompNfse.Nfse.InfNfse.OutrasInformacoes, 255)), _
                        Left(sEmailPrestador, 80), Left(sTelefonePrestador, 11), Left(sBairroPrestador, 60), lCepPrestador, lCodMunPrestador, Left(sComplementoPrestador, 60), Left(sEnderecoPrestador, 125), _
                        Left(sNumeroPrestador, 10), Left(sUFPrestador, 2), Left(sCNPJPrestador, 14), Left(sInscricaoMunicipalPrestador, 15), Left(sNomeFantasiaPrestador, 60), Left(sRazaoSocialPrestador, 115), _
                        sbRegimeEspecialTributacao, lCodigoCnae, lCodigoMunicipioServico, Left(sCodigoTributacaoMunicipio, 20), Left(sDiscriminacao, 2000), _
                        Left(sItemListaServico, 5), Round(dAliquota, 2), Round(dBaseCalculo, 2), Round(dDescontoCondicionado, 2), Round(dDescontoIncondicionado, 2), Round(dISSRetido, 2), Round(dOutrasRetencoes, 2), _
                        Round(dValorCofins, 2), Round(dValorCsll, 2), Round(dValorDeducoes, 2), Round(dValorInss, 2), Round(dValorIr, 2), Round(dValorIss, 2), Round(dValorIssRetido, 2), Round(dValorLiquidoNfse, 2), _
                        Round(dValorPis, 2), Round(dValorServicos, 2), Left(sEmailTomador, 80), Left(sTelefoneTomador, 11), Left(sBairroTomador, 60), lCepTomador, lCodigoMunicipioTomador, _
                        Left(sComplementoTomador, 60), Left(sEnderecoTomador, 125), Left(sNumeroTomador, 10), Left(sUFTomador, 2), Left(sCPFCNPJTomador, 14), Left(sInscricaoMunicipalTomador, 15), _
                        Left(sRazaoSocialTomador, 115), Round(dValorCredito, 2), lNumIntNFiscal)
                        'sItemListaServico, objCompNfse.Nfse.InfNfse.Servico.Valores.Aliquota, objCompNfse.Nfse.InfNfse.Servico.Valores.BaseCalculo, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoCondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoIncondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.IssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.OutrasRetencoes, _
                        ' objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCofins, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCsll, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorDeducoes, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorInss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIr, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorLiquidoNfse, _
                        ' objCompNfse.Nfse.InfNfse.Servico.Valores.ValorPis, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorServicos, sEmailTomador, sTelefoneTomador, sBairroTomador, lCepTomador, lCodigoMunicipioTomador, _

                        'End If


                        lRPSWEBProtNumIntDoc = lRPSWEBProtNumIntDoc + 1


                    End If


            End Select

            ConsultaRPS = 0

        Catch ex As Exception

            ConsultaRPS = 1

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Replace(Left(ex.Message, 255), "'", "*"), lNumIntNFiscal)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "ERRO - Encerrado a consulta deste lote", 0)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            Form1.Msg.Items.Add("ERRO - " & sErro & " - " & sMsg1)
            Form1.Msg.Items.Add("ERRO - a consulta do lote " & CStr(lLote) & " foi encerrado por erro. Erro = " & ex.Message)

            If Form1.Msg.Items.Count - 15 < 1 Then
                Form1.Msg.TopIndex = 1
            Else
                Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
            End If

            Application.DoEvents()


        End Try

    End Function

    Function ConsultaRPS_Aux(ByVal objConsNfseRpsEnvio As ConsultarNfseRpsEnvio, ByVal sCertificado As String, ByVal iRPSAmbiente As Integer, ByVal objEndFilial As Endereco, ByVal XMLStringCabec As String, ByVal db2 As SGEDadosDataContext, ByRef lRPSWEBConsLoteNumIntDoc As Long, ByRef lRPSWEBLoteLogNumIntDoc As Long, ByVal iFilialEmpresa As Integer, ByVal db1 As SGEDadosDataContext, ByVal sLote As String, ByVal sProtocolo As String, ByRef lRPSWEBProtNumIntDoc As Long, ByVal lLote As Long, ByVal iDebug As Integer, ByVal objFilialEmpresa As FiliaisEmpresa) As Long

        Dim resRPSWEBLote As IEnumerable(Of RPSWEBLote)
        Dim objNFiscal As NFiscal
        Dim resNFiscal As IEnumerable(Of NFiscal)
        Dim lErro As Long
        Dim sErro As String
        Dim sMSg1 As String
        Dim iResult As Integer
        Dim lNumIntNFiscal As Integer

        Try

            objConsNfseRpsEnvio.Prestador = New tcIdentificacaoPrestador

#If ABRASF2 Then
            If Len(objFilialEmpresa.CGC) = 14 Then
                objConsNfseRpsEnvio.Prestador.CpfCnpj = New tcCpfCnpj
                objConsNfseRpsEnvio.Prestador.CpfCnpj.ItemElementName = ItemChoiceType.Cnpj
                objConsNfseRpsEnvio.Prestador.CpfCnpj.Item = objFilialEmpresa.CGC
            ElseIf Len(objFilialEmpresa.CGC) = 11 Then
                objConsNfseRpsEnvio.Prestador.CpfCnpj = New tcCpfCnpj
                objConsNfseRpsEnvio.Prestador.CpfCnpj.ItemElementName = ItemChoiceType.Cpf
                objConsNfseRpsEnvio.Prestador.CpfCnpj.Item = objFilialEmpresa.CGC
            End If

#Else
        objConsNfseRpsEnvio.Prestador.Cnpj = objFilialEmpresa.CGC

#End If

            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then
                Call ADM.Formata_String_AlfaNumerico(objFilialEmpresa.InscricaoMunicipal, objConsNfseRpsEnvio.Prestador.InscricaoMunicipal)
            Else
                Call ADM.Formata_String_Numero(objFilialEmpresa.InscricaoMunicipal, objConsNfseRpsEnvio.Prestador.InscricaoMunicipal)
            End If

            '            Call Formata_String_Numero(objFilialEmpresa.InscricaoMunicipal, objConsNfseRpsEnvio.Prestador.InscricaoMunicipal)

            objConsNfseRpsEnvio.IdentificacaoRps = New tcIdentificacaoRps

            '    iInicio = InStr(objMsgRetorno.Correcao, "RPS Número:")
            'iFim = InStr(objMsgRetorno.Correcao, " Série:")

            'iInicio = 1

            'If iInicio = 0 Then
            '    GoTo Label_118
            'End If

            ' objConsNfseRpsEnvio.IdentificacaoRps.Numero = Mid(objMsgRetorno.Correcao, iInicio + 12, iFim - (iInicio + 12))
            ' objConsNfseRpsEnvio.IdentificacaoRps.Serie = Mid(objMsgRetorno.Correcao, iFim + 8, Len(objMsgRetorno.Correcao) - (iFim + 8))
            'RPS Número: 881 Série: 1)
            'objConsNfseRpsEnvio.IdentificacaoRps.Numero = "115"
            'objConsNfseRpsEnvio.IdentificacaoRps.Serie = "A"

            resRPSWEBLote = db2.ExecuteQuery(Of RPSWEBLote) _
                ("SELECT * FROM RPSWEBLote WHERE Lote = {0} ORDER BY NumIntNF", lLote)

            For Each objRPSWEBLote In resRPSWEBLote

                resNFiscal = db1.ExecuteQuery(Of NFiscal) _
                    ("SELECT * FROM NFiscal WHERE NumINtDoc = {0}", objRPSWEBLote.NumIntNF)

                lNumIntNFiscal = objRPSWEBLote.NumIntNF

                objNFiscal = resNFiscal(0)

                objConsNfseRpsEnvio.IdentificacaoRps.Numero = objNFiscal.NumNotaFiscal
                objConsNfseRpsEnvio.IdentificacaoRps.Serie = objNFiscal.Serie
                objConsNfseRpsEnvio.IdentificacaoRps.Tipo = "1"

                sErro = "117"
                sMSg1 = "vai iniciar ConsultaRPS"

                If iDebug = 1 Then MsgBox(sErro & " " & sMSg1)

                lErro = ConsultaRPS(objConsNfseRpsEnvio, sCertificado, iRPSAmbiente, objEndFilial, XMLStringCabec, db2, lRPSWEBConsLoteNumIntDoc, lRPSWEBLoteLogNumIntDoc, iFilialEmpresa, db1, sLote, sProtocolo, lRPSWEBProtNumIntDoc, lLote)


                If lErro = 1 Then
                    Throw New System.Exception("Erro na execução da rotina ConsultaRPS")
                End If

            Next

            ConsultaRPS_Aux = 0

        Catch ex As Exception

            ConsultaRPS_Aux = 1

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Replace(Left(ex.Message, 255), "'", "*"), lNumIntNFiscal)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "ERRO - Encerrado a consulta deste lote", 0)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            Form1.Msg.Items.Add("ERRO - " & sErro & " - " & sMSg1)
            Form1.Msg.Items.Add("ERRO - a consulta do lote " & CStr(lLote) & " foi encerrado por erro. Erro = " & ex.Message)

            If Form1.Msg.Items.Count - 15 < 1 Then
                Form1.Msg.TopIndex = 1
            Else
                Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
            End If

            Application.DoEvents()


        End Try

    End Function

#End If
End Class

