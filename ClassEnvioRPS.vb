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


Public Class ClassEnvioRPS

    Public Const RPS_AMBIENTE_HOMOLOGACAO As Integer = 2
    Public Const RPS_AMBIENTE_PRODUCAO As Integer = 1

    Public Const TRIB_TIPO_CALCULO_VALOR = 0
    Public Const TRIB_TIPO_CALCULO_PERCENTUAL = 1

    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Long
    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)


    Const TIPODOC_TRIB_NF = 0

    Public Function Envia_Lote_RPS(ByVal sEmpresa As String, ByVal lLote As Long, ByVal iFilialEmpresa As Integer) As Long

        '====================================================================
        ' ROTAS NOVAS (mínima interferência no ABRASF)
        ' - São Paulo (Nota Fiscal Paulistana): SOAP/XML conforme exemplo + XSD v02.2
        ' - Nacional: API (JSON) - wrapper (mantido separado)
        '====================================================================
        Try
            Dim rota As String = NfseRouter.DetectarRota(sEmpresa, iFilialEmpresa)
            If rota = NfseRouter.ROTA_SAO_PAULO Then
                Dim sp As New NfsePaulistanaEmitter()
                Return sp.Envia_Lote_RPS_Paulistana(sEmpresa, lLote, iFilialEmpresa)
            ElseIf rota = NfseRouter.ROTA_NACIONAL Then
                Dim nac As New NfseNacionalEmitter()
                Return nac.Envia_Lote_RPS_Nacional(sEmpresa, lLote, iFilialEmpresa)
            End If
        Catch ex As Exception
            'Se falhar a detecção/rota nova, cai no fluxo ABRASF existente.
        End Try

#If Not TATUI2 Then

#If Not RJ Then
        Dim a4 As cabecalho = New cabecalho
#End If

        Dim sArquivo As String

        Dim iIndice As Integer
        Dim iAchou As Integer

        Dim XMLStream As MemoryStream = New MemoryStream(10000)
        Dim XMLStream1 As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamCabec As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamDados As MemoryStream = New MemoryStream(10000)

        Dim db1 As SGEDadosDataContext = New SGEDadosDataContext
        Dim db2 As SGEDadosDataContext = New SGEDadosDataContext

        Dim dic As SGEDicDataContext = New SGEDicDataContext

        Dim results As IEnumerable(Of NFeNFiscal)
        Dim resCodTribMun As IEnumerable(Of CodTribMun)
        Dim resFiliaisClientes As IEnumerable(Of FiliaisCliente)
        Dim resEndereco As IEnumerable(Of Endereco)
        Dim resEstado As IEnumerable(Of Estado)
        Dim resFilialEmpresa As IEnumerable(Of FiliaisEmpresa)
        Dim resCidade As IEnumerable(Of Cidade)
        Dim resCliente As IEnumerable(Of Cliente)
        Dim resEndDest As IEnumerable(Of Endereco)
        Dim resItemNF As IEnumerable(Of ItensNFiscal)
        Dim resTribDocItem As IEnumerable(Of TributacaoDocItem)
        Dim resTributacaoDoc As IEnumerable(Of TributacaoDoc)
        Dim resFatConfig As IEnumerable(Of FATConfig)
        Dim resRPSWEBLote As IEnumerable(Of RPSWEBLote)
        Dim resISSQN As IEnumerable(Of ISSQN)
        Dim resProduto As IEnumerable(Of Produto)
        Dim resAdmConfig As IEnumerable(Of AdmConfig)
        Dim resMensagensRegra As IEnumerable(Of MensagensRegra)
        Dim resdicControle As IEnumerable(Of Controle)
        Dim resProdutoCNAE As IEnumerable(Of ProdutoCNAE)

        Dim objProdutoCNAE As ProdutoCNAE
        Dim objControle As Controle
        Dim objMensagensRegra As MensagensRegra
        Dim objAdmConfig As AdmConfig
        Dim objISSQN As ISSQN
        Dim objCodTribMun As CodTribMun
        Dim objRPSWEBLote As RPSWEBLote
        Dim objFatConfig As FATConfig
        Dim objTributacaoDoc As TributacaoDoc
        Dim objTribDocItem As TributacaoDocItem
        Dim objItemNF As ItensNFiscal
        Dim objCliente As Cliente
        Dim objCidade As Cidade
        Dim objFilialEmpresa As FiliaisEmpresa
        Dim objEstado As Estado
        Dim objEndereco As Endereco
        Dim objFiliaisClientes As FiliaisCliente
        Dim lEndereco As Long
        Dim lEndDest As Long
        Dim objProduto As Produto

        Dim colNFiscal As Collection = New Collection


        Dim sIM As String

        Dim XMLString As String
        Dim XMLString1 As String
        Dim XMLStringRpses As String
        Dim XMLStringCabec As String

        Dim iResult As Integer

        Dim cert As X509Certificate2 = New X509Certificate2
        Dim certificado As Certificado = New Certificado

        Dim homologacao As New NotaCariocaHomologacaoAssinc.NfseSoapClient
        Dim producao As New NotaCariocaProducaoAssinc.NfseSoapClient

        Dim producaoosasco As New br.com.nfeosasco.www.NotaFiscalEletronicaServico
        Dim emitir As New nfseabrasf.br.com.nfeosasco.www.EmissaoNotaFiscalRequest

        'emitir.NotaFiscal.Atividade = "1.01"

        'producaoosasco.Emitir(emitir)


        Dim homologacaobh = New GetWebRequest_bhHomologacao
        Dim producaobh = New GetWebRequest_bhProducao


        Dim homosalvadorEnvio As New br.gov.ba.salvador.sefaz.nfsehml.EnvioLoteRPS
        Dim homosalvadorSitLote As New br.gov.ba.salvador.sefaz.nfsehml1.ConsultaSituacaoLoteRPS
        Dim prodsalvadorEnvio As New br.gov.ba.salvador.sefaz.nfse.EnvioLoteRPS
        Dim prodsalvadorSitLote As New br.gov.ba.salvador.sefaz.nfse1.ConsultaSituacaoLoteRPS

        Dim homologacaoginfes = New br.com.ginfes.homologacao.ServiceGinfesImplService
        Dim producaoginfes = New br.com.ginfes.producao.ServiceGinfesImplService

        Dim homsistema4rEnvio = New br.com.sistemas4r.abrasf.nfsehml.RecepcionarLoteRpsSincrono


        Dim prodsistema4rEnvio As New br.com.sistemas4r.abrasf.nfse.RecepcionarLoteRpsSincrono


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
        Dim lNumNotaFiscal As Long

        Dim sCertificado As String
        Dim iRPSAmbiente As Integer

        Dim dValorServPIS As Double
        Dim dValorServCOFINS As Double

        Dim dValorPIS As Double
        Dim dValorCOFINS As Double

        Dim dServNTribICMS As Double
        Dim iItem As Integer
        Dim xRet As Byte()

        Dim resdicEmpresa As IEnumerable(Of Empresa)


        Dim objEmpresa As Empresa

        Dim dtDataRecebimento As Date


        Dim objListaMsgRetorno As ListaMensagemRetorno
        Dim objMsgRetorno As New tcMensagemRetorno

#If ABRASF Or ABRASF2 Then
        Dim objListaMsgRetornoLote As ListaMensagemRetornoLote
        Dim objMsgRetornoLote As New tcMensagemRetornoLote
        Dim objConsultarLoteRpsRespostaListaNFse As ListaMensagemRetornoLote
#End If

#If ABRASF2 Then
        Dim objEnviarLoteRpsEnvio As New EnviarLoteRpsSincronoEnvio

#End If


        Dim objCompNfse As tcCompNfse
        Dim sProtocolo As String
        Dim sLote As String
        Dim dHoraRecebimento As Double

        Dim lRPSWEBConsSitLoteNumIntDoc As Long
        Dim lRPSWEBConsLoteNumIntDoc As Long
        Dim lRPSWEBProtNumIntDoc As Long
        Dim lRPSWEBLoteLogNumIntDoc As Long
        Dim lRPSWEBRetEnviNumIntDoc As Long

        Dim iDebug As Integer

        Dim sEmailPrestador As String
        Dim sTelefonePrestador As String
        Dim sEmailTomador As String
        Dim sTelefoneTomador As String
        Dim sAux As String
        Dim objNFiscal As Object
        Dim sProduto As String
        Dim objEndFilial As Endereco
        Dim sMsg As String

        Dim sErro As String
        Dim sMsg1 As String
        Dim iAchouCNAE As Integer

        sSerie = ""
        sErro = ""
        sMsg1 = ""

        Try

#If Not TATUI2 Then

            iTamanho = 255
            sRetorno = StrDup(iTamanho, Chr(0))

            iDebug = 0

            Call GetPrivateProfileString("Geral", "Debug", 0, sRetorno, iTamanho, "Adm100.ini")

            If IsNumeric(sRetorno) Then
                iDebug = CInt(sRetorno)
            End If

            lNumIntNF = 0

            '********** pega o diretorio do log para colocar os arquivos xml *************
            'iTamanho = 255
            'sRetorno = StrDup(iTamanho, Chr(0))

            'Call GetPrivateProfileString("Geral", "Arqlog", -1, sRetorno, iTamanho, "Adm100.ini")

            'iPos = InStr(sRetorno, Chr(0))

            'sRetorno = Mid(sRetorno, 1, iPos - 1)

            'sFile = Dir(sRetorno)

            'iPos = InStr(sRetorno, sFile)

            'sDir = Mid(sRetorno, 1, iPos - 1)

            '********** pega o diretorio dos executaveis para ler os arquivos xsd *************
            iTamanho = 255
            sRetorno = StrDup(iTamanho, Chr(0))

            Call GetPrivateProfileString("Forprint", "DirBin", -1, sRetorno, iTamanho, "ADM100.INI")

            iPos = InStr(sRetorno, Chr(0))

            sDir1 = Mid(sRetorno, 1, iPos - 1)

            sErro = "1"
            sMsg1 = "vai abrir o bd odbc"

            If iDebug = 1 Then MsgBox("1")




            '***** coloca a string de conexao apontando para o SGEDados em questao *****
            odbc.ConnectionString = "DSN=SGEDados" & sEmpresa & ";UID=sa;PWD=SAPWD"
            '            odbc.ConnectionString = "DSN=SGEDados" & sEmpresa & ";UID=sa;PWD=SAPWD"

            If iDebug = 1 Then MsgBox(odbc.ConnectionString)

            odbc.Open()

            sErro = "1.1"
            sMsg1 = "abriu o bd odbc"

            If iDebug = 1 Then MsgBox("abriu BD")

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

            sErro = "2"
            sMsg1 = "vai abrir o dic odbc"

            If iDebug = 1 Then MsgBox("2")

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

            sErro = "3"
            sMsg1 = "vai abrir o bd"

            If iDebug = 1 Then MsgBox("3")


            db1.Connection.Open()

            sErro = "4"
            sMsg1 = "vai abrir o bd"

            If iDebug = 1 Then MsgBox("4")


            db2.Connection.Open()

            sErro = "5"
            sMsg1 = "vai abrir o dic"

            If iDebug = 1 Then MsgBox("5")

            dic.Connection.Open()

            sErro = "6"
            sMsg1 = "vai abrir a transacao"

            If iDebug = 1 Then MsgBox("6")

            db2.Transaction = db2.Connection.BeginTransaction()
            db1.Transaction = db1.Connection.BeginTransaction()

            sErro = "7"
            sMsg1 = "vai consultar a tabela Controle"

            If iDebug = 1 Then MsgBox("7")

            sDir = ""

            resdicControle = dic.ExecuteQuery(Of Controle) _
            ("SELECT * FROM Controle WHERE Codigo = {0}", 101)

            For Each objControle In resdicControle
                sDir = objControle.Conteudo
                Exit For
            Next

            sErro = "7.1"

            If sDir = "" Then

                sErro = "7.2"

                resdicControle = dic.ExecuteQuery(Of Controle) _
                ("SELECT * FROM Controle WHERE Codigo = {0}", 1)

                For Each objControle In resdicControle
                    sRetorno = objControle.Conteudo

                    sFile = Dir(sRetorno)

                    iPos = InStr(sRetorno, sFile)

                    sDir = Mid(sRetorno, 1, iPos - 1)

                    Exit For
                Next

            End If

            sErro = "7.3"

            resAdmConfig = db1.ExecuteQuery(Of AdmConfig) _
            ("SELECT * FROM AdmConfig WHERE  Codigo = {0} ", "VERSAO_MSG")

            If resAdmConfig.Count > 0 Then

                sErro = "7.4"

                resAdmConfig = db1.ExecuteQuery(Of AdmConfig) _
                ("SELECT * FROM AdmConfig WHERE  Codigo = {0} ", "VERSAO_MSG")

                objAdmConfig = resAdmConfig(0)
            Else
                objAdmConfig = New AdmConfig
            End If

            sErro = "7.5"

            resFatConfig = db2.ExecuteQuery(Of FATConfig) _
            ("SELECT * FROM FatConfig WHERE Codigo = {0} ", "NUM_INT_PROX_RPSWEBLOTELOG")

            objFatConfig = resFatConfig(0)

            sErro = "8"
            sMsg1 = "vai consultar FatConfig - NUM_INT_PROX_RPSWEBCONSSITLOTE"


            If iDebug = 1 Then MsgBox("8")

            lRPSWEBLoteLogNumIntDoc = CLng(objFatConfig.Conteudo)

            resFatConfig = db2.ExecuteQuery(Of FATConfig) _
            ("SELECT * FROM FatConfig WHERE Codigo = {0} ", "NUM_INT_PROX_RPSWEBCONSSITLOTE")

            objFatConfig = resFatConfig(0)

            sErro = "9"
            sMsg1 = "vai consultar FatConfig - NUM_INT_PROX_RPSWEBCONSLOTE"

            If iDebug = 1 Then MsgBox("9")

            lRPSWEBConsSitLoteNumIntDoc = CLng(objFatConfig.Conteudo)

            resFatConfig = db2.ExecuteQuery(Of FATConfig) _
            ("SELECT * FROM FatConfig WHERE Codigo = {0} ", "NUM_INT_PROX_RPSWEBCONSLOTE")

            objFatConfig = resFatConfig(0)

            sErro = "10"
            sMsg1 = "vai consultar FatConfig - NUM_INT_PROX_RPSWEBPROT"

            If iDebug = 1 Then MsgBox("10")

            lRPSWEBConsLoteNumIntDoc = CLng(objFatConfig.Conteudo)

            resFatConfig = db2.ExecuteQuery(Of FATConfig) _
            ("SELECT * FROM FatConfig WHERE Codigo = {0} ", "NUM_INT_PROX_RPSWEBPROT")

            objFatConfig = resFatConfig(0)

            sErro = "11"
            sMsg1 = "vai consultar FatConfig - NUM_INT_PROX_RPSWEBRETENVI"

            If iDebug = 1 Then MsgBox("11")

            lRPSWEBProtNumIntDoc = CLng(objFatConfig.Conteudo)


            resFatConfig = db2.ExecuteQuery(Of FATConfig) _
            ("SELECT * FROM FatConfig WHERE Codigo = {0} ", "NUM_INT_PROX_RPSWEBRETENVI")

            objFatConfig = resFatConfig(0)

            sErro = "12"
            sMsg1 = "vai consultar FiliaisEmpresa"

            If iDebug = 1 Then MsgBox("12")

            lRPSWEBRetEnviNumIntDoc = CLng(objFatConfig.Conteudo)



            '
            '  seleciona certificado do repositório MY do windows
            '

            resFilialEmpresa = db1.ExecuteQuery(Of FiliaisEmpresa) _
            ("SELECT * FROM FiliaisEmpresa WHERE FilialEmpresa = {0} ", iFilialEmpresa)

            For Each objFilialEmpresa In resFilialEmpresa

                sCertificado = objFilialEmpresa.CertificadoA1A3
                iRPSAmbiente = objFilialEmpresa.RPSAmbiente
                Exit For
            Next

            If iDebug = 1 Then MsgBox("13")


            sErro = "13"
            sMsg1 = "vai inserir RPSWEBLoteLog"


            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "Iniciando processamento...", 0)

            sErro = "13.01"
            sMsg1 = "vai iniciar o processamento"

            If iDebug = 1 Then MsgBox("13.01")

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            Form1.Msg.Items.Add("Iniciando processamento...")

            If Form1.Msg.Items.Count - 15 < 1 Then
                Form1.Msg.TopIndex = 1
            Else
                Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
            End If

            Application.DoEvents()

            sErro = "13.1"
            sMsg1 = "vai pesquisar o certificado"

            If iDebug = 1 Then MsgBox("13.1")

            cert = certificado.BuscaNome(sCertificado)

            sErro = "13.2"
            sMsg1 = "vai consultar a tabela Enderecos"

            If iDebug = 1 Then MsgBox("13.2")

            resEndereco = db1.ExecuteQuery(Of Endereco) _
            ("SELECT * FROM Enderecos WHERE Codigo = {0}", objFilialEmpresa.Endereco)

            objEndFilial = resEndereco(0)

#If Not RJ Then
            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then

                a4.versao = "1.00"
                a4.versaoDados = "1.00"

            End If

            If UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then
                '                If UCase(objEndFilial.Cidade) = "TATUÍ" Or UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then

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
                doccabec.Load(XMLStreamCabec)

                doccabec.Save(sDir & "XmlRPSCabec.xml")

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

                sErro = "15"
                sMsg1 = "vai consultar a tabela Empresas"


                If iDebug = 1 Then MsgBox("15")

                Dim xmcabec As Byte()

                xmcabec = XMLStreamCabec.ToArray

                XMLStringCabec = System.Text.Encoding.UTF8.GetString(xmcabec)

                XMLStringCabec = Mid(XMLStringCabec, 1, 19) & " encoding=""utf-8"" " & Mid(XMLStringCabec, 20)

                '          XMLStringCabec = Replace(XMLStringCabec, " xmlns=""""", "")



            End If
#End If





            resdicEmpresa = dic.ExecuteQuery(Of Empresa) _
            ("SELECT * FROM Empresas WHERE Codigo = {0}", CLng(sEmpresa))

            For Each objEmpresa In resdicEmpresa
                Exit For
            Next

            sErro = "16"
            sMsg1 = "vai consultar a tabela RPSWEBLote"

            If iDebug = 1 Then MsgBox("16")

            resRPSWEBLote = db2.ExecuteQuery(Of RPSWEBLote) _
            ("SELECT * FROM RPSWEBLote WHERE Lote = {0} ORDER BY NumIntNF", lLote)


            Form1.ProgressBar1.Maximum = resRPSWEBLote.Count
            Form1.ProgressBar1.Minimum = 0
            Form1.ProgressBar1.Value = 0

            sErro = "17"
            sMsg1 = "vai consultar a tabela RPSWEBLote"

            If iDebug = 1 Then MsgBox("17")

            resRPSWEBLote = db2.ExecuteQuery(Of RPSWEBLote) _
            ("SELECT * FROM RPSWEBLote WHERE Lote = {0} ORDER BY NumIntNF", lLote)

            iItem = -1

            For Each objRPSWEBLote In resRPSWEBLote

                sErro = "18"
                sMsg1 = "vai consultar a tabela NFeNFiscal"

                If iDebug = 1 Then MsgBox("18")

                iItem = iItem + 1

                Dim objInfRps As tcInfRps = New tcInfRps

#If ABRASF2 Then

                objEnviarLoteRpsEnvio.LoteRps = New tcLoteRps

                objEnviarLoteRpsEnvio.LoteRps.ListaRps = New tcLoteRpsListaRps

                objEnviarLoteRpsEnvio.LoteRps.ListaRps.Rps = New tcDeclaracaoPrestacaoServico

                objEnviarLoteRpsEnvio.LoteRps.ListaRps.Rps.InfDeclaracaoPrestacaoServico = New tcInfDeclaracaoPrestacaoServico

                objEnviarLoteRpsEnvio.LoteRps.ListaRps.Rps.InfDeclaracaoPrestacaoServico.Rps = objInfRps

                '                objEnviarLoteRpsEnvio.LoteRps.ListaRps.Rps.InfDeclaracaoPrestacaoServico.Servico = New tcDadosServico
#End If

                lNumIntNF = objRPSWEBLote.NumIntNF

                'lNumIntNFiscalParam
                results = db1.ExecuteQuery(Of NFeNFiscal) _
                ("SELECT * FROM NFeNFiscal WHERE NumIntDoc = {0} ", objRPSWEBLote.NumIntNF)

                For Each objNFiscal In results

                    sErro = "19"
                    sMsg1 = "vai gravar RPSWEBLoteLog"

                    If iDebug = 1 Then MsgBox("19")

                    sSerie = ADM.Serie_Sem_E(objNFiscal.Serie)

                    lNumNotaFiscal = objNFiscal.NumNotaFiscal

                    objInfRps.Id = "rps_" & sSerie & "_" & lNumNotaFiscal


                    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog (NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
                    lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "Iniciando o processamento da Nota Fiscal = " & objNFiscal.NumNotaFiscal & " Série = " & objNFiscal.Serie, lNumIntNF)

                    lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                    Form1.Msg.Items.Add("Iniciando o processamento da Nota Fiscal = " & objNFiscal.NumNotaFiscal & " Série = " & ADM.Serie_Sem_E(objNFiscal.Serie))

                    If Form1.Msg.Items.Count - 15 < 1 Then
                        Form1.Msg.TopIndex = 1
                    Else
                        Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                    End If

                    Application.DoEvents()

                    colNFiscal.Add(objNFiscal)

                    objInfRps.IdentificacaoRps = New tcIdentificacaoRps

                    objInfRps.IdentificacaoRps.Numero = lNumNotaFiscal
                    objInfRps.IdentificacaoRps.Serie = sSerie
                    objInfRps.IdentificacaoRps.Tipo = 1
                    objInfRps.DataEmissao = Format(objNFiscal.DataEmissao, "yyyy-MM-dd") & "T00:00:00"


#If Not ABRASF2 Then

                    objInfRps.NaturezaOperacao = 1
                    If objFilialEmpresa.RegimeEspecialTrib <> 0 Then
                        objInfRps.RegimeEspecialTributacaoSpecified = True
                        objInfRps.RegimeEspecialTributacao = objFilialEmpresa.RegimeEspecialTrib
                    End If
                    objInfRps.OptanteSimplesNacional = IIf(objFilialEmpresa.SuperSimples = 0, 2, 1)
                    objInfRps.IncentivadorCultural = 2

#Else
                    If objFilialEmpresa.RegimeEspecialTrib <> 0 Then
                        objEnviarLoteRpsEnvio.LoteRps.ListaRps.Rps.InfDeclaracaoPrestacaoServico.RegimeEspecialTributacaoSpecified = True
                        objEnviarLoteRpsEnvio.LoteRps.ListaRps.Rps.InfDeclaracaoPrestacaoServico.RegimeEspecialTributacao = objFilialEmpresa.RegimeEspecialTrib
                    End If
                    objEnviarLoteRpsEnvio.LoteRps.ListaRps.Rps.InfDeclaracaoPrestacaoServico.OptanteSimplesNacional = IIf(objFilialEmpresa.SuperSimples = 0, 2, 1)
                    objEnviarLoteRpsEnvio.LoteRps.ListaRps.Rps.InfDeclaracaoPrestacaoServico.Competencia = Format(objNFiscal.DataEmissao, "yyyy-MM-dd") & "T00:00:00"
                    objEnviarLoteRpsEnvio.LoteRps.ListaRps.Rps.InfDeclaracaoPrestacaoServico.IncentivoFiscal = 2
#End If
                    objInfRps.Status = 1

                    sErro = "20"
                    sMsg1 = "vai consultar a tabela TributacaoDoc"

                    If iDebug = 1 Then MsgBox("20")

                    Dim objDadosServico As tcDadosServico = New tcDadosServico

#If Not ABRASF2 Then
                        objInfRps.Servico = objDadosServico
                        Dim objValores As tcValores = New tcValores

#Else
                    objEnviarLoteRpsEnvio.LoteRps.ListaRps.Rps.InfDeclaracaoPrestacaoServico.Servico = objDadosServico
                    Dim objValores As tcValoresDeclaracaoServico = New tcValoresDeclaracaoServico
#End If

                    objDadosServico.Valores = objValores


                    resTributacaoDoc = db1.ExecuteQuery(Of TributacaoDoc) _
                    ("SELECT *  FROM TributacaoDoc WHERE TipoDoc = {0} AND NumIntDoc = {1}", TIPODOC_TRIB_NF, objNFiscal.NumIntDoc)

                    objTributacaoDoc = resTributacaoDoc(0)

                    sErro = "21"
                    sMsg1 = "vai consultar a tabela ItensNFiscal"

                    If iDebug = 1 Then MsgBox("21")


                    resItemNF = db1.ExecuteQuery(Of ItensNFiscal) _
                    ("SELECT * FROM ItensNFiscal WHERE  NumIntNF = {0} ORDER BY Item", lNumIntNF)

                    iIndice = -1

                    sErro = "22"
                    sMsg1 = "consultou ItensNFiscal"

                    If iDebug = 1 Then MsgBox("22")

                    dValorServPIS = 0
                    dValorServCOFINS = 0
                    dServNTribICMS = 0
                    dValorPIS = 0
                    dValorCOFINS = 0

                    'importante quando tiver mais de uma rps sendo enviada    
                    iAchou = 0

                    For Each objItemNF In resItemNF


                        sErro = "23"
                        sMsg1 = "vai consultar TributacaoDocItem"

                        If iDebug = 1 Then MsgBox("23")

                        iIndice = iIndice + 1


                        resTribDocItem = db1.ExecuteQuery(Of TributacaoDocItem) _
                        ("SELECT * FROM TributacaoDocItem WHERE TipoDoc = {0} AND NumIntDoc = {1} AND Item = {2}", TIPODOC_TRIB_NF, lNumIntNF, objItemNF.Item)

                        For Each objTribDocItem In resTribDocItem

                            sErro = "24"
                            sMsg1 = "leu TributacaoDocItem"

                            If iDebug = 1 Then MsgBox("24")

                            'resNaturezaOP = db1.ExecuteQuery(Of NaturezaOp) _
                            '("SELECT * FROM NaturezaOP WHERE  Codigo = {0}", objTribDocItem.NaturezaOp)

                            'objNaturezaOP = resNaturezaOP(0)

                            'a5.infNFe.det(iIndice).prod.CFOP = objNaturezaOP.codnfe

                            Exit For
                        Next

                        'se ainda nao encontrou nenhum produto q tenha natureza servico
                        If iAchou = 0 Then

                            resProduto = db1.ExecuteQuery(Of Produto) _
                            ("SELECT * FROM Produtos WHERE  Codigo = {0}", objItemNF.Produto)

                            For Each objProduto In resProduto
                                If objProduto.Natureza = 8 Then
                                    iAchou = 1
                                    Exit For
                                End If
                            Next

                            If iAchou = 1 Then
                                iAchou = 2

                                sErro = "25"
                                sMsg1 = "vai ler ISSQN"

                                If iDebug = 1 Then MsgBox("25")

                                resISSQN = db1.ExecuteQuery(Of ISSQN) _
                                ("SELECT * FROM ISSQN WHERE  Codigo = {0} ", objTribDocItem.ISSQN)

                                objISSQN = resISSQN(0)

                                If objISSQN Is Nothing Then Throw New System.Exception("o ISSQN do produto " & objItemNF.Produto & " provavelmente não foi preenchido quando a nota foi gravada.")

                                sErro = "25.1"
                                sMsg1 = "vai pegar os dados de ItemListaServico e CNAE"

                                If iDebug = 1 Then MsgBox("25.1")

                                ''classificacao do servico conforme tabela da lei complementar 116 de 2003 (LC 116/03)
                                If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then
                                    objDadosServico.ItemListaServico = objISSQN.CListServ
                                    '                                ElseIf UCase(objEndFilial.Cidade) = "SALVADOR" Or UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then
                                ElseIf UCase(objEndFilial.Cidade) = "SALVADOR" Or UCase(objEndFilial.Cidade) = "TATUÍ" Or UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then
                                    objDadosServico.ItemListaServico = Replace(Format(CDbl(objISSQN.Codigo) / 100, "00.00"), ",", ".")
                                Else
                                    objDadosServico.ItemListaServico = objISSQN.Codigo
                                End If
                                ' objDadosServico.ItemListaServico = Replace(objTribDocItem.ISSQN, ".", "")

                                If objDadosServico.ItemListaServico = "" Then

                                    resProduto = db1.ExecuteQuery(Of Produto) _
                                    ("SELECT * FROM Produtos WHERE Codigo = {0}", objItemNF.Produto)
                                    objProduto = resProduto(0)

                                    'If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Or UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then
                                    If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then
                                        objDadosServico.ItemListaServico = Replace(Format(CDbl(objProduto.ISSQN) / 100, "0.00"), ",", ".")
                                    ElseIf UCase(objEndFilial.Cidade) = "TATUÍ" Or UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then
                                        objDadosServico.ItemListaServico = Replace(Format(CDbl(objProduto.ISSQN) / 100, "00.00"), ",", ".")
                                    Else
                                        objDadosServico.ItemListaServico = Replace(objProduto.ISSQN, ".", "")
                                    End If

                                End If

                                iAchouCNAE = 0

                                resProdutoCNAE = db1.ExecuteQuery(Of ProdutoCNAE) _
                                ("SELECT * FROM ProdutoCNAE WHERE Produto = {0}", objItemNF.Produto)

                                For Each objProdutoCNAE In resProdutoCNAE
                                    objDadosServico.CodigoCnae = objProdutoCNAE.CNAE '???? depois pode pegar de produtocnae
                                    iAchouCNAE = 1
                                    Exit For
                                Next

                                If iAchouCNAE = 0 Then
                                    objDadosServico.CodigoCnae = objFilialEmpresa.CNAE '???? depois pode pegar de produtocnae
                                    iAchouCNAE = 1
                                End If

                                objDadosServico.CodigoCnaeSpecified = True
                                sProduto = Trim(objItemNF.Produto)
                            End If
                        End If

                        If Len(objDadosServico.Discriminacao) = 0 Then
                            objDadosServico.Discriminacao = objItemNF.DescricaoItem & " Quant: " & Format(objItemNF.Quantidade, "###.###,###") & " P.Unit: " & Format(objItemNF.PrecoUnitario, "fixed") & " Total: " & Format(objItemNF.Quantidade * objItemNF.PrecoUnitario - objItemNF.ValorDesconto, "fixed")
                        Else
                            objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|" & objItemNF.DescricaoItem & " Quant: " & Format(objItemNF.Quantidade, "###.###,###") & " P.Unit: " & Format(objItemNF.PrecoUnitario, "fixed") & " Total: " & Format(objItemNF.Quantidade * objItemNF.PrecoUnitario - objItemNF.ValorDesconto, "fixed")
                        End If

                        sErro = "26"
                        sMsg1 = "vai pegar os dados de PIS e COFINS"

                        If iDebug = 1 Then MsgBox("26")

                        dValorPIS = dValorPIS + objTribDocItem.PISValor
                        dValorCOFINS = dValorCOFINS + objTribDocItem.COFINSValor

                    Next

                    If iAchou = 0 Then Throw New System.Exception("Não há nenhum produto nesta nota que tenha natureza serviço.")

                    sErro = "27"
                    sMsg1 = "vai transferir os valores"

                    If iDebug = 1 Then MsgBox("27")


                    objDadosServico.Valores.ValorPisSpecified = True
#If TATUI1 Then
                    objDadosServico.Valores.ValorPis = Replace(Format(IIf(objTributacaoDoc.PISRetido <> 0, dValorPIS, 0), "fixed"), ",", ".")
#Else
                    objDadosServico.Valores.ValorPis = Format(IIf(objTributacaoDoc.PISRetido <> 0, objTributacaoDoc.PISRetido, 0), "fixed")
#End If
                    objDadosServico.Valores.ValorCofinsSpecified = True
#If TATUI1 Then
                    objDadosServico.Valores.ValorCofins = Replace(Format(IIf(objTributacaoDoc.COFINSRetido <> 0, dValorCOFINS, 0), "fixed"), ",", ".")
#Else
                    objDadosServico.Valores.ValorCofins = Format(IIf(objTributacaoDoc.COFINSRetido <> 0, objTributacaoDoc.COFINSRetido, 0), "fixed")
#End If
                    objDadosServico.Valores.ValorInssSpecified = True
#If TATUI1 Then
                    objDadosServico.Valores.ValorInss = Replace(Format(IIf(objTributacaoDoc.INSSRetido <> 0, objTributacaoDoc.ValorINSS, 0), "fixed"), ",", ".")
#Else
                    objDadosServico.Valores.ValorInss = Format(IIf(objTributacaoDoc.INSSRetido <> 0, objTributacaoDoc.ValorINSS, 0), "fixed")
#End If

                    objDadosServico.Valores.ValorIrSpecified = True
#If TATUI1 Then
                    objDadosServico.Valores.ValorIr = Replace(Format(objTributacaoDoc.IRRFValor, "fixed"), ",", ".")
#Else
                    objDadosServico.Valores.ValorIr = Format(objTributacaoDoc.IRRFValor, "fixed")
#End If

                    objDadosServico.Valores.ValorCsllSpecified = True
#If TATUI1 Then
                    objDadosServico.Valores.ValorCsll = Replace(Format(objTributacaoDoc.CSLLRetido, "fixed"), ",", ".")
#Else
                    objDadosServico.Valores.ValorCsll = Format(objTributacaoDoc.CSLLRetido, "fixed")
#End If




#If TATUI1 Then
                    objDadosServico.Valores.ValorIss = Replace(Format(objTributacaoDoc.ISSValor, "fixed"), ",", ".")
#Else
                    If UCase(objEndFilial.Cidade) <> "TATUÍ" Then
                        objDadosServico.Valores.ValorIss = Format(objTributacaoDoc.ISSValor, "fixed")
                    End If
#End If

                        objDadosServico.Valores.ValorIssSpecified = True

#If TATUI1 Then
                    If objTributacaoDoc.ISSIncluso = 1 Then
                        objDadosServico.Valores.ValorServicos = Replace(Format(objNFiscal.ValorProdutos, "fixed"), ",", ".")
                    Else
                        objDadosServico.Valores.ValorServicos = Replace(Format(objNFiscal.ValorProdutos + CDbl(Format(objTributacaoDoc.ISSValor, "fixed")), "fixed"), ",", ".")
                    End If
#Else
                        If objTributacaoDoc.ISSIncluso = 1 Then
                            objDadosServico.Valores.ValorServicos = Format(objNFiscal.ValorProdutos, "fixed")
                        Else
                            objDadosServico.Valores.ValorServicos = Format(objNFiscal.ValorProdutos + CDbl(Format(objTributacaoDoc.ISSValor, "fixed")), "fixed")
                        End If
#End If

#If Not ABRASF2 Then

                    objDadosServico.Valores.BaseCalculoSpecified = True

#If TATUI1 Then
                    objDadosServico.Valores.BaseCalculo = Replace(Format(objTributacaoDoc.ISSBase, "fixed"), ",", ".")
#Else
                    objDadosServico.Valores.BaseCalculo = Format(objTributacaoDoc.ISSBase, "fixed")
#End If
#End If

                        objDadosServico.Valores.AliquotaSpecified = True
                        If UCase(objEndFilial.Cidade) = "TATUÍ" Then
                        '                        objDadosServico.Valores.Aliquota = Replace(Format(objTribDocItem.ISSAliquota, "##0.##"), ",", ".")
#If ABRASF2 Then
                            objDadosServico.ExigibilidadeISS = 1
#End If
                        Else
                            objDadosServico.Valores.Aliquota = objTribDocItem.ISSAliquota
                        End If

#If Not ABRASF2 Then
                    objDadosServico.Valores.ValorIssRetidoSpecified = True

                                            'isto foi colocado pois em tatui no ambiente de producao apesar da flag issretido estar com nao
                        'estava exigindo que o valor do iss fosse igual ao valor do issretido.
                        '#If Not TATUI Then
                        objDadosServico.Valores.ValorIssRetido = CDbl(Format(IIf(objTributacaoDoc.ISSRetido > 0, objTributacaoDoc.ISSValor, 0), "fixed"))
                        '#Else
                        '                    objDadosServico.Valores.ValorIssRetido = Replace(Format(objTributacaoDoc.ISSValor, "fixed"), ",", ".")
                        '#End If

                    objDadosServico.Valores.IssRetido = IIf(objTributacaoDoc.ISSRetido > 0, 1, 2)

#Else

                        If objTributacaoDoc.ISSRetido > 0 Then
                            objDadosServico.IssRetido = 1
                        Else
                            objDadosServico.IssRetido = 2
                        End If
#End If


                        objDadosServico.Valores.DescontoIncondicionadoSpecified = True
#If TATUI1 Then
                    objDadosServico.Valores.DescontoIncondicionado = Replace(Format(objNFiscal.ValorDesconto, "fixed"), ",", ".")
#Else
                        objDadosServico.Valores.DescontoIncondicionado = Format(objNFiscal.ValorDesconto, "fixed")
#End If


#If Not ABRASF2 Then
                    If UCase(objEndFilial.Cidade) = "GUARULHOS" And objInfRps.OptanteSimplesNacional = 1 Then
                        objDadosServico.Valores.BaseCalculo = objDadosServico.Valores.ValorServicos - objDadosServico.Valores.DescontoIncondicionado
                    End If
                    objDadosServico.Valores.ValorLiquidoNfseSpecified = True

                    '#If TATUI Then
                    '                    objDadosServico.Valores.ValorLiquidoNfse = Replace(Format(CDec(Replace(objDadosServico.Valores.ValorServicos, ".", ",")) - CDec(Replace(objDadosServico.Valores.ValorPis, ".", ",")) - CDec(Replace(objDadosServico.Valores.ValorCofins, ".", ",")) - CDec(Replace(objDadosServico.Valores.ValorInss, ".", ",")) - CDec(Replace(objDadosServico.Valores.ValorIr, ".", ",")) - CDec(Replace(objDadosServico.Valores.ValorCsll, ".", ",")) - CDec(Replace(objDadosServico.Valores.ValorIssRetido, ".", ",")) - CDec(Replace(objDadosServico.Valores.DescontoIncondicionado, ".", ",")), "fixed"), ",", ".")
                    '#Else
                    objDadosServico.Valores.ValorLiquidoNfse = Format(objDadosServico.Valores.ValorServicos - objDadosServico.Valores.ValorPis - objDadosServico.Valores.ValorCofins - objDadosServico.Valores.ValorInss - objDadosServico.Valores.ValorIr - objDadosServico.Valores.ValorCsll - objDadosServico.Valores.ValorIssRetido - objDadosServico.Valores.DescontoIncondicionado, "fixed")
                    '#End If
#End If



#If TATUI1 Then

                    If CDbl(objDadosServico.Valores.ValorLiquidoNfse) <> CDbl(objDadosServico.Valores.ValorServicos) Then


                        objDadosServico.Discriminacao = objDadosServico.Discriminacao & "||VALOR TOTAL: R$ " & Replace(objDadosServico.Valores.ValorServicos, ".", ",")
                        If CDbl(objDadosServico.Valores.DescontoIncondicionado) <> 0 Then objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|DESCONTO: R$ " & Replace(objDadosServico.Valores.DescontoIncondicionado, ".", ",")
                        If CDbl(objDadosServico.Valores.ValorPis) <> 0 Then objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|PIS: R$ " & Replace(objDadosServico.Valores.ValorPis, ".", ",")
                        If CDbl(objDadosServico.Valores.ValorCofins) <> 0 Then objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|COFINS: R$ " & Replace(objDadosServico.Valores.ValorCofins, ".", ",")
                        If CDbl(objDadosServico.Valores.ValorCsll) <> 0 Then objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|CSLL: R$ " & Replace(objDadosServico.Valores.ValorCsll, ".", ",")

                        If CDbl(objDadosServico.Valores.ValorInss) <> 0 Then objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|INSS: R$ " & Replace(objDadosServico.Valores.ValorInss, ".", ",")
                        If CDbl(objDadosServico.Valores.ValorIr) <> 0 Then objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|IR: R$ " & Replace(objDadosServico.Valores.ValorIr, ".", ",")
                        If CDbl(objDadosServico.Valores.ValorIssRetido) <> 0 Then objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|ISS: R$ " & Replace(objDadosServico.Valores.ValorIssRetido, ".", ",")

                        objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|VALOR LIQUIDO: R$ " & Replace(objDadosServico.Valores.ValorLiquidoNfse, ".", ",")
                    End If
#Else


#If Not ABRASF2 Then

                    If objDadosServico.Valores.ValorLiquidoNfse <> objDadosServico.Valores.ValorServicos Then
#End If


                        objDadosServico.Discriminacao = objDadosServico.Discriminacao & "||VALOR TOTAL: R$ " & Format(objDadosServico.Valores.ValorServicos, "Fixed")
                        If objDadosServico.Valores.DescontoIncondicionado <> 0 Then objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|DESCONTO: R$ " & Format(objDadosServico.Valores.DescontoIncondicionado, "Fixed")
                        If objDadosServico.Valores.ValorPis <> 0 Then objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|PIS: R$ " & Format(objDadosServico.Valores.ValorPis, "Fixed")
                        If objDadosServico.Valores.ValorCofins <> 0 Then objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|COFINS: R$ " & Format(objDadosServico.Valores.ValorCofins, "Fixed")
                        If objDadosServico.Valores.ValorCsll <> 0 Then objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|CSLL: R$ " & Format(objDadosServico.Valores.ValorCsll, "Fixed")

                        If objDadosServico.Valores.ValorInss <> 0 Then objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|INSS: R$ " & Format(objDadosServico.Valores.ValorInss, "Fixed")
                        If objDadosServico.Valores.ValorIr <> 0 Then objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|IR: R$ " & Format(objDadosServico.Valores.ValorIr, "Fixed")
#If Not ABRASF2 Then
                        If objDadosServico.Valores.ValorIssRetido <> 0 Then objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|ISS: R$ " & Format(objDadosServico.Valores.ValorIssRetido, "Fixed")
                        objDadosServico.Discriminacao = objDadosServico.Discriminacao & "|VALOR LIQUIDO: R$ " & Format(objDadosServico.Valores.ValorLiquidoNfse, "Fixed")
#End If
#If Not ABRASF2 Then

                    End If
#End If

#End If

                        If objAdmConfig.Conteudo <> "2" Then

                            If Len(Trim(objNFiscal.MensagemCorpoNota)) <> 0 Then
                                objDadosServico.Discriminacao = objDadosServico.Discriminacao & "||" & Trim(objNFiscal.MensagemCorpoNota)
                            End If

                            If Len(Trim(objNFiscal.MensagemNota)) <> 0 Then
                                objDadosServico.Discriminacao = objDadosServico.Discriminacao & "||" & Trim(objNFiscal.MensagemNota)
                            End If

                        Else

                            sMsg = ""
                            resMensagensRegra = db1.ExecuteQuery(Of MensagensRegra) _
                                ("SELECT * FROM MensagensRegra WHERE TipoDoc = 0 And NumIntDoc = {0}", lNumIntNF)


                            For Each objMensagensRegra In resMensagensRegra
                                sMsg = sMsg & objMensagensRegra.Mensagem
                            Next

                            Replace(sMsg, "|", " ")


                            If Len(sMsg) > 0 Then
                                objDadosServico.Discriminacao = objDadosServico.Discriminacao & "||" & ADM.DesacentuaTexto(sMsg)
                            End If

                        End If

                        objDadosServico.Discriminacao = objDadosServico.Discriminacao & "||RPS: " & CStr(lNumNotaFiscal) & " Serie: " & sSerie

                        objDadosServico.Discriminacao = Replace(objDadosServico.Discriminacao, "|", vbCrLf)

                        sErro = "28"
                        sMsg1 = "vai ler FiliaisEmpresa"

                        If iDebug = 1 Then MsgBox("28")

                        resFilialEmpresa = db1.ExecuteQuery(Of FiliaisEmpresa) _
                        ("SELECT * FROM FiliaisEmpresa WHERE FilialEmpresa = {0} ", objNFiscal.FilialEmpresa)

                        For Each objFilialEmpresa In resFilialEmpresa
                            lEndereco = objFilialEmpresa.Endereco
                            Exit For
                        Next

                        sErro = "29"
                        sMsg1 = "vai ler Enderecos"

                        If iDebug = 1 Then MsgBox("29")

                        resEndereco = db1.ExecuteQuery(Of Endereco) _
                        ("SELECT * FROM Enderecos WHERE Codigo = {0}", lEndereco)

                        objEndereco = resEndereco(0)

                        sErro = "30"
                        sMsg1 = "vai ler Cidades"

                        If iDebug = 1 Then MsgBox("30")

                        resCidade = db1.ExecuteQuery(Of Cidade) _
                        ("SELECT * FROM Cidades WHERE Descricao = {0}", objEndereco.Cidade)

                        objCidade = resCidade(0)

                        sErro = "31"
                        sMsg1 = "vai ler CodTribMun"

                        If iDebug = 1 Then MsgBox("31")

                        objDadosServico.CodigoMunicipio = objCidade.CodIBGE




                        resCodTribMun = db1.ExecuteQuery(Of CodTribMun) _
                        ("SELECT * FROM CodTribMun WHERE  Cidade = {0} AND Produto = {1}", objCidade.Codigo, sProduto)

                        objCodTribMun = resCodTribMun(0)

                        If objCodTribMun Is Nothing Then Throw New System.Exception("o codigo de tributacao do municipio " & objCidade.Descricao & " para o produto " & sProduto & " não foi preenchido.")
                        If UCase(objEndFilial.Cidade) = "TATUÍ" Then
                        objDadosServico.CodigoTributacaoMunicipio = Left(objCodTribMun.CodTribMun, 7)
#If ABRASF2 Then
                            objDadosServico.MunicipioIncidenciaSpecified = True
                            objDadosServico.MunicipioIncidencia = objDadosServico.CodigoMunicipio
#End If
                        Else
                            objDadosServico.CodigoTributacaoMunicipio = objCodTribMun.CodTribMun
                        End If


#If Not ABRASF2 Then

            Dim objPrestador = New tcIdentificacaoPrestador
            objInfRps.Prestador = objPrestador

#Else
                        Dim objPrestador = New tcIdentificacaoPrestador
                        objEnviarLoteRpsEnvio.LoteRps.ListaRps.Rps.InfDeclaracaoPrestacaoServico.Prestador = objPrestador

#End If


#If Not ABRASF2 Then
                    objPrestador.Cnpj = objFilialEmpresa.CGC
#Else
                        If Len(objFilialEmpresa.CGC) = 14 Then
                            objPrestador.CpfCnpj = New tcCpfCnpj
                            objPrestador.CpfCnpj.ItemElementName = ItemChoiceType.Cnpj
                            objPrestador.CpfCnpj.Item = objFilialEmpresa.CGC
                        ElseIf Len(objFilialEmpresa.CGC) = 11 Then
                            objPrestador.CpfCnpj = New tcCpfCnpj
                            objPrestador.CpfCnpj.ItemElementName = ItemChoiceType.Cpf
                            objPrestador.CpfCnpj.Item = objFilialEmpresa.CGC
                        End If

#End If

                    If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then
                        Call ADM.Formata_String_AlfaNumerico(objFilialEmpresa.InscricaoMunicipal, objPrestador.InscricaoMunicipal)
                    Else
                        Call ADM.Formata_String_Numero(objFilialEmpresa.InscricaoMunicipal, objPrestador.InscricaoMunicipal)
                    End If


#If Not ABRASF2 Then

            Dim objDadosTomador = New tcDadosTomador
            objInfRps.Tomador = objDadosTomador

#Else
                    Dim objDadosTomador = New tcDadosTomador
                    objEnviarLoteRpsEnvio.LoteRps.ListaRps.Rps.InfDeclaracaoPrestacaoServico.Tomador = objDadosTomador

#End If



                    objDadosTomador.Endereco = New tcEndereco

                    sErro = "32"
                    sMsg1 = "vai ler FiliaisClientes"

                    If iDebug = 1 Then MsgBox("32")


                    resFiliaisClientes = db1.ExecuteQuery(Of FiliaisCliente) _
                    ("SELECT * FROM FiliaisClientes WHERE CodCliente = {0} AND CodFilial = {1}", objNFiscal.Cliente, objNFiscal.FilialCli)

                    For Each objFiliaisClientes In resFiliaisClientes


                        sErro = "33"
                        sMsg1 = "leu FiliaisClientes"

                        If iDebug = 1 Then MsgBox("33")

                        lEndDest = objFiliaisClientes.Endereco

                        If Len(objFiliaisClientes.CGC) = 11 Then
                            objDadosTomador.IdentificacaoTomador = New tcIdentificacaoTomador

                            objDadosTomador.IdentificacaoTomador.CpfCnpj = New tcCpfCnpj
                            objDadosTomador.IdentificacaoTomador.CpfCnpj.ItemElementName = ItemChoiceType.Cpf
                            objDadosTomador.IdentificacaoTomador.CpfCnpj.Item = objFiliaisClientes.CGC
                        ElseIf Len(objFiliaisClientes.CGC) = 14 Then
                            objDadosTomador.IdentificacaoTomador = New tcIdentificacaoTomador

                            objDadosTomador.IdentificacaoTomador.CpfCnpj = New tcCpfCnpj
                            objDadosTomador.IdentificacaoTomador.CpfCnpj.ItemElementName = ItemChoiceType.Cnpj
                            objDadosTomador.IdentificacaoTomador.CpfCnpj.Item = objFiliaisClientes.CGC
                        End If



                        If Trim(objFiliaisClientes.InscricaoMunicipal) = "" Then
                            sIM = ""
                        Else
                            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then
                                Call ADM.Formata_String_AlfaNumerico(objFiliaisClientes.InscricaoMunicipal, sIM)
                            Else
                                Call ADM.Formata_String_Numero(objFiliaisClientes.InscricaoMunicipal, sIM)
                            End If
                        End If


                        Exit For
                    Next

                    sErro = "34"
                    sMsg1 = "vai ler Clientes"

                    If iDebug = 1 Then MsgBox("34")

                    resCliente = db1.ExecuteQuery(Of Cliente) _
                    ("SELECT * FROM Clientes WHERE Codigo = {0}", objNFiscal.Cliente)

                    For Each objCliente In resCliente
                        objDadosTomador.RazaoSocial = ADM.DesacentuaTexto(Trim(objCliente.RazaoSocial))
                        Exit For
                    Next

                    sErro = "35"
                    sMsg1 = "vai ler Endereco de Destino"


                    If iDebug = 1 Then MsgBox("35")

                    resEndDest = db1.ExecuteQuery(Of Endereco) _
                    ("SELECT * FROM Enderecos WHERE Codigo = {0}", lEndDest)

                    For Each objEndDest In resEndDest

                        resEstado = db1.ExecuteQuery(Of Estado) _
                        ("SELECT * FROM Estados WHERE Sigla = {0}", objEndDest.SiglaEstado)

                        For Each objEstado In resEstado

                            Exit For
                        Next

                        sErro = "40"
                        sMsg1 = "vai ler Cidade de Destino"


                        If iDebug = 1 Then MsgBox("40")


                        If objEstado Is Nothing Then Throw New System.Exception("o Estado do Destinatário não foi encontrado. Estado = " & objEndDest.SiglaEstado)

                        'se o Estado do tomador for o mesmo do prestador
                        If objEndereco.SiglaEstado = objEstado.Sigla Then
                            If Len(sIM) > 0 Then
                                objDadosTomador.IdentificacaoTomador.InscricaoMunicipal = sIM
                            End If

                        End If

                        resCidade = db1.ExecuteQuery(Of Cidade) _
                        ("SELECT * FROM Cidades WHERE Descricao = {0}", objEndDest.Cidade)

                        For Each objCidade In resCidade
                            Exit For
                        Next

                        sErro = "41"
                        sMsg1 = "vai transferir os dados de endereco destino"

                        If iDebug = 1 Then MsgBox("41")


                        If Len(objEndDest.Logradouro) > 0 Then


                            objDadosTomador.Endereco.Endereco = ADM.DesacentuaTexto(Left(IIf(Len(objEndDest.TipoLogradouro) > 0, objEndDest.TipoLogradouro & " ", "") & objEndDest.Logradouro, 60))
                            objDadosTomador.Endereco.Numero = objEndDest.Numero
                            objDadosTomador.Endereco.Complemento = ADM.DesacentuaTexto(objEndDest.Complemento)
                        Else
                            objDadosTomador.Endereco.Endereco = ADM.DesacentuaTexto(objEndDest.Endereco)
                            objDadosTomador.Endereco.Numero = "0"
                        End If


                        If Len(objEndDest.TelNumero1) > 0 Then
                            If objDadosTomador.Contato Is Nothing Then objDadosTomador.Contato = New tcContato
                            Call ADM.Formata_String_Numero(IIf(Len(CStr(objEndDest.TelDDD1)) > 0, CStr(objEndDest.TelDDD1), "") + objEndDest.TelNumero1, objDadosTomador.Contato.Telefone)
                            If Len(objDadosTomador.Contato.Telefone) > 11 Then
                                objDadosTomador.Contato.Telefone = Left(objDadosTomador.Contato.Telefone, 10)
                            End If
                        ElseIf Len(objEndDest.Telefone1) > 0 Then
                            If objDadosTomador.Contato Is Nothing Then objDadosTomador.Contato = New tcContato
                            Call ADM.Formata_String_Numero(objEndDest.Telefone1, objDadosTomador.Contato.Telefone)
                            If Len(objDadosTomador.Contato.Telefone) > 11 Then
                                objDadosTomador.Contato.Telefone = Left(objDadosTomador.Contato.Telefone, 10)
                            End If
                        End If

                        sErro = "42"
                        sMsg1 = "vai continuar transferindo os dados de endereco destino"

                        If iDebug = 1 Then MsgBox("42")

                        If Len(objEndDest.Email) > 0 Then
                            If objDadosTomador.Contato Is Nothing Then objDadosTomador.Contato = New tcContato
                            objDadosTomador.Contato.Email = objEndDest.Email
                        End If

                        objDadosTomador.Endereco.Bairro = ADM.DesacentuaTexto(objEndDest.Bairro)

                        If Len(objCidade.CodIBGE) = 0 Then Throw New System.Exception("O codigo IBGE da cidade do endereço do tomador não foi setado. Cidade = " & objCidade.Descricao)
                        objDadosTomador.Endereco.CodigoMunicipio = objCidade.CodIBGE
                        objDadosTomador.Endereco.CodigoMunicipioSpecified = True

                        If Len(objEndDest.SiglaEstado) <> 2 Then Throw New System.Exception("A UF do tomador do serviço tem sigla diferente de 2 caracteres")

                        objDadosTomador.Endereco.Uf = objEndDest.SiglaEstado

                        If Len(objEndDest.CEP) > 0 Then
                            Call ADM.Formata_String_Numero(objEndDest.CEP, objDadosTomador.Endereco.Cep)
#If Not ABRASF2 Then
                            objDadosTomador.Endereco.CepSpecified = True
#End If
                        End If

                        sErro = "43"
                        sMsg1 = "transferiu os dados de endereco de destino"

                        If iDebug = 1 Then MsgBox("43")

                        Exit For
                    Next

                    Exit For

                Next

#If ABRASF2 Then
                Dim objRps As tcDeclaracaoPrestacaoServico

                objRps = objEnviarLoteRpsEnvio.LoteRps.ListaRps.Rps
#Else

                Dim objRps As tcRps = New tcRps

                objRps.InfRps = objInfRps

#End If


                Dim AD As AssinaturaDigital = New AssinaturaDigital

                sErro = "44"
                sMsg1 = "vai serializar e assinar a msg"

                If iDebug = 1 Then MsgBox("44")

#If ABRASF2 Then
                Dim mySerializer As New XmlSerializer(GetType(tcDeclaracaoPrestacaoServico))
#Else
                Dim mySerializer As New XmlSerializer(GetType(tcRps))
#End If
                XMLStream = New MemoryStream(10000)

                mySerializer.Serialize(XMLStream, objRps)

                Dim xm1 As Byte()
                xm1 = XMLStream.ToArray

                XMLString = System.Text.Encoding.UTF8.GetString(xm1)

                If UCase(objEndFilial.Cidade) = "SALVADOR" Then
                    '    iPos1 = InStr(XMLString, "<tcRps")
                    '    iPos2 = InStr(iPos1, XMLString, ">")

                    '    XMLString = Mid(XMLString, 1, iPos1 + 5) & Mid(XMLString, iPos2)

                    '    iPos1 = InStr(XMLString, "xmlns")
                    '    iPos2 = InStr(iPos1, XMLString, ">")

                    '    XMLString = Mid(XMLString, 1, iPos1 - 2) & Mid(XMLString, iPos2)
                    XMLString = Replace(XMLString, "Id=", "id=")

                End If

                If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then
                    XMLString = Replace(XMLString, "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd", "http://www.abrasf.org.br/nfse.xsd")
                End If


                If iDebug = 1 Then

                    Dim xDados100 As Byte()
                    Dim XMLStreamDados100 As MemoryStream = New MemoryStream(10000)

                    xDados100 = System.Text.Encoding.UTF8.GetBytes(XMLString)

                    XMLStreamDados100.Write(xDados100, 0, xDados100.Length)

                    Dim DocDados100 As XmlDocument = New XmlDocument

                    XMLStreamDados100.Position = 0
                    DocDados100.Load(XMLStreamDados100)
                    sArquivo = sDir & "assina_" & objInfRps.Id & ".xml"

                    Dim writer100 As New XmlTextWriter(sArquivo, Nothing)

                    writer100.Formatting = Formatting.None
                    DocDados100.WriteTo(writer100)
                    writer100.Close()

                    MsgBox(sArquivo)

                End If


#If ABRASF2 Then

                iResult = AD.Assinar(XMLString, "InfDeclaracaoPrestacaoServico", cert, objEndFilial.Cidade)
#Else

                iResult = AD.Assinar(XMLString, "InfRps", cert, objEndFilial.Cidade)
#End If

                If iResult <> 0 Then Throw New System.Exception("Ocorreu um erro durante a assinatura do lote. " & AD.mensagemResultado)

                sErro = "45"
                sMsg1 = "vai pegar a msg assinada"

                If iDebug = 1 Then MsgBox("45")

                Dim xMlAD As XmlDocument

                xMlAD = AD.XMLDocAssinado()

                Dim xString10 As String
                xString10 = AD.XMLStringAssinado


#If ABRASF2 Then
                xString10 = Replace(xString10, "<tcDeclaracaoPrestacaoServico", "<Rps")

                xString10 = Replace(xString10, "</tcDeclaracaoPrestacaoServico", "</Rps")
#Else
                xString10 = Replace(xString10, "<tcRps", "<Rps")

                xString10 = Replace(xString10, "</tcRps", "</Rps")
#End If

                XMLStringRpses = XMLStringRpses & Mid(xString10, 22) & " "

                sErro = "46"
                sMsg1 = "vai gravar o xml"

                If iDebug = 1 Then MsgBox("46")


                '****************  salva o arquivo 

                XMLStreamDados = New MemoryStream(10000)

                Dim xDados1 As Byte()

                xDados1 = System.Text.Encoding.UTF8.GetBytes(Mid(xString10, 22))

                XMLStreamDados.Write(xDados1, 0, xDados1.Length)

                Dim DocDados1 As XmlDocument = New XmlDocument

                XMLStreamDados.Position = 0
                DocDados1.Load(XMLStreamDados)
                sArquivo = sDir & objInfRps.Id & ".xml"

                Dim writer As New XmlTextWriter(sArquivo, Nothing)

                writer.Formatting = Formatting.None
                DocDados1.WriteTo(writer)

                writer.Close()

                Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 1

                sErro = "47"
                sMsg1 = "gravou o xml"

                If iDebug = 1 Then MsgBox("47")


                '                    DocDados1.Save(sArquivo)


            Next


            Dim a5 As EnviarLoteRpsEnvio = New EnviarLoteRpsEnvio

            Dim objLoteRPS As tcLoteRps = New tcLoteRps

            a5.LoteRps = objLoteRPS

            a5.LoteRps.NumeroLote = lLote
            a5.LoteRps.Id = "lote_rps_" & lLote
#If Not ABRASF2 Then
            a5.LoteRps.Cnpj = objFilialEmpresa.CGC
#Else
            If Len(objFilialEmpresa.CGC) = 14 Then
                a5.LoteRps.CpfCnpj = New tcCpfCnpj
                a5.LoteRps.CpfCnpj.ItemElementName = ItemChoiceType.Cnpj
                a5.LoteRps.CpfCnpj.Item = objFilialEmpresa.CGC
            ElseIf Len(objFilialEmpresa) = 11 Then
                a5.LoteRps.CpfCnpj = New tcCpfCnpj
                a5.LoteRps.CpfCnpj.ItemElementName = ItemChoiceType.Cpf
                a5.LoteRps.CpfCnpj.Item = objFilialEmpresa.CGC
            End If


#End If
            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then
#If ABRASF Then
                a5.LoteRps.versao = "1.00"
#End If
            End If


            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then
                Call ADM.Formata_String_AlfaNumerico(objFilialEmpresa.InscricaoMunicipal, a5.LoteRps.InscricaoMunicipal)
            Else
                Call ADM.Formata_String_Numero(objFilialEmpresa.InscricaoMunicipal, a5.LoteRps.InscricaoMunicipal)
            End If

            sErro = "48"
            sMsg1 = "vai ler a tabela RPSWEBLote"

            If iDebug = 1 Then MsgBox("48")


            resRPSWEBLote = db2.ExecuteQuery(Of RPSWEBLote) _
            ("SELECT * FROM RPSWEBLote WHERE Lote = {0} ORDER BY NumIntNF", lLote)

            a5.LoteRps.QuantidadeRps = resRPSWEBLote.Count


            sErro = "49"
            sMsg1 = "vai serializar EnviarLoteRpsEnvio e assinar o lote"

            If iDebug = 1 Then MsgBox("49")

#If ABRASF2 Then
            Dim objListaRps As tcLoteRpsListaRps
            objListaRps = New tcLoteRpsListaRps
#Else
            Dim objListaRps(0) As tcRps
            objListaRps(0) = New tcRps
#End If
            a5.LoteRps.ListaRps = objListaRps


            Dim mySerializer1 As New XmlSerializer(GetType(EnviarLoteRpsEnvio))

            XMLStream = New MemoryStream(10000)

            mySerializer1.Serialize(XMLStream, a5)

            Dim xm As Byte()
            xm = XMLStream.ToArray

            XMLString = System.Text.Encoding.UTF8.GetString(xm)

#If ABRASF2 Then
            XMLStringRpses = "<ListaRps>" & XMLStringRpses & "</ListaRps>"
            XMLString = Replace(XMLString, "<ListaRps />", XMLStringRpses)
#Else
            XMLString = Replace(XMLString, "<Rps />", XMLStringRpses)
#End If

            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then
                XMLString = Replace(XMLString, "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd", "http://www.abrasf.org.br/nfse.xsd")
            End If

            If UCase(objEndFilial.Cidade) = "SALVADOR" Then

                '    iPos1 = InStr(XMLString, "xmlns")
                '    iPos2 = InStr(XMLString, "xmlns=")
                '    XMLString = Mid(XMLString, 1, iPos1 - 1) & Mid(XMLString, iPos2)

                '    iPos1 = InStr(XMLString, "?>")
                '    XMLString = Mid(XMLString, 1, iPos1 - 1) & " encoding=""utf-8""" & Mid(XMLString, iPos1)
                XMLString = Replace(XMLString, "Id=", "id=")

            End If


            Dim AD1 As AssinaturaDigital = New AssinaturaDigital

            AD1.Assinar(XMLString, "LoteRps", cert, objEndFilial.Cidade)

            sErro = "50"
            sMsg1 = "assinou o lote"

            If iDebug = 1 Then MsgBox("50")

            Dim xMlD As XmlDocument

            xMlD = AD1.XMLDocAssinado()

            Dim xString As String
            xString = AD1.XMLStringAssinado

            '            If UCase(objEndFilial.Cidade) <> "SALVADOR" Then
            xString = Mid(xString, 22)
            '            End If


            sErro = "51"
            sMsg1 = "vai gravar xml do lote"


            If iDebug = 1 Then MsgBox("51")



            sSerie = ""
            lNumNotaFiscal = 0

            'iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5})", _
            'iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "Iniciando a validação do lote", 0)

            'Form1.Msg.Items.Add("Iniciando a validação do lote")

            'Application.DoEvents()

            lNumIntNF = 0

            'envioNFe.versao = "1.10"
            'envioNFe.idLote = lLote

            'Dim mySerializerw As New XmlSerializer(GetType(TEnviNFe))

            'XMLStream1 = New MemoryStream(10000)

            'mySerializerw.Serialize(XMLStream1, envioNFe)

            'Dim xmw As Byte()
            'xmw = XMLStream1.ToArray

            'XMLString1 = System.Text.Encoding.UTF8.GetString(xmw)

            'XMLString2 = Mid(XMLString1, 1, Len(XMLString1) - 10) & XMLStringNFes & Mid(XMLString1, Len(XMLString1) - 10)

            'XMLString2 = Mid(XMLString2, 1, 19) & " encoding=""utf-8"" " & Mid(XMLString2, 20)

            'Dim XMLStringRetEnvNFE As String

            ''Load the client certificate from a file.
            ''Dim x509 As X509Certificate = X509Certificate.CreateFromSignedFile("c:\nfe\ecnpj.cer")

            '************* valida dados antes do envio **********************
            Dim xDados As Byte()

            ''            xDados = System.Text.Encoding.UTF8.GetBytes(xString)
            xDados = System.Text.Encoding.UTF8.GetBytes(xString)

            XMLStreamDados = New MemoryStream(10000)

            XMLStreamDados.Write(xDados, 0, xDados.Length)

            Dim DocDados As XmlDocument = New XmlDocument
            XMLStreamDados.Position = 0
            DocDados.Load(XMLStreamDados)
            sArquivo = sDir & a5.LoteRps.Id & ".xml"
            DocDados.Save(sArquivo)

            sErro = "52"
            sMsg1 = "vai gravar RPSWEBLoteLog"

            If iDebug = 1 Then MsgBox("52")

            Dim lErro As Long

            'lErro = objValidaXML.validaXML(sArquivo, sDir1 & "\nfse.xsd", lLote, lNumIntNF, db2, iFilialEmpresa)
            '            lErro = objValidaXML.validaXML(sArquivo, "c:\nfeservico\tatui\schemas_v301\servico_enviar_lote_rps_envio_v03.xsd", lLote, lNumIntNF, db2, iFilialEmpresa)
            '            If lErro = 1 Then

            '    '    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5})", _
            '    '    iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "ERRO - Encerrado o envio deste lote", 0)

            '    '    Form1.Msg.Items.Add("ERRO - o envio do lote " & CStr(lLote) & " foi encerrado por erro.")

            '    '    Application.DoEvents()

            '    '    Exit Try
            '           End If


            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "Iniciando o envio do lote", 0)

            sErro = "53"
            sMsg1 = "vai iniciar o envio do lote"

            If iDebug = 1 Then MsgBox("53")


            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            Form1.Msg.Items.Add("Iniciando o envio do lote")

            If Form1.Msg.Items.Count - 15 < 1 Then
                Form1.Msg.TopIndex = 1
            Else
                Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
            End If

            Application.DoEvents()

            Dim XMLStringRetEnvRPS As String

            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then

                XMLStringCabec = Replace(XMLStringCabec, "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd", "http://www.abrasf.org.br/nfse.xsd")


                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homologacaobh.ClientCertificates.Add(cert)
                    'homologacaobh.Endpoint.Address=
                    XMLStringRetEnvRPS = homologacaobh.RecepcionarLoteRps(XMLStringCabec, xString)
                Else
                    '  producaobh.ClientCredentials.ClientCertificate.Certificate = cert
                    '  producaobh.ClientCredentials.ServiceCertificate.DefaultCertificate = cert
                    producaobh.ClientCertificates.Add(cert)
                    XMLStringRetEnvRPS = producaobh.RecepcionarLoteRps(XMLStringCabec, xString)
                End If

                XMLStringRetEnvRPS = Replace(XMLStringRetEnvRPS, "http://www.abrasf.org.br/nfse.xsd", "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd")

            ElseIf UCase(objEndFilial.Cidade) = "SALVADOR" Then

                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homosalvadorEnvio.ClientCertificates.Add(cert)
                    XMLStringRetEnvRPS = homosalvadorEnvio.EnviarLoteRPS(xString)
                Else
                    prodsalvadorEnvio.ClientCertificates.Add(cert)
                    XMLStringRetEnvRPS = prodsalvadorEnvio.EnviarLoteRPS(xString)
                End If

            ElseIf UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then

                ' XMLStringCabec = Replace(XMLStringCabec, "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd", "http://www.ginfes.com.br/servico_enviar_lote_rps_envio_v03.xsd")


                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homologacaoginfes.ClientCertificates.Add(cert)
                    XMLStringRetEnvRPS = homologacaoginfes.RecepcionarLoteRpsV3(XMLStringCabec, xString)
                Else
                    producaoginfes.ClientCertificates.Add(cert)
                    XMLStringRetEnvRPS = producaoginfes.RecepcionarLoteRpsV3(XMLStringCabec, xString)
                End If

                '   XMLStringRetEnvRPS = Replace(XMLStringRetEnvRPS, "http://www.abrasf.org.br/nfse.xsd", "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd")

                XMLStringRetEnvRPS = Replace(XMLStringRetEnvRPS, "ListaMensagemRetorno", "ns2:ListaMensagemRetorno")


            ElseIf UCase(objEndFilial.Cidade) = "TATUÍ" Then

                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homsistema4rEnvio.ClientCertificates.Add(cert)
                    XMLStringRetEnvRPS = homsistema4rEnvio.Execute(xString)
                Else
                    prodsistema4rEnvio.ClientCertificates.Add(cert)
                    XMLStringRetEnvRPS = prodsistema4rEnvio.Execute(xString)
                End If


            Else

                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homologacao.ClientCredentials.ClientCertificate.Certificate = cert
                    XMLStringRetEnvRPS = homologacao.RecepcionarLoteRps(xString)
                Else
                    producao.ClientCredentials.ClientCertificate.Certificate = cert
                    XMLStringRetEnvRPS = producao.RecepcionarLoteRps(xString)
                End If

            End If


            sErro = "55"
            sMsg1 = "vai deserializar a resposta"


            If iDebug = 1 Then MsgBox("55")

            Dim xRet1 As Byte()



            xRet1 = System.Text.Encoding.UTF8.GetBytes(XMLStringRetEnvRPS)

            XMLStreamRet = New MemoryStream(10000)

            XMLStreamRet.Write(xRet1, 0, xRet1.Length)



#If ABRASF2 Then
            Dim mySerializerRetEnvNFe As New XmlSerializer(GetType(EnviarLoteRpsSincronoResposta))
            Dim objRetEnviRPS As EnviarLoteRpsSincronoResposta = New EnviarLoteRpsSincronoResposta
#Else
            Dim mySerializerRetEnvNFe As New XmlSerializer(GetType(EnviarLoteRpsResposta))
            Dim objRetEnviRPS As EnviarLoteRpsResposta = New EnviarLoteRpsResposta
#End If


            XMLStreamRet.Position = 0

            objRetEnviRPS = mySerializerRetEnvNFe.Deserialize(XMLStreamRet)

            sErro = "56"
            sMsg1 = "vai tratar os varios tipos de resposta"

            If iDebug = 1 Then MsgBox("56")


#If Not ABRASF2 Then
            For iIndice1 = 0 To objRetEnviRPS.Items.Count - 1

                dtDataRecebimento = CDate("07/09/1822")
                dHoraRecebimento = 0.5
#End If

#If ABRASF2 Then
            dtDataRecebimento = objRetEnviRPS.DataRecebimento
            dHoraRecebimento = 0.5
            sProtocolo = objRetEnviRPS.Protocolo
            sLote = objRetEnviRPS.NumeroLote
            sAux = objRetEnviRPS.Item.GetType.ToString()
            If InStr(sAux, ".") <> 0 Then
                sAux = Mid(sAux, InStr(sAux, ".") + 1)
            End If

            If sAux = "ListaMensagemRetorno" Then
                objListaMsgRetorno = objRetEnviRPS.Item

#ElseIf GINFES Or RJ Then
                If objRetEnviRPS.ItemsElementName(iIndice1) = ItemsChoiceType3.DataRecebimento Then
                    dtDataRecebimento = CDate(objRetEnviRPS.Items(iIndice1)).Date
                    dHoraRecebimento = CDate(Format(objRetEnviRPS.Items(iIndice1), "T")).ToOADate()
                ElseIf objRetEnviRPS.ItemsElementName(iIndice1) = ItemsChoiceType3.Protocolo Then
                    sProtocolo = objRetEnviRPS.Items(iIndice1)
                ElseIf objRetEnviRPS.ItemsElementName(iIndice1) = ItemsChoiceType3.NumeroLote Then
                    sLote = objRetEnviRPS.Items(iIndice1)
                ElseIf objRetEnviRPS.ItemsElementName(iIndice1) = ItemsChoiceType3.ListaMensagemRetorno Then
                    objListaMsgRetorno = objRetEnviRPS.Items(iIndice1)

#ElseIf ABRASF Then
                If objRetEnviRPS.ItemsElementName(iIndice1) = ItemsChoiceType.DataRecebimento Then
                    dtDataRecebimento = CDate(objRetEnviRPS.Items(iIndice1)).Date
                    dHoraRecebimento = CDate(Format(objRetEnviRPS.Items(iIndice1), "T")).ToOADate()
                ElseIf objRetEnviRPS.ItemsElementName(iIndice1) = ItemsChoiceType.Protocolo Then
                    sProtocolo = objRetEnviRPS.Items(iIndice1)
                ElseIf objRetEnviRPS.ItemsElementName(iIndice1) = ItemsChoiceType.NumeroLote Then
                    sLote = objRetEnviRPS.Items(iIndice1)
                ElseIf objRetEnviRPS.ItemsElementName(iIndice1) = ItemsChoiceType.ListaMensagemRetorno Then

                                objListaMsgRetorno = objRetEnviRPS.Items(iIndice1)

#End If

                If Not objListaMsgRetorno.MensagemRetorno Is Nothing Then
                    For iIndice = 0 To objListaMsgRetorno.MensagemRetorno.Count - 1
                        objMsgRetorno = objListaMsgRetorno.MensagemRetorno(iIndice)

                        sLote = lLote
                        sProtocolo = ""
                        objMsgRetorno.Correcao = ""

                        sErro = "57"
                        sMsg1 = "vai gravar RPSWEBRetEnvi"

                        If iDebug = 1 Then MsgBox("57")


                        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBRetEnvi ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Correcao, data, hora, datarecebimento, horarecebimento) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11} )", _
                        lRPSWEBRetEnviNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objMsgRetorno.Codigo, Left(objMsgRetorno.Mensagem, 200), IIf(objMsgRetorno.Correcao Is Nothing, "", objMsgRetorno.Correcao), Now.Date, TimeOfDay.ToOADate, dtDataRecebimento, dHoraRecebimento)

                        lRPSWEBRetEnviNumIntDoc = lRPSWEBRetEnviNumIntDoc + 1

                        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                        lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Retorno do envio do lote - " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & IIf(objMsgRetorno.Correcao Is Nothing, "", objMsgRetorno.Correcao), 255), 0)

                        lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                        Form1.Msg.Items.Add("Retorno do envio do lote - " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & IIf(objMsgRetorno.Correcao Is Nothing, "", objMsgRetorno.Correcao))

                        If Form1.Msg.Items.Count - 15 < 1 Then
                            Form1.Msg.TopIndex = 1
                        Else
                            Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                        End If

                        Application.DoEvents()

                        sErro = "58"
                        sMsg1 = "tratou RPSWEBRetEnvi"


                        If iDebug = 1 Then MsgBox("58")

                    Next
                    Throw New System.Exception("Aconteceu um erro no envio. Veja mensagens acima.")
                End If

                If iDebug = 1 Then MsgBox("59")


            End If

#If ABRASF2 Then

            If sAux = "ListaMensagemRetornoLote" Then

                objListaMsgRetornoLote = objRetEnviRPS.Item


                If Not objListaMsgRetornoLote.MensagemRetorno Is Nothing Then
                    For iIndice = 0 To objListaMsgRetornoLote.MensagemRetorno.Count - 1
                        objMsgRetornoLote = objListaMsgRetornoLote.MensagemRetorno(iIndice)

                        sLote = lLote
                        sProtocolo = ""
                        objMsgRetorno.Correcao = ""

                        sErro = "59.1"
                        sMsg1 = "vai gravar RPSWEBRetEnvi"

                        If iDebug = 1 Then MsgBox("59.1")


                        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBRetEnvi ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Correcao, data, hora, datarecebimento, horarecebimento) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11} )", _
                        lRPSWEBRetEnviNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objMsgRetornoLote.Codigo, Left(objMsgRetornoLote.Mensagem, 200), IIf(objMsgRetornoLote.IdentificacaoRps Is Nothing, "", "RPS - Numero: " & objMsgRetornoLote.IdentificacaoRps.Numero & " Serie: " & objMsgRetornoLote.IdentificacaoRps.Serie & " Tipo: " & objMsgRetornoLote.IdentificacaoRps.Tipo), Now.Date, TimeOfDay.ToOADate, dtDataRecebimento, dHoraRecebimento)

                        lRPSWEBRetEnviNumIntDoc = lRPSWEBRetEnviNumIntDoc + 1

                        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                        lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Retorno do envio do lote - " & objMsgRetornoLote.Codigo & " - " & objMsgRetornoLote.Mensagem & " - " & IIf(objMsgRetornoLote.IdentificacaoRps Is Nothing, "", "RPS - Numero: " & objMsgRetornoLote.IdentificacaoRps.Numero & " Serie: " & objMsgRetornoLote.IdentificacaoRps.Serie & " Tipo: " & objMsgRetornoLote.IdentificacaoRps.Tipo), 255), 0)

                        lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                        Form1.Msg.Items.Add("Retorno do envio do lote - " & objMsgRetornoLote.Codigo & " - " & objMsgRetornoLote.Mensagem & " - " & IIf(objMsgRetornoLote.IdentificacaoRps Is Nothing, "", "RPS - Numero: " & objMsgRetornoLote.IdentificacaoRps.Numero & " Serie: " & objMsgRetornoLote.IdentificacaoRps.Serie & " Tipo: " & objMsgRetornoLote.IdentificacaoRps.Tipo))

                        If Form1.Msg.Items.Count - 15 < 1 Then
                            Form1.Msg.TopIndex = 1
                        Else
                            Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                        End If

                        Application.DoEvents()

                        sErro = "59.2"
                        sMsg1 = "tratou RPSWEBRetEnvi"


                        If iDebug = 1 Then MsgBox("59.2")

                    Next
                    Throw New System.Exception("Aconteceu um erro no envio. Veja mensagens acima.")
                End If
            End If

#End If

#If Not ABRASF2 Then
            Next
#End If

            'vai iniciar a consulta a situacao do lote 
            If sProtocolo <> "" Then

                objMsgRetorno.Codigo = ""
                objMsgRetorno.Mensagem = "Lote enviado"
                objMsgRetorno.Correcao = ""

                sErro = "60"
                sMsg1 = "vai gravar RPSWEBRetEnvi"


                If iDebug = 1 Then MsgBox("60")


                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBRetEnvi ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Correcao, data, hora, datarecebimento, horarecebimento) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11} )", _
                lRPSWEBRetEnviNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objMsgRetorno.Codigo, Left(objMsgRetorno.Mensagem, 200), objMsgRetorno.Correcao, Now.Date, TimeOfDay.ToOADate, dtDataRecebimento, dHoraRecebimento)

                lRPSWEBRetEnviNumIntDoc = lRPSWEBRetEnviNumIntDoc + 1

                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Retorno do envio do lote - " & objMsgRetorno.Mensagem & " protocolo = " & sProtocolo, 255), 0)


                sErro = "61"
                sMsg1 = "vai gravar RPSWEBRetEnvi e iniciar a consulta de situacao do lote"

                If iDebug = 1 Then MsgBox("61")

                lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                Form1.Msg.Items.Add("Retorno do envio do lote - " & objMsgRetorno.Mensagem & " protocolo = " & sProtocolo)

                If Form1.Msg.Items.Count - 15 < 1 Then
                    Form1.Msg.TopIndex = 1
                Else
                    Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                End If

                Application.DoEvents()

                '    If 1 = 2 Then

                'If UCase(objEndFilial.Cidade) <> "SALVADOR" Then


                '    Dim objConsSituacaoLoteRpsEnvio As ConsultarSituacaoLoteRpsEnvio = New ConsultarSituacaoLoteRpsEnvio


                '    objConsSituacaoLoteRpsEnvio.Prestador = New tcIdentificacaoPrestador

                '    objConsSituacaoLoteRpsEnvio.Prestador.Cnpj = objFilialEmpresa.CGC
                '    Call ADM.Formata_String_Numero(objFilialEmpresa.InscricaoMunicipal, objConsSituacaoLoteRpsEnvio.Prestador.InscricaoMunicipal)

                '    objConsSituacaoLoteRpsEnvio.Protocolo = sProtocolo


                '    Dim mySerializerx2 As New XmlSerializer(GetType(ConsultarSituacaoLoteRpsEnvio))

                '    XMLStream1 = New MemoryStream(10000)
                '    mySerializerx2.Serialize(XMLStream1, objConsSituacaoLoteRpsEnvio)

                '    Dim xm11 As Byte()
                '    xm11 = XMLStream1.ToArray

                '    XMLString1 = System.Text.Encoding.UTF8.GetString(xm11)

                '    XMLString1 = Mid(XMLString1, 1, 19) & " encoding=""utf-8"" " & Mid(XMLString1, 20)


                '    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                '    lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "Iniciando a consulta da situação do lote", 0)

                '    lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                '    Form1.Msg.Items.Add("Iniciando a consulta de situacao do lote - Aguarde")

                '    If Form1.Msg.Items.Count - 15 < 1 Then
                '        Form1.Msg.TopIndex = 1
                '    Else
                '        Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                '    End If

                '    Application.DoEvents()

                '    Dim XMLStringConsultarSituacaoLoteRPSResposta As String


                '    If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then

                '        XMLString1 = Replace(XMLString1, "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd", "http://www.abrasf.org.br/nfse.xsd")

                '    End If

                '    If UCase(objEndFilial.Cidade) = "TATUÍ" Or UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then

                '        Dim AD2 As AssinaturaDigital = New AssinaturaDigital

                '        AD2.Assinar(XMLString1, "Protocolo", cert, objEndFilial.Cidade)

                '        sErro = "50"
                '        sMsg1 = "vai assinar a msg de verficacao de situacao do lote de tatui"


                '        If iDebug = 1 Then MsgBox("50")

                '        Dim xMlD1 As XmlDocument

                '        xMlD1 = AD2.XMLDocAssinado()

                '        XMLString1 = AD2.XMLStringAssinado

                '        '                        XMLString1 = Mid(XMLString1, 22)
                '        '            End If

                '    End If

                '    If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then


                '        If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then



                '            homologacaobh.ClientCertificates.Add(cert)
                '            XMLStringConsultarSituacaoLoteRPSResposta = homologacaobh.ConsultarSituacaoLoteRps(XMLStringCabec, XMLString1)
                '        Else
                '            '                    producao.ClientCredentials.ClientCertificate.Certificate = cert
                '            producaobh.ClientCertificates.Add(cert)
                '            XMLStringConsultarSituacaoLoteRPSResposta = producaobh.ConsultarSituacaoLoteRps(XMLStringCabec, XMLString1)
                '        End If

                '        XMLStringConsultarSituacaoLoteRPSResposta = Replace(XMLStringConsultarSituacaoLoteRPSResposta, "http://www.abrasf.org.br/nfse.xsd", "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd")


                '    ElseIf UCase(objEndFilial.Cidade) = "SALVADOR" Then

                '        If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                '            'homologacao.ClientCredentials.ClientCertificate.Certificate = cert
                '            XMLStringConsultarSituacaoLoteRPSResposta = homosalvadorSitLote.ConsultarSituacaoLoteRPS(XMLString1)
                '        Else
                '            '                producao.ClientCredentials.ClientCertificate.Certificate = cert
                '            XMLStringConsultarSituacaoLoteRPSResposta = prodsalvadorSitLote.ConsultarSituacaoLoteRPS(XMLString1)
                '        End If

                '    ElseIf UCase(objEndFilial.Cidade) = "TATUÍ" Or UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then


                '        If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then

                '            homologacaoginfes.ClientCertificates.Add(cert)
                '            XMLStringConsultarSituacaoLoteRPSResposta = homologacaoginfes.ConsultarSituacaoLoteRpsV3(XMLStringCabec, XMLString1)
                '        Else
                '            '                    producao.ClientCredentials.ClientCertificate.Certificate = cert
                '            producaoginfes.ClientCertificates.Add(cert)
                '            XMLStringConsultarSituacaoLoteRPSResposta = producaoginfes.ConsultarSituacaoLoteRpsV3(XMLStringCabec, XMLString1)
                '        End If

                '        'XMLStringConsultarSituacaoLoteRPSResposta = Replace(XMLStringConsultarSituacaoLoteRPSResposta, "http://www.abrasf.org.br/nfse.xsd", "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd")
                '        XMLStringConsultarSituacaoLoteRPSResposta = Replace(XMLStringConsultarSituacaoLoteRPSResposta, "ListaMensagemRetorno", "ns2:ListaMensagemRetorno")
                '    Else

                '        If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                '            'homologacao.ClientCredentials.ClientCertificate.Certificate = cert
                '            XMLStringConsultarSituacaoLoteRPSResposta = homologacao.ConsultarSituacaoLoteRps(XMLString1)
                '        Else
                '            '                producao.ClientCredentials.ClientCertificate.Certificate = cert
                '            XMLStringConsultarSituacaoLoteRPSResposta = producao.ConsultarSituacaoLoteRps(XMLString1)
                '        End If

                '    End If

                '    xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringConsultarSituacaoLoteRPSResposta)

                '    XMLStreamRet = New MemoryStream(10000)
                '    XMLStreamRet.Write(xRet, 0, xRet.Length)

                '    Dim mySerializerConsultarSituacaoLoteRPSResposta As New XmlSerializer(GetType(ConsultarSituacaoLoteRpsResposta))

                '    Dim objConsultarSituacaoLoteRPSResposta = New ConsultarSituacaoLoteRpsResposta

                '    XMLStreamRet.Position = 0

                '    objConsultarSituacaoLoteRPSResposta = mySerializerConsultarSituacaoLoteRPSResposta.Deserialize(XMLStreamRet)

                '    sAux = objConsultarSituacaoLoteRPSResposta.GetType.ToString()
                '    If InStr(sAux, ".") <> 0 Then
                '        sAux = Mid(sAux, InStr(sAux, ".") + 1)
                '    End If

                '    Select Case sAux

                '        Case "ListaMensagemRetorno"

                '            '                objListaMsgRetorno = objConsultarLoteRPSResposta.Item

                '            '                For iIndice1 = 0 To objListaMsgRetorno.MensagemRetorno.Count - 1

                '            '                    objMsgRetorno = objListaMsgRetorno.MensagemRetorno(iIndice1)

                '            '                    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Correcao, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9})", _
                '            '                    lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objMsgRetorno.Codigo, Left(objMsgRetorno.Mensagem, 255), objMsgRetorno.Correcao, Now.Date, TimeOfDay.ToOADate)

                '            '                    lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                '            '                    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                '            '                    lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Retorno da consulta do lote - " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao, 255), 0)

                '            '                    lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                '            '                    Form1.Msg.Items.Add("Retorno da consulta do lote - " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao)

                '            '                    Application.DoEvents()

                '            '                Next

                '        Case "ConsultarSituacaoLoteRpsResposta"



                '            '' ''Dim objListaMensagemRetorno As ListaMensagemRetorno

                '            '' ''objListaMensagemRetorno = objConsultarSituacaoLoteRPSResposta.Items(0)

                '            ' '' ''                objListaMsgRetornoLote = objConsultarLoteRPSResposta.Item

                '            ' '' ''                For iIndice1 = 0 To objListaMsgRetornoLote.MensagemRetorno.Count - 1

                '            ' '' ''                    objMsgRetornoLote = objListaMsgRetornoLote.MensagemRetorno(iIndice1)

                '            '' ''iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsSitLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Situacao, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9})", _
                '            '' ''lRPSWEBConsSitLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objListaMensagemRetorno.MensagemRetorno(0).Codigo, objListaMensagemRetorno.MensagemRetorno(0).Mensagem, 0, Now.Date, TimeOfDay.ToOADate)

                '            '' ''lRPSWEBConsSitLoteNumIntDoc = lRPSWEBConsSitLoteNumIntDoc + 1

                '            '' ''iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                '            '' ''lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Retorno da consulta de situação do lote - " & objListaMensagemRetorno.MensagemRetorno(0).Codigo & " " & objListaMensagemRetorno.MensagemRetorno(0).Mensagem & " " & objListaMensagemRetorno.MensagemRetorno(0).Correcao, 255), 0)

                '            '' ''lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                '            '' ''Form1.Msg.Items.Add("Retorno da consulta de situação do lote - " & objListaMensagemRetorno.MensagemRetorno(0).Codigo & " " & objListaMensagemRetorno.MensagemRetorno(0).Mensagem & " " & objListaMensagemRetorno.MensagemRetorno(0).Correcao)

                '            '' ''If Form1.Msg.Items.Count - 15 < 1 Then
                '            '' ''    Form1.Msg.TopIndex = 1
                '            '' ''Else
                '            '' ''    Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                '            '' ''End If

                '            '' ''Application.DoEvents()

                '            '                Next



                '            '            Case "ConsultarLoteRpsRespostaListaNfse"

                '            '                objConsultarLoteRpsRespostaListaNFse = objConsultarLoteRPSResposta.Item

                '            '                For iIndice1 = 0 To objConsultarLoteRpsRespostaListaNFse.CompNfse.Count - 1

                '            '                    objCompNfse = objConsultarLoteRpsRespostaListaNFse.CompNfse(iIndice1)

                '            '                    iAchou = 0

                '            '                    For Each objNFiscal In colNFiscal
                '            '                        If Format(CInt(ADM.Serie_Sem_E(objNFiscal.Serie)), "000") = Format(CInt(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie), "000") And _
                '            '                           Format(objNFiscal.NumNotaFiscal, "000000000") = Format(CLng(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero), "000000000") Then
                '            '                            iAchou = 1
                '            '                            Exit For
                '            '                        End If

                '            '                    Next

                '            '                    If iAchou = 0 Then
                '            '                        Throw New System.Exception("A nota consultada nao corresponde a nenhuma das enviadas. Serie Consulta = " & Format(CInt(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie), "000") & " Numero Consulta = " & Format(CLng(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero), "000000000"))
                '            '                    End If

                '            '                    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, tipo, serie, numero, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11})", _
                '            '                    lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, 0, "Nota Fiscal de Servico gerada - " & objCompNfse.Nfse.InfNfse.Numero, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, Now.Date, TimeOfDay.ToOADate)

                '            '                    lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                '            '                    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                '            '                    lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, 0)

                '            '                    lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                '            '                    Form1.Msg.Items.Add("Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero)

                '            '                    Application.DoEvents()

                '            '                    If objCompNfse.Nfse.InfNfse.PrestadorServico.Contato Is Nothing Then
                '            '                        sEmailPrestador = ""
                '            '                        sTelefonePrestador = ""
                '            '                    Else
                '            '                        sEmailPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Contato.Email
                '            '                        If sEmailPrestador = Nothing Then sEmailPrestador = ""
                '            '                        sTelefonePrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Contato.Telefone
                '            '                        If sTelefonePrestador = Nothing Then sTelefonePrestador = ""
                '            '                    End If

                '            '                    If objCompNfse.Nfse.InfNfse.TomadorServico.Contato Is Nothing Then
                '            '                        sEmailTomador = ""
                '            '                        sTelefoneTomador = ""
                '            '                    Else
                '            '                        sEmailTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Contato.Email
                '            '                        If sEmailTomador = Nothing Then sEmailTomador = ""
                '            '                        sTelefoneTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Contato.Telefone
                '            '                        If sTelefoneTomador = Nothing Then sTelefoneTomador = ""
                '            '                    End If

                '            '                    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBProt ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, data, hora, Versao, CodigoVerificacao, Competencia, DataEmissao, DataEmissaoRPS, Id,  TipoRPS, SerieRPS, NumeroRPS, NaturezaOperacao, Numero, OptanteSimplesNacional, OrgaoGeradorCodMun, OrgaoGeradorUF, " & _
                '            '                    "OutrasInformacoes, ContatoEmail, ContatoTelefone, PrestadorBairro, PrestadorCEP, PrestadorCodMun, PrestadorComplemento, PrestadorEndereco, PrestadorNumero, PrestadorUF, PrestadorCNPJ, PrestadorInscMun, PrestadorNomeFantasia, PrestadorRazaoSocial, " & _
                '            '                    "RegimeEspecialTributaco, ServicoCodigoCNAE,ServicoCodMun, ServicoCodTrib, ServicoDiscriminacao, ServicoItemLista, Aliquota, BaseCalculo, DescontoCondicionado, DescontoIncondicionado, ISSRetido, OutrasRetencoes, " & _
                '            '                    "ValorCofins, ValorCsll, ValorDeducoes, ValorINSS, ValorIR, ValorISS, ValorISSRetido, ValorLiquidoNfse, ValorPIS, ValorServicos,  TomadorEmail, TomadorTelefone, TomadorBairro, TomadorCep, TomadorCodMun, " & _
                '            '                    "TomadorComplemento, TomandorEndereco, TomadorNumero, TomadorUF, TomadorCPFCNPJ, TomadorInscMun, TomadorRazaoSocial, ValorCredito, NumIntNF)  VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, " & _
                '            '                    "{20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, {59}, {60}, {61}, {62}, {63}, {64}, {65}, {66}, {67}, {68}, {69}, {70})", _
                '            '                    lRPSWEBProtNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, Now.Date, TimeOfDay.ToOADate, "", objCompNfse.Nfse.InfNfse.CodigoVerificacao, objCompNfse.Nfse.InfNfse.Competencia, objCompNfse.Nfse.InfNfse.DataEmissao, objCompNfse.Nfse.InfNfse.DataEmissaoRps, IIf(objCompNfse.Nfse.InfNfse.Id = Nothing, "", objCompNfse.Nfse.InfNfse.Id), objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, _
                '            '                    objCompNfse.Nfse.InfNfse.NaturezaOperacao, objCompNfse.Nfse.InfNfse.Numero, objCompNfse.Nfse.InfNfse.OptanteSimplesNacional, objCompNfse.Nfse.InfNfse.OrgaoGerador.CodigoMunicipio, objCompNfse.Nfse.InfNfse.OrgaoGerador.Uf, IIf(objCompNfse.Nfse.InfNfse.OutrasInformacoes = Nothing, "", objCompNfse.Nfse.InfNfse.OutrasInformacoes), _
                '            '                    sEmailPrestador, sTelefonePrestador, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Bairro, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Cep, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.CodigoMunicipio, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Complemento, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Endereco, _
                '            '                    objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Numero, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Uf, objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.Cnpj, objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.InscricaoMunicipal, IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia = Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia), objCompNfse.Nfse.InfNfse.PrestadorServico.RazaoSocial, _
                '            '                    objCompNfse.Nfse.InfNfse.RegimeEspecialTributacao, objCompNfse.Nfse.InfNfse.Servico.CodigoCnae, objCompNfse.Nfse.InfNfse.Servico.CodigoMunicipio, objCompNfse.Nfse.InfNfse.Servico.CodigoTributacaoMunicipio, Left(objCompNfse.Nfse.InfNfse.Servico.Discriminacao, 255), _
                '            '                    objCompNfse.Nfse.InfNfse.Servico.ItemListaServico, objCompNfse.Nfse.InfNfse.Servico.Valores.Aliquota, objCompNfse.Nfse.InfNfse.Servico.Valores.BaseCalculo, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoCondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoIncondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.IssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.OutrasRetencoes, _
                '            '                    objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCofins, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCsll, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorDeducoes, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorInss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIr, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorLiquidoNfse, _
                '            '                    objCompNfse.Nfse.InfNfse.Servico.Valores.ValorPis, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorServicos, sEmailTomador, sTelefoneTomador, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Bairro, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Cep, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.CodigoMunicipio, _
                '            '                    IIf(objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento = Nothing, "", objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento), objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Endereco, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Numero, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Uf, objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.CpfCnpj.Item, IIf(objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal = Nothing, "", objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal), _
                '            '                    objCompNfse.Nfse.InfNfse.TomadorServico.RazaoSocial, objCompNfse.Nfse.InfNfse.ValorCredito, objNFiscal.NumIntDoc)

                '            '                    lRPSWEBProtNumIntDoc = lRPSWEBProtNumIntDoc + 1
                '            '                Next

                '    End Select

                '    'If objConsultarSituacaoLoteRPSResposta.Items(1) = 3 Then

                '    '    Dim objConsLoteRpsEnvio As ConsultarLoteRpsEnvio = New ConsultarLoteRpsEnvio


                '    '    objConsLoteRpsEnvio.Prestador = New tcIdentificacaoPrestador

                '    '    objConsLoteRpsEnvio.Prestador.Cnpj = objFilialEmpresa.CGC
                '    '    Call ADM.Formata_String_Numero(objFilialEmpresa.InscricaoMunicipal, objConsLoteRpsEnvio.Prestador.InscricaoMunicipal)

                '    '    objConsLoteRpsEnvio.Protocolo = sProtocolo


                '    '    Dim mySerializerx3 As New XmlSerializer(GetType(ConsultarLoteRpsEnvio))

                '    '    XMLStream1 = New MemoryStream(10000)
                '    '    mySerializerx3.Serialize(XMLStream1, objConsLoteRpsEnvio)

                '    '    Dim xm12 As Byte()
                '    '    xm12 = XMLStream1.ToArray

                '    '    XMLString1 = System.Text.Encoding.UTF8.GetString(xm12)

                '    '    XMLString1 = Mid(XMLString1, 1, 19) & " encoding=""utf-8"" " & Mid(XMLString1, 20)


                '    '    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                '    '    lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "Iniciando a consulta do lote", 0)

                '    '    lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                '    '    Form1.Msg.Items.Add("Iniciando a consulta do lote - Aguarde")

                '    '    Application.DoEvents()

                '    '    Dim XMLStringConsultarLoteRPSResposta As String

                '    '    If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then

                '    '        If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then

                '    '            homologacaobh.ClientCertificates.Add(cert)
                '    '            XMLStringConsultarLoteRPSResposta = homologacaobh.ConsultarLoteRps(XMLStringCabec, XMLString1)
                '    '        Else
                '    '            '                    producao.ClientCredentials.ClientCertificate.Certificate = cert
                '    '            producaobh.ClientCertificates.Add(cert)
                '    '            XMLStringConsultarLoteRPSResposta = producaobh.ConsultarLoteRps(XMLStringCabec, XMLString1)
                '    '        End If


                '    '    Else

                '    '        If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                '    '            homologacao.ClientCredentials.ClientCertificate.Certificate = cert
                '    '            XMLStringConsultarLoteRPSResposta = homologacao.ConsultarLoteRps(XMLString1)
                '    '        Else
                '    '            producao.ClientCredentials.ClientCertificate.Certificate = cert
                '    '            XMLStringConsultarLoteRPSResposta = producao.ConsultarLoteRps(XMLString1)
                '    '        End If

                '    '    End If

                '    '    xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringConsultarLoteRPSResposta)

                '    '    XMLStreamRet = New MemoryStream(10000)
                '    '    XMLStreamRet.Write(xRet, 0, xRet.Length)

                '    '    Dim mySerializerConsultarLoteRPSResposta As New XmlSerializer(GetType(ConsultarLoteRpsResposta))

                '    '    Dim objConsultarLoteRPSResposta = New ConsultarLoteRpsResposta

                '    '    XMLStreamRet.Position = 0

                '    '    objConsultarLoteRPSResposta = mySerializerConsultarLoteRPSResposta.Deserialize(XMLStreamRet)

                '    '    sAux = objConsultarLoteRPSResposta.GetType.ToString()
                '    '    If InStr(sAux, ".") <> 0 Then
                '    '        sAux = Mid(sAux, InStr(sAux, ".") + 1)
                '    '    End If

                '    '    Select Case sAux

                '    '        Case "MensagemListaLoteRpsResposta"

                '    '            objListaMsgRetorno = objConsultarLoteRPSResposta.Item

                '    '            For iIndice1 = 0 To objListaMsgRetorno.MensagemRetorno.Count - 1

                '    '                objMsgRetorno = objListaMsgRetorno.MensagemRetorno(iIndice1)

                '    '                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Correcao, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9})", _
                '    '                lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objMsgRetorno.Codigo, Left(objMsgRetorno.Mensagem, 255), objMsgRetorno.Correcao, Now.Date, TimeOfDay.ToOADate)

                '    '                lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                '    '                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                '    '                lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Left("Retorno da consulta do lote - " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao, 255), 0)

                '    '                lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                '    '                Form1.Msg.Items.Add("Retorno da consulta do lote - " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao)

                '    '                Application.DoEvents()

                '    '            Next

                '    '        Case "ConsultarLoteRpsResposta"


                '    '            objListaMsgRetornoLote = objConsultarLoteRPSResposta.Item

                '    '            For iIndice1 = 0 To objListaMsgRetornoLote.MensagemRetorno.Count - 1

                '    '                objMsgRetornoLote = objListaMsgRetornoLote.MensagemRetorno(iIndice1)

                '    '                If objMsgRetornoLote.Codigo = "E10" Then


                '    '                End If

                '    '                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Serie, Numero, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10})", _
                '    '                lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objMsgRetornoLote.Codigo, objMsgRetornoLote.Mensagem, objMsgRetornoLote.IdentificacaoRps.Serie, objMsgRetornoLote.IdentificacaoRps.Numero, Now.Date, TimeOfDay.ToOADate)

                '    '                lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                '    '                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                '    '                lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "Retorno da consulta do lote - Serie = " & objMsgRetornoLote.IdentificacaoRps.Serie & " RPS = " & objMsgRetornoLote.IdentificacaoRps.Numero & " - " & objMsgRetornoLote.Codigo & " - " & objMsgRetornoLote.Mensagem, 0)

                '    '                lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                '    '                Form1.Msg.Items.Add("Retorno da consulta do lote - Serie = " & objMsgRetornoLote.IdentificacaoRps.Serie & " RPS = " & objMsgRetornoLote.IdentificacaoRps.Numero & " - " & objMsgRetornoLote.Codigo & " - " & objMsgRetornoLote.Mensagem)

                '    '                Application.DoEvents()

                '    '            Next




                '    '            'Case "ConsultarLoteRpsResposta"

                '    '            '    objConsultarLoteRpsRespostaListaNFse = objConsultarLoteRPSResposta.Item

                '    '            '    For iIndice1 = 0 To objConsultarLoteRpsRespostaListaNFse.MensagemRetorno.Count - 1

                '    '            '        objCompNfse = objConsultarLoteRpsRespostaListaNFse.CompNfse(iIndice1)

                '    '            '        iAchou = 0

                '    '            '        For Each objNFiscal In colNFiscal
                '    '            '            If Format(CInt(ADM.Serie_Sem_E(objNFiscal.Serie)), "000") = Format(CInt(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie), "000") And _
                '    '            '               Format(objNFiscal.NumNotaFiscal, "000000000") = Format(CLng(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero), "000000000") Then
                '    '            '                iAchou = 1
                '    '            '                Exit For
                '    '            '            End If

                '    '            '        Next

                '    '            '        If iAchou = 0 Then
                '    '            '            Throw New System.Exception("A nota consultada nao corresponde a nenhuma das enviadas. Serie Consulta = " & Format(CInt(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie), "000") & " Numero Consulta = " & Format(CLng(objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero), "000000000"))
                '    '            '        End If

                '    '            '        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBConsLote ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, tipo, serie, numero, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11})", _
                '    '            '        lRPSWEBConsLoteNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, 0, "Nota Fiscal de Servico gerada - " & objCompNfse.Nfse.InfNfse.Numero, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, Now.Date, TimeOfDay.ToOADate)

                '    '            '        lRPSWEBConsLoteNumIntDoc = lRPSWEBConsLoteNumIntDoc + 1

                '    '            '        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
                '    '            '        lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, 0)

                '    '            '        lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                '    '            '        Form1.Msg.Items.Add("Nota Fiscal de Serviço " & objCompNfse.Nfse.InfNfse.Numero & " gerada. RPS - Tipo = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo & " Série = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie & " Número = " & objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero)

                '    '            '        Application.DoEvents()

                '    '            '        If objCompNfse.Nfse.InfNfse.PrestadorServico.Contato Is Nothing Then
                '    '            '            sEmailPrestador = ""
                '    '            '            sTelefonePrestador = ""
                '    '            '        Else
                '    '            '            sEmailPrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Contato.Email
                '    '            '            If sEmailPrestador = Nothing Then sEmailPrestador = ""
                '    '            '            sTelefonePrestador = objCompNfse.Nfse.InfNfse.PrestadorServico.Contato.Telefone
                '    '            '            If sTelefonePrestador = Nothing Then sTelefonePrestador = ""
                '    '            '        End If

                '    '            '        If objCompNfse.Nfse.InfNfse.TomadorServico.Contato Is Nothing Then
                '    '            '            sEmailTomador = ""
                '    '            '            sTelefoneTomador = ""
                '    '            '        Else
                '    '            '            sEmailTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Contato.Email
                '    '            '            If sEmailTomador = Nothing Then sEmailTomador = ""
                '    '            '            sTelefoneTomador = objCompNfse.Nfse.InfNfse.TomadorServico.Contato.Telefone
                '    '            '            If sTelefoneTomador = Nothing Then sTelefoneTomador = ""
                '    '            '        End If

                '    '            '        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBProt ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, data, hora, Versao, CodigoVerificacao, Competencia, DataEmissao, DataEmissaoRPS, Id,  TipoRPS, SerieRPS, NumeroRPS, NaturezaOperacao, Numero, OptanteSimplesNacional, OrgaoGeradorCodMun, OrgaoGeradorUF, " & _
                '    '            '        "OutrasInformacoes, ContatoEmail, ContatoTelefone, PrestadorBairro, PrestadorCEP, PrestadorCodMun, PrestadorComplemento, PrestadorEndereco, PrestadorNumero, PrestadorUF, PrestadorCNPJ, PrestadorInscMun, PrestadorNomeFantasia, PrestadorRazaoSocial, " & _
                '    '            '        "RegimeEspecialTributaco, ServicoCodigoCNAE,ServicoCodMun, ServicoCodTrib, ServicoDiscriminacao, ServicoItemLista, Aliquota, BaseCalculo, DescontoCondicionado, DescontoIncondicionado, ISSRetido, OutrasRetencoes, " & _
                '    '            '        "ValorCofins, ValorCsll, ValorDeducoes, ValorINSS, ValorIR, ValorISS, ValorISSRetido, ValorLiquidoNfse, ValorPIS, ValorServicos,  TomadorEmail, TomadorTelefone, TomadorBairro, TomadorCep, TomadorCodMun, " & _
                '    '            '        "TomadorComplemento, TomandorEndereco, TomadorNumero, TomadorUF, TomadorCPFCNPJ, TomadorInscMun, TomadorRazaoSocial, ValorCredito, NumIntNF)  VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, " & _
                '    '            '        "{20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, {59}, {60}, {61}, {62}, {63}, {64}, {65}, {66}, {67}, {68}, {69}, {70})", _
                '    '            '        lRPSWEBProtNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, Now.Date, TimeOfDay.ToOADate, "", objCompNfse.Nfse.InfNfse.CodigoVerificacao, objCompNfse.Nfse.InfNfse.Competencia, objCompNfse.Nfse.InfNfse.DataEmissao, objCompNfse.Nfse.InfNfse.DataEmissaoRps, IIf(objCompNfse.Nfse.InfNfse.Id = Nothing, "", objCompNfse.Nfse.InfNfse.Id), objCompNfse.Nfse.InfNfse.IdentificacaoRps.Tipo, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Serie, objCompNfse.Nfse.InfNfse.IdentificacaoRps.Numero, _
                '    '            '        objCompNfse.Nfse.InfNfse.NaturezaOperacao, objCompNfse.Nfse.InfNfse.Numero, objCompNfse.Nfse.InfNfse.OptanteSimplesNacional, objCompNfse.Nfse.InfNfse.OrgaoGerador.CodigoMunicipio, objCompNfse.Nfse.InfNfse.OrgaoGerador.Uf, IIf(objCompNfse.Nfse.InfNfse.OutrasInformacoes = Nothing, "", objCompNfse.Nfse.InfNfse.OutrasInformacoes), _
                '    '            '        sEmailPrestador, sTelefonePrestador, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Bairro, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Cep, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.CodigoMunicipio, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Complemento, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Endereco, _
                '    '            '        objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Numero, objCompNfse.Nfse.InfNfse.PrestadorServico.Endereco.Uf, objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.Cnpj, objCompNfse.Nfse.InfNfse.PrestadorServico.IdentificacaoPrestador.InscricaoMunicipal, IIf(objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia = Nothing, "", objCompNfse.Nfse.InfNfse.PrestadorServico.NomeFantasia), objCompNfse.Nfse.InfNfse.PrestadorServico.RazaoSocial, _
                '    '            '        objCompNfse.Nfse.InfNfse.RegimeEspecialTributacao, objCompNfse.Nfse.InfNfse.Servico.CodigoCnae, objCompNfse.Nfse.InfNfse.Servico.CodigoMunicipio, objCompNfse.Nfse.InfNfse.Servico.CodigoTributacaoMunicipio, Left(objCompNfse.Nfse.InfNfse.Servico.Discriminacao, 255), _
                '    '            '        objCompNfse.Nfse.InfNfse.Servico.ItemListaServico, objCompNfse.Nfse.InfNfse.Servico.Valores.Aliquota, objCompNfse.Nfse.InfNfse.Servico.Valores.BaseCalculo, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoCondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.DescontoIncondicionado, objCompNfse.Nfse.InfNfse.Servico.Valores.IssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.OutrasRetencoes, _
                '    '            '        objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCofins, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorCsll, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorDeducoes, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorInss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIr, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIss, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorIssRetido, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorLiquidoNfse, _
                '    '            '        objCompNfse.Nfse.InfNfse.Servico.Valores.ValorPis, objCompNfse.Nfse.InfNfse.Servico.Valores.ValorServicos, sEmailTomador, sTelefoneTomador, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Bairro, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Cep, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.CodigoMunicipio, _
                '    '            '        IIf(objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento = Nothing, "", objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Complemento), objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Endereco, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Numero, objCompNfse.Nfse.InfNfse.TomadorServico.Endereco.Uf, objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.CpfCnpj.Item, IIf(objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal = Nothing, "", objCompNfse.Nfse.InfNfse.TomadorServico.IdentificacaoTomador.InscricaoMunicipal), _
                '    '            '        objCompNfse.Nfse.InfNfse.TomadorServico.RazaoSocial, objCompNfse.Nfse.InfNfse.ValorCredito, objNFiscal.NumIntDoc)

                '    '            '        lRPSWEBProtNumIntDoc = lRPSWEBProtNumIntDoc + 1
                '    '            '    Next

                '    '    End Select



                '    'End If

                'End If
            End If

            db1.Transaction.Commit()

#End If

            Envia_Lote_RPS = ADM.SUCESSO

        Catch ex As Exception

            If ex.InnerException Is Nothing Then
                sMsg = ""
            Else
                sMsg = " - " & ex.InnerException.Message
            End If

            Form1.Msg.Items.Add("ERRO - " & ex.Message & sMsg & IIf(lNumNotaFiscal <> 0, "Serie = " & sSerie & " Nota Fiscal = " & lNumNotaFiscal, ""))
            Form1.Msg.Items.Add("ERRO - " & sErro & " - " & sMsg1 & IIf(lNumNotaFiscal <> 0, " Serie = " & sSerie & " Nota Fiscal = " & lNumNotaFiscal, ""))
            Form1.Msg.Items.Add("ERRO - o envio do lote " & CStr(lLote) & " foi encerrado por erro.")

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Replace(Left(ex.Message & sMsg, 255), "'", "*"), lNumIntNF)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "ERRO - Encerrado o envio deste lote", 0)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1


            If Form1.Msg.Items.Count - 15 < 1 Then
                Form1.Msg.TopIndex = 1
            Else
                Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
            End If

            Application.DoEvents()

            db1.Transaction.Rollback()

            Envia_Lote_RPS = 1

        Finally

            If lRPSWEBLoteLogNumIntDoc <> 0 Then
                iResult = db2.ExecuteCommand("UPDATE FatConfig Set Conteudo = {0} WHERE Codigo = {1}", lRPSWEBLoteLogNumIntDoc, "NUM_INT_PROX_RPSWEBLOTELOG")
            End If

            If lRPSWEBConsLoteNumIntDoc <> 0 Then
                iResult = db2.ExecuteCommand("UPDATE FatConfig Set Conteudo = {0} WHERE Codigo = {1}", lRPSWEBConsLoteNumIntDoc, "NUM_INT_PROX_RPSWEBCONSLOTE")
            End If

            If lRPSWEBProtNumIntDoc <> 0 Then
                iResult = db2.ExecuteCommand("UPDATE FatConfig Set Conteudo = {0} WHERE Codigo = {1}", lRPSWEBProtNumIntDoc, "NUM_INT_PROX_RPSWEBPROT")
            End If

            If lRPSWEBConsSitLoteNumIntDoc <> 0 Then
                iResult = db2.ExecuteCommand("UPDATE FatConfig Set Conteudo = {0} WHERE Codigo = {1}", lRPSWEBConsSitLoteNumIntDoc, "NUM_INT_PROX_RPSWEBCONSSITLOTE")
            End If

            If lRPSWEBRetEnviNumIntDoc <> 0 Then
                iResult = db2.ExecuteCommand("UPDATE FatConfig Set Conteudo = {0} WHERE Codigo = {1}", lRPSWEBRetEnviNumIntDoc, "NUM_INT_PROX_RPSWEBRETENVI")
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


            db1.Dispose()
            dic.Dispose()
            db2.Dispose()

        End Try



#End If


    End Function

    Sub CalculaDV_Modulo11(ByVal sString As String, ByRef iDigito As Integer)
        Dim iIndice As Integer
        Dim iMult As Integer
        Dim iTotal As Integer

        iMult = 2

        For iIndice = Len(sString) To 1 Step -1

            iTotal = iTotal + (Mid(sString, iIndice, 1) * iMult)

            If iMult = 9 Then
                iMult = 2
            Else
                iMult = iMult + 1
            End If

        Next

        iDigito = iTotal Mod 11

        iDigito = 11 - iDigito

        If iDigito > 9 Then iDigito = 0

    End Sub

    Function PIS_CST(ByRef iCST As Integer, ByVal objTributacaoDocItem As TributacaoDocItem) As Long
        If objTributacaoDocItem.PISCredito > 0 Then
            iCST = 1
        Else
            iCST = 4
        End If
        PIS_CST = ADM.SUCESSO
    End Function

    Function PIS_Aliquota(ByRef dAliquota As Double, ByVal objFilialEmpresa As FiliaisEmpresa) As Long
        If objFilialEmpresa.PISNaoCumulativo = 1 Then
            dAliquota = 0.0165
        Else
            dAliquota = 0.0065
        End If
        PIS_Aliquota = ADM.SUCESSO
    End Function

    Function COFINS_CST(ByRef iCST As Integer, ByVal objTributacaoDocItem As TributacaoDocItem) As Long
        If objTributacaoDocItem.COFINSCredito > 0 Then
            iCST = 1
        Else
            iCST = 4
        End If
        COFINS_CST = ADM.SUCESSO
    End Function

    Function COFINS_Aliquota(ByRef dAliquota As Double, ByVal objFilialEmpresa As FiliaisEmpresa) As Long
        If objFilialEmpresa.COFINSNaoCumulativo = 1 Then
            dAliquota = 0.076
        Else
            dAliquota = 0.03
        End If
        COFINS_Aliquota = ADM.SUCESSO
    End Function


End Class
