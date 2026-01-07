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
Imports System.Configuration
Imports System.Net.Http
Imports System.Globalization
Imports nfseabrasf.NFeXsd


Public Class ClassEnvioRPS

    Public Const RPS_AMBIENTE_HOMOLOGACAO As Integer = 2
    Public Const RPS_AMBIENTE_PRODUCAO As Integer = 1

    Public Const TRIB_TIPO_CALCULO_VALOR = 0
    Public Const TRIB_TIPO_CALCULO_PERCENTUAL = 1

    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Long
    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

    Private Const NS_DPS As String = "http://www.nfse.gov.br/dps"
    Private Const URL_HOMO As String = "http://sefin.producaorestrita.nfse.gov.br/SefinNacional/nfse"

    ' Ajuste conforme sua política (verProc etc.)
    Private Const VER_PROC As String = "SGE-NFSE-NACIONAL"


    Const TIPODOC_TRIB_NF = 0

    Public Function Envia_Lote_RPS(ByVal sEmpresa As String, ByVal lLote As Long, ByVal iFilialEmpresa As Integer) As Long

        Dim url As String = Nothing
        Try
            url = ConfigurationManager.AppSettings("NFSE_NACIONAL_URL")
        Catch
        End Try

        If String.IsNullOrWhiteSpace(url) Then
            Throw New Exception("NFSE_NACIONAL_URL não configurado no app.config (rota NACIONAL acionada).")
        End If


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
        Dim objCidadeFilial As Cidade
        Dim objFilialEmpresa As FiliaisEmpresa
        Dim objEstado As Estado
        Dim objEndereco As Endereco
        Dim objFiliaisClientes As FiliaisCliente
        Dim lEndereco As Long
        Dim lEndDest As Long
        Dim objProduto As Produto
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
        Dim objEndDest As Endereco

        sSerie = ""
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


            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})",
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
            resCidade = db1.ExecuteQuery(Of Cidade) _
                    ("SELECT * FROM Cidades WHERE Descricao = {0}", objEndFilial.Cidade)

            objCidadeFilial = resCidade(0)


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

                    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog (NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})",
                    lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "Iniciando o processamento da Nota Fiscal = " & objNFiscal.NumNotaFiscal & " Série = " & objNFiscal.Serie, lNumIntNF)

                    lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                    Form1.Msg.Items.Add("Iniciando o processamento da Nota Fiscal = " & objNFiscal.NumNotaFiscal & " Série = " & ADM.Serie_Sem_E(objNFiscal.Serie))

                    If Form1.Msg.Items.Count - 15 < 1 Then
                        Form1.Msg.TopIndex = 1
                    Else
                        Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                    End If

                    Application.DoEvents()

                    Dim dps As New NFeXsd.TCDPS() With {
                        .versao = "1.00",
                        .infDPS = New NFeXsd.TCInfDPS()
                    }

                    Dim inf As TCInfDPS = dps.infDPS

                    ' atributo Id do infDPS
                    dps.infDPS.Id = "DPS" + objCidadeFilial.CodIBGE + "2" + objFilialEmpresa.CGC + PadLeftZeros(sSerie, 5) + PadLeftZeros(lNumNotaFiscal.ToString(), 15)

                    ' dCompet: use a competência (mês) - aqui por DataEmissao
                    Dim comp = objNFiscal.DataEmissao
                    ' elementos do TCInfDPS (nomes reais do seu XSD)
                    dps.infDPS.tpAmb = If(iRPSAmbiente = 1, NFeXsd.TSTipoAmbiente.Item1, NFeXsd.TSTipoAmbiente.Item2)
                    dps.infDPS.dhEmi = comp.ToString("yyyy-MM-ddTHH:mm:sszzz")
                    dps.infDPS.verAplic = "1.00"
                    dps.infDPS.dCompet = New DateTime(comp.Year, comp.Month, 1).ToString("yyyy-MM-dd")
                    dps.infDPS.tpEmit = NFeXsd.TSEmitenteDPS.Item1
                    dps.infDPS.cLocEmi = objCidadeFilial.CodIBGE

                    ' =========================
                    ' prest (TCInfoPrestador)
                    ' =========================
                    inf.prest = New TCInfoPrestador()

                    ' -------- choice: CNPJ / CPF / cNaoNIF --------
                    Dim cgc As String = SoNumeros(objFilialEmpresa.CGC)

                    If cgc.Length = 14 Then
                        inf.prest.Item = cgc
                        inf.prest.ItemElementName = ItemChoiceType1.CNPJ

                    ElseIf cgc.Length = 11 Then
                        inf.prest.Item = cgc
                        inf.prest.ItemElementName = ItemChoiceType1.CPF
                    End If

                    ' -------- dados básicos --------
                    inf.prest.IM = SoNumeros(objFilialEmpresa.InscricaoMunicipal)
                    inf.prest.xNome = Trunc(objEmpresa.Nome, 115)

                    ' =========================
                    ' end (TCEndereco)
                    ' =========================
                    inf.prest.end = New TCEndereco()

                    ' --- endNac (choice correto) ---
                    Dim endNac As New TCEnderNac()
                    endNac.cMun = objCidadeFilial.CodIBGE
                    endNac.CEP = PadLeftZeros(SoNumeros(objEndFilial.CEP), 8)

                    inf.prest.end.Item = endNac

                    ' --- demais campos do endereço ---
                    inf.prest.end.xLgr = Trunc(objEndFilial.Logradouro, 255)
                    inf.prest.end.nro = Trunc(objEndFilial.Numero, 60)

                    If Not String.IsNullOrWhiteSpace(objEndFilial.Complemento) Then
                        inf.prest.end.xCpl = Trunc(objEndFilial.Complemento, 156)
                    End If

                    inf.prest.end.xBairro = Trunc(objEndFilial.Bairro, 60)

                    If Not String.IsNullOrWhiteSpace(objEndFilial.Email) Then
                        inf.prest.email = Trunc(objEndFilial.Email, 80)
                    End If

                    inf.prest.regTrib = New TCRegTrib
                    inf.prest.regTrib.opSimpNac = IIf(objFilialEmpresa.SuperSimples = 1, TSOpSimpNac.Item3, TSOpSimpNac.Item1)
                    inf.prest.regTrib.regEspTrib = TSRegEspTrib.Item0

                    sErro = "32"
                    sMsg1 = "vai ler FiliaisClientes"

                    If iDebug = 1 Then MsgBox("32")

                    resFiliaisClientes = db1.ExecuteQuery(Of FiliaisCliente) _
            ("SELECT * FROM FiliaisClientes WHERE CodCliente = {0} AND CodFilial = {1}", objNFiscal.Cliente, objNFiscal.FilialCli)

                    For Each objFiliaisClientes In resFiliaisClientes


                        sErro = "33"
                        sMsg1 = "leu FiliaisClientes"
                        Exit For
                    Next

                    sErro = "34"
                    sMsg1 = "vai ler Clientes"

                    If iDebug = 1 Then MsgBox("34")

                    resCliente = db1.ExecuteQuery(Of Cliente) _
            ("SELECT * FROM Clientes WHERE Codigo = {0}", objNFiscal.Cliente)

                    For Each objCliente In resCliente
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

                        resCidade = db1.ExecuteQuery(Of Cidade) _
                ("SELECT * FROM Cidades WHERE Descricao = {0}", objEndDest.Cidade)

                        For Each objCidade In resCidade
                            Exit For
                        Next

                        Exit For
                    Next

                    ' =========================
                    ' toma (TCInfoPessoa)
                    ' =========================
                    inf.toma = New TCInfoPessoa()

                    ' -------- choice: CNPJ / CPF --------
                    Dim doc As String = SoNumeros(objFiliaisClientes.CGC)

                    If doc.Length = 14 Then
                        inf.toma.Item = doc
                        inf.toma.ItemElementName = ItemChoiceType1.CNPJ

                    ElseIf doc.Length = 11 Then
                        inf.toma.Item = doc
                        inf.toma.ItemElementName = ItemChoiceType1.CPF
                    End If

                    ' -------- nome --------
                    inf.toma.xNome = Trunc(objCliente.RazaoSocial, 150)

                    ' =========================
                    ' end (TCEndereco)
                    ' =========================
                    Dim endNacToma As New TCEnderNac()
                    endNacToma.cMun = objCidade.CodIBGE.ToString()
                    endNacToma.CEP = PadLeftZeros(SoNumeros(objEndDest.CEP), 8)

                    Dim ender As New TCEndereco()
                    ender.Item = endNacToma          ' escolhe endNac (não endExt)
                    ender.xLgr = Trunc(objEndDest.Logradouro, 255)
                    ender.nro = Trunc(objEndDest.Numero, 60)

                    If Not String.IsNullOrWhiteSpace(objEndDest.Complemento) Then
                        ender.xCpl = Trunc(objEndDest.Complemento, 156)
                    End If

                    ender.xBairro = Trunc(objEndDest.Bairro, 60)

                    inf.toma.end = ender

                    ' -------- contatos (opcionais) --------
                    If Not String.IsNullOrWhiteSpace(objEndDest.Telefone1) Then
                        inf.toma.fone = PadLeftZeros(SoNumeros(objEndDest.Telefone1), 11)
                    End If

                    If Not String.IsNullOrWhiteSpace(objEndDest.Email) Then
                        inf.toma.email = Trunc(objEndDest.Email, 80)
                    End If


                    sErro = "20"
                    sMsg1 = "vai consultar a tabela TributacaoDoc"

                    If iDebug = 1 Then MsgBox("20")


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

                    Dim sDiscriminacao As String = ""

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

                            End If
                        End If

                        If Len(sDiscriminacao) = 0 Then
                            sDiscriminacao = objItemNF.DescricaoItem & " Quant: " & Format(objItemNF.Quantidade, "###.###,###") & " P.Unit: " & Format(objItemNF.PrecoUnitario, "fixed") & " Total: " & Format(objItemNF.Quantidade * objItemNF.PrecoUnitario - objItemNF.ValorDesconto, "fixed")
                        Else
                            sDiscriminacao = sDiscriminacao & "|" & objItemNF.DescricaoItem & " Quant: " & Format(objItemNF.Quantidade, "###.###,###") & " P.Unit: " & Format(objItemNF.PrecoUnitario, "fixed") & " Total: " & Format(objItemNF.Quantidade * objItemNF.PrecoUnitario - objItemNF.ValorDesconto, "fixed")
                        End If

                        sErro = "26"
                        sMsg1 = "vai pegar os dados de PIS e COFINS"

                        If iDebug = 1 Then MsgBox("26")

                        dValorPIS = dValorPIS + objTribDocItem.PISValor
                        dValorCOFINS = dValorCOFINS + objTribDocItem.COFINSValor

                    Next

                    If iAchou = 0 Then Throw New System.Exception("Não há nenhum produto nesta nota que tenha natureza serviço.")


                    Dim cServ As TCServ = New TCServ
                    inf.serv = cServ
                    cServ.locPrest = New TCLocPrest
                    cServ.locPrest.ItemElementName = ItemChoiceType3.cLocPrestacao
                    cServ.locPrest.Item = objCidadeFilial.CodIBGE

                    cServ.cServ = New TCCServ

                    Select Case objTribDocItem.ISSCodServ

                        Case Else
                            cServ.cServ.cTribNac = objTribDocItem.ISSCodServ

                    End Select

                    sErro = "27"
                    sMsg1 = "vai transferir os valores"

                    If iDebug = 1 Then MsgBox("27")

                    inf.valores = New TCInfoValores
                    inf.valores.vServPrest = New TCVServPrest
                    If objTributacaoDoc.ISSIncluso = 1 Then
                        inf.valores.vServPrest.vServ = Format(objNFiscal.ValorProdutos, "fixed")
                    Else
                        inf.valores.vServPrest.vServ = Format(objNFiscal.ValorProdutos + CDbl(Format(objTributacaoDoc.ISSValor, "fixed")), "fixed")
                    End If

                    If objNFiscal.ValorDesconto <> 0 Then
                        inf.valores.vDescCondIncond = New TCVDescCondIncond
                        inf.valores.vDescCondIncond.vDescIncond = Format(objNFiscal.ValorDesconto, "fixed")
                    End If

                    inf.valores.trib = New TCInfoTributacao
                    inf.valores.trib.tribMun = New TCTribMunicipal
                    inf.valores.trib.tribMun.tribISSQN = TSTribISSQN.Item1
                    inf.valores.trib.tribMun.tpRetISSQN = TSTipoRetISSQN.Item1

                    If (objTributacaoDoc.PISRetido <> 0 Or objTributacaoDoc.COFINSRetido <> 0 Or objTributacaoDoc.IRRFValor <> 0 Or objTributacaoDoc.CSLLRetido <> 0) Then
                        inf.valores.trib.tribFed = New TCTribFederal

                        If (objTributacaoDoc.PISRetido <> 0 Or objTributacaoDoc.COFINSRetido <> 0) Then

                            inf.valores.trib.tribFed.piscofins = New TCTribOutrosPisCofins
                            inf.valores.trib.tribFed.piscofins.CST = TSTipoCST.Item01 '????

                            inf.valores.trib.tribFed.piscofins.vPis = Format(IIf(objTributacaoDoc.PISRetido <> 0, objTributacaoDoc.PISRetido, 0), "fixed")
                            inf.valores.trib.tribFed.piscofins.vCofins = Format(IIf(objTributacaoDoc.COFINSRetido <> 0, objTributacaoDoc.COFINSRetido, 0), "fixed")

                            inf.valores.trib.tribFed.piscofins.tpRetPisCofins = TSTipoRetPISCofins.Item1
                            inf.valores.trib.tribFed.piscofins.tpRetPisCofinsSpecified = True

                        End If
                        If objTributacaoDoc.INSSRetido <> 0 Then
                            inf.valores.trib.tribFed.vRetCP = Format(objTributacaoDoc.ValorINSS, "fixed")
                        End If
                        If objTributacaoDoc.IRRFValor <> 0 Then
                            inf.valores.trib.tribFed.vRetIRRF = Format(objTributacaoDoc.IRRFValor, "fixed")
                        End If
                        If objTributacaoDoc.CSLLRetido <> 0 Then
                            inf.valores.trib.tribFed.vRetCSLL = Format(objTributacaoDoc.CSLLRetido, "fixed")
                        End If
                    End If

                    objDadosServico.Valores.ValorIss = Format(objTributacaoDoc.ISSValor, "fixed")
                    objDadosServico.Valores.ValorIssSpecified = True

                    objDadosServico.Valores.AliquotaSpecified = True
                    objDadosServico.Valores.Aliquota = objTribDocItem.ISSAliquota
                    If objTributacaoDoc.ISSRetido > 0 Then
                        objDadosServico.IssRetido = 1
                    Else
                        objDadosServico.IssRetido = 2
                    End If




                    sDiscriminacao = sDiscriminacao & "||VALOR TOTAL: R$ " & Format(objDadosServico.Valores.ValorServicos, "Fixed")
                    If objDadosServico.Valores.DescontoIncondicionado <> 0 Then sDiscriminacao = sDiscriminacao & "|DESCONTO: R$ " & Format(objDadosServico.Valores.DescontoIncondicionado, "Fixed")
                    If objDadosServico.Valores.ValorPis <> 0 Then sDiscriminacao = sDiscriminacao & "|PIS: R$ " & Format(objDadosServico.Valores.ValorPis, "Fixed")
                    If objDadosServico.Valores.ValorCofins <> 0 Then sDiscriminacao = sDiscriminacao & "|COFINS: R$ " & Format(objDadosServico.Valores.ValorCofins, "Fixed")
                    If objDadosServico.Valores.ValorCsll <> 0 Then sDiscriminacao = sDiscriminacao & "|CSLL: R$ " & Format(objDadosServico.Valores.ValorCsll, "Fixed")

                    If objDadosServico.Valores.ValorInss <> 0 Then sDiscriminacao = sDiscriminacao & "|INSS: R$ " & Format(objDadosServico.Valores.ValorInss, "Fixed")
                    If objDadosServico.Valores.ValorIr <> 0 Then sDiscriminacao = sDiscriminacao & "|IR: R$ " & Format(objDadosServico.Valores.ValorIr, "Fixed")

                    If objDadosServico.Valores.ValorInss <> 0 Then sDiscriminacao = sDiscriminacao & "|INSS: R$ " & Format(objDadosServico.Valores.ValorInss, "Fixed")
                    If objDadosServico.Valores.ValorIr <> 0 Then sDiscriminacao = sDiscriminacao & "|IR: R$ " & Format(objDadosServico.Valores.ValorIr, "Fixed")
                    If objAdmConfig.Conteudo <> "2" Then

                        If Len(Trim(objNFiscal.MensagemCorpoNota)) <> 0 Then
                            sDiscriminacao = sDiscriminacao & "||" & Trim(objNFiscal.MensagemCorpoNota)
                        End If

                        If Len(Trim(objNFiscal.MensagemNota)) <> 0 Then
                            sDiscriminacao = sDiscriminacao & "||" & Trim(objNFiscal.MensagemNota)
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
                            sDiscriminacao = sDiscriminacao & "||" & ADM.DesacentuaTexto(sMsg)
                        End If

                    End If

                    sDiscriminacao = sDiscriminacao & "||RPS: " & CStr(lNumNotaFiscal) & " Serie: " & sSerie

                    sDiscriminacao = Replace(sDiscriminacao, "|", vbCrLf)

                    cServ.cServ.xDescServ = sDiscriminacao

                    sErro = "28"
                    sMsg1 = "vai ler FiliaisEmpresa"

                    If iDebug = 1 Then MsgBox("28")


                    Dim ns As New XmlSerializerNamespaces()
                    ns.Add("", "http://www.sped.fazenda.gov.br/nfse") ' namespace real do seu XSD
                    ns.Add("ds", "http://www.w3.org/2000/09/xmldsig#")

                    Dim ser As New XmlSerializer(GetType(NFeXsd.TCDPS))

                    Using ms As New MemoryStream()
                        Using tw As New StreamWriter(ms, New UTF8Encoding(False))
                            ser.Serialize(tw, dps, ns)
                        End Using
                        Return Encoding.UTF8.GetString(ms.ToArray())
                    End Using

                    Dim AD As AssinaturaDigital = New AssinaturaDigital

                    sErro = "44"
                    sMsg1 = "vai serializar e assinar a msg"

                    If iDebug = 1 Then MsgBox("44")

                    Dim mySerializer As New XmlSerializer(GetType(tcDeclaracaoPrestacaoServico))
                    XMLStream = New MemoryStream(10000)

                    mySerializer.Serialize(XMLStream, objRps)

                    Dim xm1 As Byte()
                    xm1 = XMLStream.ToArray

                    XMLString = System.Text.Encoding.UTF8.GetString(xm1)


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


                    iResult = AD.Assinar(XMLString, "InfDeclaracaoPrestacaoServico", cert, objEndFilial.Cidade)

                    If iResult <> 0 Then Throw New System.Exception("Ocorreu um erro durante a assinatura do lote. " & AD.mensagemResultado)

                    sErro = "45"
                    sMsg1 = "vai pegar a msg assinada"

                    If iDebug = 1 Then MsgBox("45")

                    Dim xMlAD As XmlDocument

                    xMlAD = AD.XMLDocAssinado()

                    Dim xString10 As String
                    xString10 = AD.XMLStringAssinado


                    xString10 = Replace(xString10, "<tcDeclaracaoPrestacaoServico", "<Rps")

                    xString10 = Replace(xString10, "</tcDeclaracaoPrestacaoServico", "</Rps")

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
                If Len(objFilialEmpresa.CGC) = 14 Then
                    a5.LoteRps.CpfCnpj = New tcCpfCnpj
                    a5.LoteRps.CpfCnpj.ItemElementName = ItemChoiceType.Cnpj
                    a5.LoteRps.CpfCnpj.Item = objFilialEmpresa.CGC
                ElseIf Len(objFilialEmpresa) = 11 Then
                    a5.LoteRps.CpfCnpj = New tcCpfCnpj
                    a5.LoteRps.CpfCnpj.ItemElementName = ItemChoiceType.Cpf
                    a5.LoteRps.CpfCnpj.Item = objFilialEmpresa.CGC
                End If


                Call ADM.Formata_String_Numero(objFilialEmpresa.InscricaoMunicipal, a5.LoteRps.InscricaoMunicipal)

                sErro = "48"
                sMsg1 = "vai ler a tabela RPSWEBLote"

                If iDebug = 1 Then MsgBox("48")


                resRPSWEBLote = db2.ExecuteQuery(Of RPSWEBLote) _
                ("SELECT * FROM RPSWEBLote WHERE Lote = {0} ORDER BY NumIntNF", lLote)

                a5.LoteRps.QuantidadeRps = resRPSWEBLote.Count


                sErro = "49"
                sMsg1 = "vai serializar EnviarLoteRpsEnvio e assinar o lote"

                If iDebug = 1 Then MsgBox("49")

                Dim objListaRps As tcLoteRpsListaRps
                objListaRps = New tcLoteRpsListaRps

                a5.LoteRps.ListaRps = objListaRps


                Dim mySerializer1 As New XmlSerializer(GetType(EnviarLoteRpsEnvio))

                XMLStream = New MemoryStream(10000)

                mySerializer1.Serialize(XMLStream, a5)

                Dim xm As Byte()
                xm = XMLStream.ToArray

                XMLString = System.Text.Encoding.UTF8.GetString(xm)

                XMLStringRpses = "<ListaRps>" & XMLStringRpses & "</ListaRps>"
                XMLString = Replace(XMLString, "<ListaRps />", XMLStringRpses)

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


                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})",
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
                If iRPSAmbiente = RPS_AMBIENTE_HOMOLOGACAO Then
                    homologacao.ClientCredentials.ClientCertificate.Certificate = cert
                    XMLStringRetEnvRPS = homologacao.RecepcionarLoteRps(xString)
                Else
                    producao.ClientCredentials.ClientCertificate.Certificate = cert
                    XMLStringRetEnvRPS = producao.RecepcionarLoteRps(xString)
                End If


                sErro = "55"
                sMsg1 = "vai deserializar a resposta"


                If iDebug = 1 Then MsgBox("55")

                Dim xRet1 As Byte()



                xRet1 = System.Text.Encoding.UTF8.GetBytes(XMLStringRetEnvRPS)

                XMLStreamRet = New MemoryStream(10000)

                XMLStreamRet.Write(xRet1, 0, xRet1.Length)



                Dim mySerializerRetEnvNFe As New XmlSerializer(GetType(EnviarLoteRpsSincronoResposta))
                Dim objRetEnviRPS As EnviarLoteRpsSincronoResposta = New EnviarLoteRpsSincronoResposta


                XMLStreamRet.Position = 0

                objRetEnviRPS = mySerializerRetEnvNFe.Deserialize(XMLStreamRet)

                sErro = "56"
                sMsg1 = "vai tratar os varios tipos de resposta"

                If iDebug = 1 Then MsgBox("56")


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


                    If Not objListaMsgRetorno.MensagemRetorno Is Nothing Then
                        For iIndice = 0 To objListaMsgRetorno.MensagemRetorno.Count - 1
                            objMsgRetorno = objListaMsgRetorno.MensagemRetorno(iIndice)
                            sLote = lLote
                            sProtocolo = ""
                            objMsgRetorno.Correcao = ""

                            sErro = "57"
                            sMsg1 = "vai gravar RPSWEBRetEnvi"

                            If iDebug = 1 Then MsgBox("57")


                            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBRetEnvi ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Correcao, data, hora, datarecebimento, horarecebimento) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11} )",
    lRPSWEBRetEnviNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objMsgRetorno.Codigo, Left(objMsgRetorno.Mensagem, 200), IIf(objMsgRetorno.Correcao Is Nothing, "", objMsgRetorno.Correcao), Now.Date, TimeOfDay.ToOADate, dtDataRecebimento, dHoraRecebimento)

                            lRPSWEBRetEnviNumIntDoc = lRPSWEBRetEnviNumIntDoc + 1

                            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})",
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

                End If

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


                            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBRetEnvi ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Correcao, data, hora, datarecebimento, horarecebimento) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11} )",
                            lRPSWEBRetEnviNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objMsgRetornoLote.Codigo, Left(objMsgRetornoLote.Mensagem, 200), IIf(objMsgRetornoLote.IdentificacaoRps Is Nothing, "", "RPS - Numero: " & objMsgRetornoLote.IdentificacaoRps.Numero & " Serie: " & objMsgRetornoLote.IdentificacaoRps.Serie & " Tipo: " & objMsgRetornoLote.IdentificacaoRps.Tipo), Now.Date, TimeOfDay.ToOADate, dtDataRecebimento, dHoraRecebimento)

                            lRPSWEBRetEnviNumIntDoc = lRPSWEBRetEnviNumIntDoc + 1

                            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})",
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

                'vai iniciar a consulta a situacao do lote 
                If sProtocolo <> "" Then

                    objMsgRetorno.Codigo = ""
                    objMsgRetorno.Mensagem = "Lote enviado"
                    objMsgRetorno.Correcao = ""

                    sErro = "60"
                    sMsg1 = "vai gravar RPSWEBRetEnvi"


                    If iDebug = 1 Then MsgBox("60")


                    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBRetEnvi ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, CodMsg, Msg, Correcao, data, hora, datarecebimento, horarecebimento) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11} )",
                    lRPSWEBRetEnviNumIntDoc, iFilialEmpresa, iRPSAmbiente, sLote, sProtocolo, objMsgRetorno.Codigo, Left(objMsgRetorno.Mensagem, 200), objMsgRetorno.Correcao, Now.Date, TimeOfDay.ToOADate, dtDataRecebimento, dHoraRecebimento)

                    lRPSWEBRetEnviNumIntDoc = lRPSWEBRetEnviNumIntDoc + 1

                    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumINtDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})",
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

                End If

            Next
            db1.Transaction.Commit()

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

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4},{5}, {6}, {7})",
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Replace(Left(ex.Message & sMsg, 255), "'", "*"), lNumIntNF)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})",
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
