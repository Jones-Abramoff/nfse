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
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization.Json


Public Class ClassEnvioRPS

    Public Const RPS_AMBIENTE_HOMOLOGACAO As Integer = 2
    Public Const RPS_AMBIENTE_PRODUCAO As Integer = 1

    Public Const TRIB_TIPO_CALCULO_VALOR = 0
    Public Const TRIB_TIPO_CALCULO_PERCENTUAL = 1

    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

    Private Const NS_DPS As String = "http://www.nfse.gov.br/dps"
    Private Const URL_HOMO As String = "http://sefin.producaorestrita.nfse.gov.br/SefinNacional/nfse"

    ' Ajuste conforme sua política (verProc etc.)
    Private Const VER_PROC As String = "SGE-NFSE-NACIONAL"


    Const TIPODOC_TRIB_NF = 0

    Public Function Envia_Lote_RPS(ByVal sEmpresa As String, ByVal lLote As Long, ByVal iFilialEmpresa As Integer) As Long

        Dim XMLAssinado As String
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
        Dim objNFiscal As NFeNFiscal
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

            iDebug = 0

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
                    dps.infDPS.dhEmi = DateTime.SpecifyKind(comp, DateTimeKind.Local).ToString("yyyy-MM-ddTHH:mm:sszzz")
                    dps.infDPS.serie = sSerie
                    dps.infDPS.nDPS = lNumNotaFiscal.ToString()
                    dps.infDPS.verAplic = "1.00"
                    dps.infDPS.dCompet = New DateTime(comp.Value.Year, comp.Value.Month, 1).ToString("yyyy-MM-dd")
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
                    '??? Rio nao aceita inf.prest.IM = SoNumeros(objFilialEmpresa.InscricaoMunicipal)
                    'inf.prest.xNome = Trunc(objEmpresa.Nome, 115)

                    ' =========================
                    ' end (TCEndereco)
                    ' =========================
                    'inf.prest.end = New TCEndereco()

                    '' --- endNac (choice correto) ---
                    'Dim endNac As New TCEnderNac()
                    'endNac.cMun = objCidadeFilial.CodIBGE
                    'endNac.CEP = PadLeftZeros(SoNumeros(objEndFilial.CEP), 8)
                    'inf.prest.end.Item = endNac

                    ' --- demais campos do endereço ---
                    'inf.prest.end.xLgr = Trunc(objEndFilial.Logradouro, 255)
                    'inf.prest.end.nro = Trunc(objEndFilial.Numero, 60)

                    'If Not String.IsNullOrWhiteSpace(objEndFilial.Complemento) Then
                    '    inf.prest.end.xCpl = Trunc(objEndFilial.Complemento, 156)
                    'End If

                    'inf.prest.end.xBairro = Trunc(objEndFilial.Bairro, 60)
                    inf.prest.fone = SoNumeros(IIf(Len(Trim(objEndFilial.TelDDD1)) > 1, objEndFilial.TelDDD1.ToString(), "21") + objEndFilial.TelNumero1)
                    If Not String.IsNullOrWhiteSpace(objEndFilial.Email) Then
                        inf.prest.email = Trunc(objEndFilial.Email, 80)
                    End If

                    inf.prest.regTrib = New TCRegTrib
                    If objFilialEmpresa.SuperSimples = 1 Then
                        inf.prest.regTrib.opSimpNac = TSOpSimpNac.Item3
                        inf.prest.regTrib.regApTribSN = TSRegimeApuracaoSimpNac.Item1
                        inf.prest.regTrib.regApTribSNSpecified = True
                    Else
                        inf.prest.regTrib.opSimpNac = TSOpSimpNac.Item1
                    End If

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
            ("SELECT * FROM Enderecos WHERE Codigo = {0}", objFiliaisClientes.Endereco)

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
                    inf.toma.xNome = ADM.DesacentuaTexto(Trim(Trunc(objCliente.RazaoSocial, 150)))

                    ' =========================
                    ' end (TCEndereco)
                    ' =========================
                    Dim endNacToma As New TCEnderNac()
                    endNacToma.cMun = objCidade.CodIBGE.ToString()
                    endNacToma.CEP = PadLeftZeros(SoNumeros(objEndDest.CEP), 8)

                    Dim ender As New TCEndereco()
                    ender.Item = endNacToma          ' escolhe endNac (não endExt)
                    ender.xLgr = Trunc(ADM.DesacentuaTexto(objEndDest.Logradouro), 255)
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
                            sDiscriminacao = objItemNF.DescricaoItem & " Quant: " & Format(objItemNF.Quantidade, "###.###,###") & " P.Unit: " & FormatValor(objItemNF.PrecoUnitario) & " Total: " & FormatValor(objItemNF.Quantidade * objItemNF.PrecoUnitario - objItemNF.ValorDesconto)
                        Else
                            sDiscriminacao = sDiscriminacao & "|" & objItemNF.DescricaoItem & " Quant: " & Format(objItemNF.Quantidade, "###.###,###") & " P.Unit: " & FormatValor(objItemNF.PrecoUnitario) & " Total: " & FormatValor(objItemNF.Quantidade * objItemNF.PrecoUnitario - objItemNF.ValorDesconto)
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

                        'Case "010101"
                        '    cServ.cServ.cTribNac = "010101"
                        '    cServ.cServ.cTribMun = "001"
                        'Case "010501"
                        '    cServ.cServ.cTribNac = "010501"
                        '    cServ.cServ.cTribMun = "001"
                        Case "010701"
                            cServ.cServ.cTribNac = "010701"
                            cServ.cServ.cTribMun = "003"
                        Case "070298"
                            cServ.cServ.cTribNac = "070202"
                            cServ.cServ.cTribMun = "098"
                        Case "140613"
                            cServ.cServ.cTribNac = "140601"
                            cServ.cServ.cTribMun = "013"
                        'Case "310101"
                        '    cServ.cServ.cTribNac = "310101"
                        '    cServ.cServ.cTribMun = "001"
                        Case "310105"
                            cServ.cServ.cTribNac = "310104"
                            cServ.cServ.cTribMun = "001"
                        Case Else
                            cServ.cServ.cTribNac = objTribDocItem.ISSCodServ
                            cServ.cServ.cTribMun = "001"
                    End Select

                    sErro = "27"
                    sMsg1 = "vai transferir os valores"

                    If iDebug = 1 Then MsgBox("27")

                    inf.valores = New TCInfoValores
                    inf.valores.vServPrest = New TCVServPrest
                    If objTributacaoDoc.ISSIncluso = 1 Then
                        inf.valores.vServPrest.vServ = Replace(Format(objNFiscal.ValorProdutos, "fixed"), ",", ".")
                    Else
                        inf.valores.vServPrest.vServ = Replace(Format(objNFiscal.ValorProdutos + CDbl(Format(objTributacaoDoc.ISSValor, "fixed")), "fixed"), ",", ".")
                    End If

                    If objNFiscal.ValorDesconto <> 0 Then
                        inf.valores.vDescCondIncond = New TCVDescCondIncond
                        inf.valores.vDescCondIncond.vDescIncond = Replace(Format(objNFiscal.ValorDesconto, "fixed"), ",", ".")
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

                            inf.valores.trib.tribFed.piscofins.vPis = Replace(Format(IIf(objTributacaoDoc.PISRetido <> 0, objTributacaoDoc.PISRetido, 0), "fixed"), ",", ".")
                            inf.valores.trib.tribFed.piscofins.vCofins = Replace(Format(IIf(objTributacaoDoc.COFINSRetido <> 0, objTributacaoDoc.COFINSRetido, 0), "fixed"), ",", ".")

                            inf.valores.trib.tribFed.piscofins.tpRetPisCofins = TSTipoRetPISCofins.Item1
                            inf.valores.trib.tribFed.piscofins.tpRetPisCofinsSpecified = True

                        End If
                        If objTributacaoDoc.INSSRetido <> 0 Then
                            inf.valores.trib.tribFed.vRetCP = Replace(Format(objTributacaoDoc.ValorINSS, "fixed"), ",", ".")
                        End If
                        If objTributacaoDoc.IRRFValor <> 0 Then
                            inf.valores.trib.tribFed.vRetIRRF = Replace(Format(objTributacaoDoc.IRRFValor, "fixed"), ",", ".")
                        End If
                        If objTributacaoDoc.CSLLRetido <> 0 Then
                            inf.valores.trib.tribFed.vRetCSLL = Replace(Format(objTributacaoDoc.CSLLRetido, "fixed"), ",", ".")
                        End If
                    End If
                    inf.valores.trib.totTrib = New TCTribTotal
                    inf.valores.trib.totTrib.Item = "5"

                    sDiscriminacao = sDiscriminacao & "|VALOR TOTAL: "
                    If objTributacaoDoc.ISSIncluso = 1 Then
                        sDiscriminacao = sDiscriminacao & FormatValor(objNFiscal.ValorProdutos)
                    Else
                        sDiscriminacao = sDiscriminacao & FormatValor(objNFiscal.ValorProdutos + CDbl(Format(objTributacaoDoc.ISSValor, "fixed")))
                    End If
                    If objNFiscal.ValorDesconto <> 0 Then sDiscriminacao = sDiscriminacao & "|DESCONTO: " & FormatValor(objNFiscal.ValorDesconto)
                    If objTributacaoDoc.PISRetido <> 0 Then sDiscriminacao = sDiscriminacao & "|PIS: " & FormatValor(objTributacaoDoc.PISRetido)
                    If objTributacaoDoc.COFINSRetido <> 0 <> 0 Then sDiscriminacao = sDiscriminacao & "|COFINS: " & FormatValor(objTributacaoDoc.COFINSRetido)
                    If objTributacaoDoc.CSLLRetido <> 0 <> 0 Then sDiscriminacao = sDiscriminacao & "|CSLL: " & FormatValor(objTributacaoDoc.CSLLRetido)

                    If objTributacaoDoc.INSSRetido <> 0 <> 0 Then sDiscriminacao = sDiscriminacao & "|INSS: " & FormatValor(objTributacaoDoc.INSSRetido)
                    If objTributacaoDoc.IRRFValor <> 0 Then sDiscriminacao = sDiscriminacao & "|IR: " & FormatValor(objTributacaoDoc.IRRFValor)
                    If objAdmConfig.Conteudo <> "2" Then

                        If Len(Trim(objNFiscal.MensagemCorpoNota)) <> 0 Then
                            sDiscriminacao = sDiscriminacao & "|" & Trim(objNFiscal.MensagemCorpoNota)
                        End If

                        If Len(Trim(objNFiscal.MensagemNota)) <> 0 Then
                            sDiscriminacao = sDiscriminacao & "|" & Trim(objNFiscal.MensagemNota)
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
                            sDiscriminacao = sDiscriminacao & "|" & ADM.DesacentuaTexto(sMsg)
                        End If

                    End If

                    sDiscriminacao = sDiscriminacao & "|RPS: " & CStr(lNumNotaFiscal) & " Serie: " & sSerie

                    sDiscriminacao = Replace(sDiscriminacao, "|", " | ")

                    cServ.cServ.xDescServ = sDiscriminacao

                    sErro = "28"
                    sMsg1 = "vai ler FiliaisEmpresa"

                    If iDebug = 1 Then MsgBox("28")


                    Dim ns As New XmlSerializerNamespaces()
                    ns.Add("", "http://www.sped.fazenda.gov.br/nfse") ' namespace real do seu XSD
                    ns.Add("ds", "http://www.w3.org/2000/09/xmldsig#")

                    Dim ser As New XmlSerializer(GetType(NFeXsd.TCDPS))
                    Dim xm1 As Byte()

                    Using ms As New MemoryStream()
                        Using tw As New StreamWriter(ms, New UTF8Encoding(False))
                            ser.Serialize(tw, dps, ns)
                        End Using
                        XMLString = Encoding.UTF8.GetString(ms.ToArray())
                    End Using

                    Dim AD As AssinaturaDigital = New AssinaturaDigital

                    sErro = "44"
                    sMsg1 = "vai serializar e assinar a msg"

                    If iDebug = 1 Then MsgBox("44")

                    XMLAssinado = XmlUtil.AssinarDps(XMLString, cert)

                    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 1

                    sErro = "47"
                    sMsg1 = "gravou o xml"

                    If iDebug = 1 Then MsgBox("47")

                    Dim json As String = "{""dpsXmlGZipB64"": """ + XmlUtil.GZipToBase64(XMLAssinado) + """}"

                    Dim ret As (StatusCode As HttpStatusCode, Body As String)

                    ret = PostJsonWithCert_WebRequest(IIf(objFilialEmpresa.RPSAmbiente = 1, "http://sefin.nfse.gov.br/SefinNacional/nfse", "http://sefin.producaorestrita.nfse.gov.br/SefinNacional/nfse"), json, cert)

                    If ret.StatusCode <> HttpStatusCode.Created Then
                        'Continue For '?????
                        MsgBox(ret.Body)
                        Throw New Exception("error while sending the dps")
                    End If
                    'ret.Body = "{""tipoAmbiente"":2,""versaoAplicativo"":""SefinNacional_1.5.0"",""dataHoraProcessamento"":""2026-01-07T11:58:59.0403378-03:00"",""idDps"":""NFS33045572207851803000107000000000000126012138651922"",""chaveAcesso"":""33045572207851803000107000000000000126012138651922"",""nfseXmlGZipB64"":""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"",""alertas"":null}"

                    Dim retorno = DesserializarRetorno(ret.Body)
                    Dim dataProc As DateTimeOffset

                    If DateTimeOffset.TryParse(retorno.DataHoraProcessamento,
                           Globalization.CultureInfo.InvariantCulture,
                           Globalization.DateTimeStyles.AssumeUniversal,
                           dataProc) Then
                        ' OK
                    End If
                    If retorno IsNot Nothing Then
                        Console.WriteLine(retorno.ChaveAcesso)

                        If retorno.Alertas IsNot Nothing Then
                            For Each a In retorno.Alertas
                                Console.WriteLine($"{a.Codigo} - {a.Descricao}")
                            Next
                        End If

                        Dim xmlRet = XmlUtil.Base64GunzipToString(retorno.NfseXmlGZipB64)
                        Dim nfseObj As TCNFSe = DesserializarNFSe(xmlRet)

                        Dim xDados100 As Byte()
                        Dim XMLStreamDados100 As MemoryStream = New MemoryStream(10000)

                        xDados100 = System.Text.Encoding.UTF8.GetBytes(xmlRet)

                        XMLStreamDados100.Write(xDados100, 0, xDados100.Length)

                        Dim DocDados100 As XmlDocument = New XmlDocument

                        XMLStreamDados100.Position = 0
                        DocDados100.Load(XMLStreamDados100)
                        '??? sDir = "c:\sge\log\" '????
                        sArquivo = sDir & nfseObj.infNFSe.Id & ".xml"

                        Dim writer100 As New XmlTextWriter(sArquivo, Nothing)

                        writer100.Formatting = Formatting.None
                        DocDados100.WriteTo(writer100)
                        writer100.Close()

                        sErro = "176"
                        sMsg1 = "vai gravar RPSWEBProt"
                        If iDebug = 1 Then MsgBox(sErro & " " & sMsg1)

                        Dim dtDataEmissao As DateTimeOffset

                        If DateTimeOffset.TryParse(nfseObj.infNFSe.DPS.infDPS.dhEmi,
                           Globalization.CultureInfo.InvariantCulture,
                           Globalization.DateTimeStyles.AssumeUniversal,
                           dtDataEmissao) Then
                            ' OK
                        End If
                        Form1.Msg.Items.Add("NFSE autorizada com a chave " + nfseObj.infNFSe.Id)

                        If Form1.Msg.Items.Count - 15 < 1 Then
                            Form1.Msg.TopIndex = 1
                        Else
                            Form1.Msg.TopIndex = Form1.Msg.Items.Count - 15
                        End If

                        Application.DoEvents()

                        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBProt ( NumIntDoc, FilialEmpresa, Ambiente, Lote, Protocolo, data, hora, Versao, CodigoVerificacao, Competencia, DataEmissao, DataEmissaoRPS, Id,  TipoRPS, SerieRPS, NumeroRPS, NaturezaOperacao, Numero, OptanteSimplesNacional, OrgaoGeradorCodMun, OrgaoGeradorUF, " &
                            "OutrasInformacoes, ContatoEmail, ContatoTelefone, PrestadorBairro, PrestadorCEP, PrestadorCodMun, PrestadorComplemento, PrestadorEndereco, PrestadorNumero, PrestadorUF, PrestadorCNPJ, PrestadorInscMun, PrestadorNomeFantasia, PrestadorRazaoSocial, " &
                            "RegimeEspecialTributaco, ServicoCodigoCNAE,ServicoCodMun, ServicoCodTrib, ServicoDiscriminacao, ServicoItemLista, Aliquota, BaseCalculo, DescontoCondicionado, DescontoIncondicionado, ISSRetido, OutrasRetencoes, " &
                            "ValorCofins, ValorCsll, ValorDeducoes, ValorINSS, ValorIR, ValorISS, ValorISSRetido, ValorLiquidoNfse, ValorPIS, ValorServicos,  TomadorEmail, TomadorTelefone, TomadorBairro, TomadorCep, TomadorCodMun, " &
                            "TomadorComplemento, TomandorEndereco, TomadorNumero, TomadorUF, TomadorCPFCNPJ, TomadorInscMun, TomadorRazaoSocial, ValorCredito, NumIntNF)  VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, " &
                            "{20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, {59}, {60}, {61}, {62}, {63}, {64}, {65}, {66}, {67}, {68}, {69}, {70})",
                            lRPSWEBProtNumIntDoc, iFilialEmpresa, nfseObj.infNFSe.ambGer.GetXmlEnumValue(), Left(lLote.ToString(), 15), "", Now.Date, TimeOfDay.ToOADate, "", Left(nfseObj.infNFSe.nDFSe, 9), Left(nfseObj.infNFSe.DPS.infDPS.dCompet, 20), dtDataEmissao, dtDataEmissao, IIf(nfseObj.infNFSe.Id = Nothing, "", Left(nfseObj.infNFSe.Id, 255)), "1", Left(nfseObj.infNFSe.DPS.infDPS.serie, 5), Left(nfseObj.infNFSe.DPS.infDPS.nDPS, 15),
                            "", Left(nfseObj.infNFSe.nNFSe, 15), nfseObj.infNFSe.DPS.infDPS.prest.regTrib.opSimpNac.GetXmlEnumValue(), nfseObj.infNFSe.DPS.infDPS.cLocEmi, "", "",
                            "", "", "", 0, 0, "", "",
                            "", "", "", "", "", "",
                            "", 0, 0, "", Left(nfseObj.infNFSe.DPS.infDPS.serv.cServ.xDescServ, 2000),
                            "", 0, 0, 0, 0, 0, 0,
                            0, 0, 0, 0, 0, 0, 0, 0,
                            0, 0, "", "", "", 0, 0,
                            "", "", "", "", "", "",
                            "", 0, objNFiscal.NumIntDoc)

                        lRPSWEBProtNumIntDoc = lRPSWEBProtNumIntDoc + 1

                        db2.ExecuteCommand("UPDATE NFiscal SET ChvNFe = {0} WHERE NumIntDoc = {1}", nfseObj.infNFSe.Id.Substring(7), objNFiscal.NumIntDoc)

                    End If

                Next
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

    Private Shared Function PostJsonWithCert_WebRequest(
    url As String,
    json As String,
    cert As X509Certificate2
) As (StatusCode As HttpStatusCode, Body As String)

        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

        Dim req = CType(WebRequest.Create(url), HttpWebRequest)
        req.Method = "POST"
        req.ContentType = "application/json; charset=utf-8"
        req.Accept = "application/json"
        req.UserAgent = "NFSeClient/1.0"
        req.Timeout = 120000
        req.ReadWriteTimeout = 120000

        ' Certificado cliente (mTLS)
        req.ClientCertificates.Add(cert)

        ' Envia o body
        Dim bytes = Encoding.UTF8.GetBytes(json)
        req.ContentLength = bytes.Length
        Using rs = req.GetRequestStream()
            rs.Write(bytes, 0, bytes.Length)
        End Using

        ' Lê resposta
        Try
            Using resp = CType(req.GetResponse(), HttpWebResponse)
                Using sr As New StreamReader(resp.GetResponseStream(), Encoding.UTF8)
                    Return (resp.StatusCode, sr.ReadToEnd())
                End Using
            End Using

        Catch ex As WebException
            Dim resp = TryCast(ex.Response, HttpWebResponse)
            Dim body As String = ""
            If resp IsNot Nothing AndAlso resp.GetResponseStream() IsNot Nothing Then
                Using sr As New StreamReader(resp.GetResponseStream(), Encoding.UTF8)
                    body = sr.ReadToEnd()
                End Using
                Return (resp.StatusCode, body)
            End If
            Throw
        End Try
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
    Public Function DesserializarRetorno(json As String) As RetornoCriacaoNfse
        Dim serializer As New DataContractJsonSerializer(GetType(RetornoCriacaoNfse))

        Using ms As New MemoryStream(Encoding.UTF8.GetBytes(json))
            Return CType(serializer.ReadObject(ms), RetornoCriacaoNfse)
        End Using
    End Function

    Public Function DesserializarNFSe(xml As String) As TCNFSe
        Dim serializer As New XmlSerializer(
        GetType(TCNFSe),
        New XmlRootAttribute With {
            .ElementName = "NFSe",
            .Namespace = "http://www.sped.fazenda.gov.br/nfse",
            .IsNullable = False
        }
    )

        Using sr As New StringReader(xml)
            Return CType(serializer.Deserialize(sr), TCNFSe)
        End Using
    End Function


End Class
