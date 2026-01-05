Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.Schema
Imports System.Data.Odbc


Public Class ClassCancelaNFSE

    Public Const NFE_AMBIENTE_HOMOLOGACAO As Integer = 2
    Public Const NFE_AMBIENTE_PRODUCAO As Integer = 1

    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Long

    Function Cancela_NFSE(ByVal sEmpresa As String, ByVal lNumIntNF As Long, ByVal sMotivo As String, ByVal iFilialEmpresa As Integer) As Long


#If Not GINFES Then

#If Not RJ Then
        Dim a4 As cabecalho = New cabecalho
#End If

        Dim objCanc As CancelarNfseEnvio = New CancelarNfseEnvio

        Dim sArquivo As String


        'Jones        Dim resNFeFedProtNFE As IEnumerable(Of NFeFedProtNFe)


        Dim XMLStream As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamDados As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamCabec As MemoryStream = New MemoryStream(10000)

        'Jones        Dim objNFeFedProtNFE As NFeFedProtNFe


        Dim db1 As SGEDadosDataContext = New SGEDadosDataContext
        Dim db2 As SGEDadosDataContext = New SGEDadosDataContext
        Dim dic As SGEDicDataContext = New SGEDicDataContext
        Dim AD As AssinaturaDigital = New AssinaturaDigital
        Dim XMLString As String
        Dim XMLString1 As String
        Dim XMLStringRetCancNFSE As String
        Dim xString As String
        Dim XMLStringCabec As String

        Dim cert As X509Certificate2 = New X509Certificate2
        Dim certificado As Certificado = New Certificado
        Dim sCertificado As String

        Dim homologacao As New NotaCariocaHomologacao.NfseSoapClient
        Dim producao As New NotaCariocaProducao.NfseSoapClient

        Dim homologacaobh = New GetWebRequest_bhHomologacao
        Dim producaobh = New GetWebRequest_bhProducao

        Dim homologacaoginfes = New br.com.ginfes.homologacao.ServiceGinfesImplService
        Dim producaoginfes = New br.com.ginfes.producao.ServiceGinfesImplService

        Dim homsistema4rCancela = New br.com.sistemas4r.abrasf.nfsehml4.CancelarNfse

        Dim prodsistema4rCancela As New br.com.sistemas4r.abrasf.nfse4.CancelarNfse

        Dim iResult As Integer

        Dim odbc As OdbcConnection = New OdbcConnection
        Dim sString As String

        Dim iInicio As Integer
        Dim iFim As Integer

        Dim iTamanho As Integer
        Dim sRetorno As String
        Dim iPos As Integer
        Dim sDir As String
        Dim sDir1 As String
        Dim sFile As String

        Dim objValidaXML As ClassValidaXML = New ClassValidaXML
        Dim resFilialEmpresa As IEnumerable(Of FiliaisEmpresa)
        Dim resRPSWEBProt As IEnumerable(Of RPSWEBProt)
        Dim resFATConfig As IEnumerable(Of FATConfig)
        Dim resEndereco As IEnumerable(Of Endereco)
        Dim resNFeNFiscal As IEnumerable(Of NFeNFiscal)

        Dim iCodigoCancelamento As Integer

        Dim objNFeNFiscal As NFeNFiscal = New NFeNFiscal
        Dim objFATConfig As FATConfig
        Dim objRPSWEBProt As RPSWEBProt
        Dim objFilialEmpresa As FiliaisEmpresa
        Dim iRPSAmbiente As Integer

        Dim objListaMsgRetorno As ListaMensagemRetorno
        Dim objMsgRetorno As New tcMensagemRetorno
        Dim objCancelamento As New tcCancelamentoNfse
        Dim objConfirmacaoCancelamento As New tcConfirmacaoCancelamento
        Dim objPedidoCancelamento As New tcPedidoCancelamento
        Dim objIdentificacaoNfse As New tcIdentificacaoNfse
        Dim lRPSWEBRetCancNumIntDoc As Long
        Dim lRPSWEBLoteLogNumIntDoc As Long
        Dim sAux As String
        Dim objEndFilial As Endereco



        Try

            '********** pega o diretorio do log para colocar os arquivos xml *************
            iTamanho = 255
            sRetorno = StrDup(iTamanho, Chr(0))

            Call GetPrivateProfileString("Geral", "ArqLog", -1, sRetorno, iTamanho, "ADM100.INI")

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




            '***** coloca a string de conexao apontando para o SGEDados em questao *****
            odbc.ConnectionString = "DSN=SGEDados" & sEmpresa & ";UID=sa;PWD=SAPWD"

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


            db1.Connection.Open()
            db2.Connection.Open()

            db1.Transaction = db1.Connection.BeginTransaction()
            db2.Transaction = db2.Connection.BeginTransaction()


            resFilialEmpresa = db1.ExecuteQuery(Of FiliaisEmpresa) _
            ("SELECT * FROM FiliaisEmpresa WHERE FilialEmpresa = {0} ", iFilialEmpresa)

            For Each objFilialEmpresa In resFilialEmpresa
                sCertificado = objFilialEmpresa.CertificadoA1A3
                iRPSAmbiente = objFilialEmpresa.RPSAmbiente
                Exit For
            Next

            'lNumIntNFiscalParam
            resNFeNFiscal = db1.ExecuteQuery(Of NFeNFiscal) _
            ("SELECT * FROM NFeNFiscal WHERE NumIntDoc = {0} ", lNumIntNF)

            For Each objNFeNFiscal In resNFeNFiscal
                Exit For

            Next

            cert = certificado.BuscaNome(objFilialEmpresa.CertificadoA1A3)

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
                '               If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Or UCase(objEndFilial.Cidade) = "TATUÍ" Or UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then

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

            resFATConfig = db2.ExecuteQuery(Of FATConfig) _
            ("SELECT * FROM FatConfig WHERE Codigo = {0} ", "NUM_INT_PROX_RPSWEBLOTELOG")

            objFATConfig = resFATConfig(0)

            lRPSWEBLoteLogNumIntDoc = CLng(objFATConfig.Conteudo)

            resFATConfig = db2.ExecuteQuery(Of FATConfig) _
            ("SELECT * FROM FatConfig WHERE Codigo = {0} ", "NUM_INT_PROX_RPSWEBRETCANC")

            objFATConfig = resFATConfig(0)

            lRPSWEBRetCancNumIntDoc = CLng(objFATConfig.Conteudo)


            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, 0, Now.Date, TimeOfDay.ToOADate, "Iniciando o cancelamento da nota fiscal " & objNFeNFiscal.NumNotaFiscal, lNumIntNF)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            Form1.Msg.Items.Add("Iniciando o cancelamento da nota fiscal " & objNFeNFiscal.NumNotaFiscal)

            Application.DoEvents()

            resRPSWEBProt = db1.ExecuteQuery(Of RPSWEBProt) _
            ("SELECT * FROM RPSWEBProt WHERE NumIntNF = {0} AND FilialEmpresa = {1} AND Ambiente = {2} ORDER BY Numero DESC", lNumIntNF, iFilialEmpresa, iRPSAmbiente)

            objRPSWEBProt = resRPSWEBProt(0)

            If objRPSWEBProt Is Nothing Then
                Throw New System.Exception("A nota = " & objNFeNFiscal.NumNotaFiscal & " não foi autorizada no site da sefaz.")
            End If


#If Not TATUI2 Then

            Dim objInfPedidoCancelamento As New tcInfPedidoCancelamento

            objPedidoCancelamento.InfPedidoCancelamento = objInfPedidoCancelamento

            objInfPedidoCancelamento.Id = "pedido_cancelamento_" & objRPSWEBProt.Protocolo

            objInfPedidoCancelamento.IdentificacaoNfse = objIdentificacaoNfse
#If ABRASF2 Then
            If Len(objRPSWEBProt.PrestadorCNPJ) = 14 Then
                objInfPedidoCancelamento.IdentificacaoNfse.CpfCnpj = New tcCpfCnpj
                objInfPedidoCancelamento.IdentificacaoNfse.CpfCnpj.ItemElementName = ItemChoiceType.Cnpj
                objInfPedidoCancelamento.IdentificacaoNfse.CpfCnpj.Item = objRPSWEBProt.PrestadorCNPJ
            ElseIf Len(objRPSWEBProt.PrestadorCNPJ) = 11 Then
                objInfPedidoCancelamento.IdentificacaoNfse.CpfCnpj = New tcCpfCnpj
                objInfPedidoCancelamento.IdentificacaoNfse.CpfCnpj.ItemElementName = ItemChoiceType.Cpf
                objInfPedidoCancelamento.IdentificacaoNfse.CpfCnpj.Item = objRPSWEBProt.PrestadorCNPJ
            ElseIf Len(objFilialEmpresa.CGC) = 14 Then
                objInfPedidoCancelamento.IdentificacaoNfse.CpfCnpj = New tcCpfCnpj
                objInfPedidoCancelamento.IdentificacaoNfse.CpfCnpj.ItemElementName = ItemChoiceType.Cnpj
                objInfPedidoCancelamento.IdentificacaoNfse.CpfCnpj.Item = objFilialEmpresa.CGC
            ElseIf Len(objFilialEmpresa.CGC) = 11 Then
                objInfPedidoCancelamento.IdentificacaoNfse.CpfCnpj = New tcCpfCnpj
                objInfPedidoCancelamento.IdentificacaoNfse.CpfCnpj.ItemElementName = ItemChoiceType.Cpf
                objInfPedidoCancelamento.IdentificacaoNfse.CpfCnpj.Item = objFilialEmpresa.CGC

            End If
            objInfPedidoCancelamento.CodigoCancelamentoSpecified = True
            objInfPedidoCancelamento.CodigoCancelamento = Format(1, "0000")
            If Len(objRPSWEBProt.PrestadorInscMun) > 0 Then
                objInfPedidoCancelamento.IdentificacaoNfse.InscricaoMunicipal = objRPSWEBProt.PrestadorInscMun
            Else
                Call ADM.Formata_String_AlfaNumerico(objFilialEmpresa.InscricaoMunicipal, objInfPedidoCancelamento.IdentificacaoNfse.InscricaoMunicipal)
            End If

#Else
            objInfPedidoCancelamento.IdentificacaoNfse.Cnpj = objRPSWEBProt.PrestadorCNPJ
            objInfPedidoCancelamento.IdentificacaoNfse.InscricaoMunicipal = objRPSWEBProt.PrestadorInscMun
            objInfPedidoCancelamento.CodigoCancelamento = 2
#End If
            objInfPedidoCancelamento.IdentificacaoNfse.CodigoMunicipio = objRPSWEBProt.PrestadorCodMun
            objInfPedidoCancelamento.IdentificacaoNfse.Numero = objRPSWEBProt.Numero


            Dim mySerializer As New XmlSerializer(GetType(tcPedidoCancelamento))

            XMLStream = New MemoryStream(10000)

            mySerializer.Serialize(XMLStream, objPedidoCancelamento)

            Dim xm1 As Byte()
            xm1 = XMLStream.ToArray

            XMLString = System.Text.Encoding.UTF8.GetString(xm1)

            XMLString = Replace(XMLString, "<tcPedidoCancelamento", "<Pedido")
            XMLString = Replace(XMLString, "</tcPedidoCancelamento>", "</Pedido>")
            XMLString = Replace(XMLString, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""", "")
#If Not ABRASF2 Then
            XMLString = Replace(XMLString, "xmlns:xsd=""http://www.w3.org/2001/XMLSchema""", "xmlns=""http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd""")
#End If
            XMLString = Mid(XMLString, 1, InStr(XMLString, "InfPedidoCancelamento") - 1) & Replace(XMLString, "xmlns=""http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd""", "", InStr(XMLString, "InfPedidoCancelamento"))

            XMLString = Mid(XMLString, 22)

            'XMLString = Replace(XMLString, "xmlns=""http://www.abrasf.org.br/nfse.xsd""", "")

            objCanc.Pedido = New tcPedidoCancelamento

#Else


            objCanc.NumeroNfse = objRPSWEBProt.Numero

            objCanc.Prestador = New tcIdentificacaoPrestador

            objCanc.Prestador.Cnpj = objRPSWEBProt.PrestadorCNPJ
            objCanc.Prestador.InscricaoMunicipal = objRPSWEBProt.PrestadorInscMun

#End If


            Dim mySerializer2 As New XmlSerializer(GetType(CancelarNfseEnvio))

            XMLStream = New MemoryStream(10000)

            mySerializer2.Serialize(XMLStream, objCanc)

            Dim xm2 As Byte()
            xm2 = XMLStream.ToArray

            XMLString1 = System.Text.Encoding.UTF8.GetString(xm2)

            XMLString1 = Replace(XMLString1, " xmlns=""""", "")

            XMLString = Replace(XMLString1, "<Pedido />", XMLString)

            '            XMLString = Mid(XMLString, 22)

#If TATUI2 Then
            XMLString = Replace(XMLString, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""", "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.ginfes.com.br/servico_cancelar_nfse_envio_v02.xsd servico_cancelar_nfse_envio_v02.xsd""")
            XMLString = Replace(XMLString, "xmlns:xsd=""http://www.w3.org/2001/XMLSchema""", "xmlns:tipos=""http://www.ginfes.com.br/tipos""")
#Else

            XMLString = Replace(XMLString, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""", "")
            XMLString = Replace(XMLString, "xmlns:xsd=""http://www.w3.org/2001/XMLSchema""", "")
#End If
            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then

                XMLStringCabec = Replace(XMLStringCabec, "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd", "http://www.abrasf.org.br/nfse.xsd")
                XMLString = Replace(XMLString, "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd", "http://www.abrasf.org.br/nfse.xsd")

            End If

            If UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then
                '                If UCase(objEndFilial.Cidade) = "TATUÍ" Or UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then

                XMLString = Replace(XMLString, " xmlns=""http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd""", "")
                XMLString = Replace(XMLString, " xmlns=""http://www.ginfes.com.br/tipos""", "")
                XMLString = Replace(XMLString, " xmlns=""http://www.ginfes.com.br/servico_cancelar_nfse_envio""", "")
                '    XMLString = Replace(XMLString, "<CancelarNfseEnvio  >", "<CancelarNfseEnvio xmlns=""http://www.ginfes.com.br/servico_cancelar_nfse_envio"" xmlns:tipos=""http://www.ginfes.com.br/tipos"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.ginfes.com.br/servico_cancelar_nfse_envio_v02.xsd servico_cancelar_nfse_envio_v02.xsd"">")
                XMLString = Replace(XMLString, "<CancelarNfseEnvio", "<CancelarNfseEnvio xmlns=""http://www.ginfes.com.br/servico_cancelar_nfse_envio""")
                XMLString = Replace(XMLString, "Cnpj", "tipos:Cnpj")
                XMLString = Replace(XMLString, "InscricaoMunicipal", "tipos:InscricaoMunicipal")
                XMLString = Replace(XMLString, "xmlns=""http://www.ginfes.com.br/servico_cancelar_nfse_envio_v03.xsd""", "")
            End If


            If UCase(objEndFilial.Cidade) = "SALVADOR" Then

                XMLString = Replace(XMLString, "Id=", "id=")

            End If

#If TATUI2 Then
            AD.Assinar(XMLString, "Prestador", cert, objEndFilial.Cidade)
#Else
            AD.Assinar(XMLString, "InfPedidoCancelamento", cert, objEndFilial.Cidade)
#End If

            Dim xMlD As XmlDocument

            xMlD = AD.XMLDocAssinado()

            xString = AD.XMLStringAssinado

            '            xString = Mid(xString, 22)


#If Not TATUI2 Then
            xString = Replace(xString, "</Pedido>", "")
            xString = Replace(xString, "</CancelarNfseEnvio>", "</Pedido></CancelarNfseEnvio>")

#Else
            'xString = Replace(xString, "<CancelarNfseEnvio", "<CancelarNfse><CancelarNfseEnvio")
            'xString = Replace(xString, "</CancelarNfseEnvio>", "</CancelarNfseEnvio></CancelarNfse>")
#End If

            ' ''Dim a4 As cabecalho = New cabecalho
            ' ''Dim XMLStreamCabec As MemoryStream = New MemoryStream(10000)
            ' ''Dim XMLStringCabec As String

            ' ''a4.versao = "1.00"
            ' ''a4.versaoDados = "1.00"

            ' ''Dim mySerializercabec As New XmlSerializer(GetType(cabecalho))

            ' ''mySerializercabec.Serialize(XMLStreamCabec, a4)

            ' ''Dim doccabec As XmlDocument = New XmlDocument
            ' ''XMLStreamCabec.Position = 0

            ' ''Dim xmcabec As Byte()

            ' ''xmcabec = XMLStreamCabec.ToArray

            ' ''XMLStringCabec = System.Text.Encoding.UTF8.GetString(xmcabec)

            ' ''XMLStringCabec = Mid(XMLStringCabec, 1, 19) & " encoding=""utf-8"" " & Mid(XMLStringCabec, 20)

            '************* valida dados antes do envio **********************
            Dim xDados As Byte()

            xDados = System.Text.Encoding.UTF8.GetBytes(xString)

            XMLStreamDados = New MemoryStream(10000)

            XMLStreamDados.Write(xDados, 0, xDados.Length)


            Dim DocDados As XmlDocument = New XmlDocument
            XMLStreamDados.Position = 0
            DocDados.Load(XMLStreamDados)
            sArquivo = sDir & "Cancela_NFSE_" & lNumIntNF & ".xml"
            DocDados.Save(sArquivo)


            If UCase(objEndFilial.Cidade) = "BELO HORIZONTE" Then

                If iRPSAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then

                    homologacaobh.ClientCertificates.Add(cert)
                    XMLStringRetCancNFSE = homologacaobh.CancelarNfse(XMLStringCabec, xString)

                Else
                    producaobh.ClientCertificates.Add(cert)
                    XMLStringRetCancNFSE = producaobh.CancelarNfse(XMLStringCabec, xString)

                End If

                XMLStringRetCancNFSE = Replace(XMLStringRetCancNFSE, "http://www.abrasf.org.br/nfse.xsd", "http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd")

            ElseIf UCase(objEndFilial.Cidade) = "SÃO BERNARDO DO CAMPO" Or UCase(objEndFilial.Cidade) = "GUARULHOS" Then

                If iRPSAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then

                    homologacaoginfes.ClientCertificates.Add(cert)
                    XMLStringRetCancNFSE = homologacaoginfes.CancelarNfse(xString)

                Else
                    producaoginfes.ClientCertificates.Add(cert)
                    XMLStringRetCancNFSE = producaoginfes.CancelarNfse(xString)


                End If

                XMLStringRetCancNFSE = Replace(XMLStringRetCancNFSE, "ns2:", "")
                XMLStringRetCancNFSE = Replace(XMLStringRetCancNFSE, "ns3:", "")

                xString = Mid(XMLStringRetCancNFSE, 1, InStr(XMLStringRetCancNFSE, "<CancelarNfseResposta") + 20) & Mid(XMLStringRetCancNFSE, InStr(XMLStringRetCancNFSE, "><Sucesso>"))

                XMLStringRetCancNFSE = xString

                Dim sSucesso As String

                sSucesso = Mid(xString, InStr(xString, "<Sucesso>") + 9, InStr(xString, "</Sucesso>") - (InStr(xString, "<Sucesso>") + 9))
                objMsgRetorno.Codigo = Mid(xString, InStr(xString, "<Codigo>") + 8, InStr(xString, "</Codigo>") - (InStr(xString, "<Codigo>") + 8))
                objMsgRetorno.Mensagem = Mid(xString, InStr(xString, "<Mensagem>") + 10, InStr(xString, "</Mensagem>") - (InStr(xString, "<Mensagem>") + 10))
                objMsgRetorno.Correcao = Mid(xString, InStr(xString, "<Correcao>") + 10, InStr(xString, "</Correcao>") - (InStr(xString, "<Correcao>") + 10))

                iCodigoCancelamento = 0

                If objMsgRetorno.Codigo = "E79" Then
                    'esse codigo esta sendo colocado pois se a nota tiver sido cancelada e por alguma razao nao tiver sido gravado o codigo, o sistema possa indicar q foi feito o cancelamento
                    iCodigoCancelamento = 1
                End If

                If UCase(sSucesso) = "TRUE" Then iCodigoCancelamento = 1

                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBRetCanc ( NumIntDoc, FilialEmpresa, tpAmb, CodigoCancelamento, NumIntNF, CodMsg, Msg, Correcao, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9})", _
                lRPSWEBRetCancNumIntDoc, iFilialEmpresa, iRPSAmbiente, iCodigoCancelamento, lNumIntNF, objMsgRetorno.Codigo, Left(objMsgRetorno.Mensagem, 200), Left(IIf(Len(objMsgRetorno.Correcao) = 0, "", objMsgRetorno.Correcao), 200), Now.Date, TimeOfDay.ToOADate)

                lRPSWEBRetCancNumIntDoc = lRPSWEBRetCancNumIntDoc + 1

                iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
                lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, 0, Now.Date, TimeOfDay.ToOADate, Left("Retorno do cancelamento da nota fiscal " & objNFeNFiscal.NumNotaFiscal & "  Codigo = " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao, 255), 0)

                lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                Form1.Msg.Items.Add("Retorno do cancelamento da nota fiscal " & objNFeNFiscal.NumNotaFiscal & "  Codigo = " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao)

                Application.DoEvents()


                'If InStr(XMLStringRetCancNFSE, "<MensagemRetorno>") <> 0 Then

                '    xString = Mid(XMLStringRetCancNFSE, 1, InStr(XMLStringRetCancNFSE, "<Sucesso>") - 1)

                '    XMLStringRetCancNFSE = xString & Mid(XMLStringRetCancNFSE, InStr(XMLStringRetCancNFSE, "<MensagemRetorno>"))

                '    XMLStringRetCancNFSE = Replace(XMLStringRetCancNFSE, "<CancelarNfseResposta>", "<CancelarNfseResposta xmlns=""http://www.ginfes.com.br/servico_cancelar_nfse_resposta_v03.xsd"">")

                '    XMLStringRetCancNFSE = Replace(XMLStringRetCancNFSE, "</CancelarNfseResposta>", "</ListaMensagemRetorno></CancelarNfseResposta>")

                '    '                    XMLStringRetCancNFSE = Replace(XMLStringRetCancNFSE, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>", "<?xml version=""1.0""?>")



                ' End If


            ElseIf UCase(objEndFilial.Cidade) = "TATUÍ" Then

                If iRPSAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then

                    homsistema4rCancela.ClientCertificates.Add(cert)
                    XMLStringRetCancNFSE = homsistema4rCancela.Execute(xString)

                Else
                    prodsistema4rCancela.ClientCertificates.Add(cert)
                    XMLStringRetCancNFSE = prodsistema4rCancela.Execute(xString)
                End If


            Else


                If iRPSAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then

                    homologacao.ClientCredentials.ClientCertificate.Certificate = cert
                    XMLStringRetCancNFSE = homologacao.CancelarNfse(xString)

                Else
                    producao.ClientCredentials.ClientCertificate.Certificate = cert
                    XMLStringRetCancNFSE = producao.CancelarNfse(xString)

                End If

            End If

#If Not TATUI2 Then

            Dim xRet As Byte()

            '            XMLStringRetCancNFSE = Mid(XMLStringRetCancNFSE, InStr(XMLStringRetCancNFSE, "<CancelarNfseResposta>"))

            xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetCancNFSE)

            XMLStreamRet = New MemoryStream(10000)

            XMLStreamRet.Write(xRet, 0, xRet.Length)

            Dim mySerializerRetEnvNFe As New XmlSerializer(GetType(CancelarNfseResposta))

            Dim objCancelarNfseResposta As CancelarNfseResposta = New CancelarNfseResposta

            XMLStreamRet.Position = 0

            objCancelarNfseResposta = mySerializerRetEnvNFe.Deserialize(XMLStreamRet)

            sAux = objCancelarNfseResposta.Item.GetType.ToString()
            If InStr(sAux, ".") <> 0 Then
                sAux = Mid(sAux, InStr(sAux, ".") + 1)
            End If


#If Not RJ Then
            If sAux = "RetCancelamento" Then


                Dim objRetCancelamento As RetCancelamento

                objRetCancelamento = objCancelarNfseResposta.Item

                objCancelamento = objRetCancelamento.NfseCancelamento(0)
            End If

#Else
            If sAux = "tcCancelamentoNfse" Then

                objCancelamento = objCancelarNfseResposta.Item

            End If
#End If

            Select Case sAux


                Case "ListaMensagemRetorno"

                    objListaMsgRetorno = objCancelarNfseResposta.Item

                    For iIndice1 = 0 To objListaMsgRetorno.MensagemRetorno.Count - 1

                        objMsgRetorno = objListaMsgRetorno.MensagemRetorno(iIndice1)

                        iCodigoCancelamento = 0

                        If objMsgRetorno.Codigo = "E79" Then
                            'esse codigo esta sendo colocado pois se a nota tiver sido cancelada e por alguma razao nao tiver sido gravado o codigo, o sistema possa indicar q foi feito o cancelamento
                            iCodigoCancelamento = 1
                        End If



                        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBRetCanc ( NumIntDoc, FilialEmpresa, tpAmb, CodigoCancelamento, NumIntNF, CodMsg, Msg, Correcao, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9})", _
                        lRPSWEBRetCancNumIntDoc, iFilialEmpresa, iRPSAmbiente, iCodigoCancelamento, lNumIntNF, objMsgRetorno.Codigo, Left(objMsgRetorno.Mensagem, 200), Left(IIf(Len(objMsgRetorno.Correcao) = 0, "", objMsgRetorno.Correcao), 200), Now.Date, TimeOfDay.ToOADate)

                        lRPSWEBRetCancNumIntDoc = lRPSWEBRetCancNumIntDoc + 1

                        iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
                        lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, 0, Now.Date, TimeOfDay.ToOADate, Left("Retorno do cancelamento da RPS " & objNFeNFiscal.NumNotaFiscal & "  Codigo = " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao, 255), 0)

                        lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                        Form1.Msg.Items.Add("Retorno do cancelamento da RPS " & objNFeNFiscal.NumNotaFiscal & "  Codigo = " & objMsgRetorno.Codigo & " - " & objMsgRetorno.Mensagem & " - " & objMsgRetorno.Correcao)

                        Application.DoEvents()

                    Next

                Case "tcCancelamentoNfse", "RetCancelamento"


                    objConfirmacaoCancelamento = objCancelamento.Confirmacao

                    If objConfirmacaoCancelamento Is Nothing Then
                        objConfirmacaoCancelamento = New nfseabrasf.tcConfirmacaoCancelamento
                    End If

                    objPedidoCancelamento = objConfirmacaoCancelamento.Pedido

                    If objPedidoCancelamento Is Nothing Then
                        objPedidoCancelamento = New nfseabrasf.tcPedidoCancelamento
                    End If

                    If Not objPedidoCancelamento.InfPedidoCancelamento Is Nothing Then
                        objIdentificacaoNfse = objPedidoCancelamento.InfPedidoCancelamento.IdentificacaoNfse
                    Else
                        objPedidoCancelamento.InfPedidoCancelamento = New nfseabrasf.tcInfPedidoCancelamento
                    End If

                    If objIdentificacaoNfse Is Nothing Then
                        objIdentificacaoNfse = New nfseabrasf.tcIdentificacaoNfse
                    End If


                    Dim sCNPJ As String

#If ABRASF2 Then
                    If objIdentificacaoNfse.CpfCnpj.Item = Nothing Then
                        sCNPJ = ""
                    Else
                        sCNPJ = objIdentificacaoNfse.CpfCnpj.Item
                    End If

#Else
                    If objIdentificacaoNfse.Cnpj = Nothing Then
                        sCNPJ = ""
                    Else
                        sCNPJ = objIdentificacaoNfse.Cnpj
                    End If
#End If


#If TATUI2 Then
                    iResult = db1.ExecuteCommand("INSERT INTO RPSWEBRetCanc ( NumIntDoc, FilialEmpresa, NumIntNF, versao, tpAmb, CodigoCancelamento, Id,CNPJ, CodigoMunicipio, InscricaoMunicipal, Numero, IdConfirmacao, Data, Hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, (9), {10}, {11}, {12}, {13} )", _
                      lRPSWEBRetCancNumIntDoc, iFilialEmpresa, lNumIntNF, "", iRPSAmbiente, IIf(objPedidoCancelamento.InfPedidoCancelamento.CodigoCancelamento = Nothing, 0, objPedidoCancelamento.InfPedidoCancelamento.CodigoCancelamento), IIf(objPedidoCancelamento.InfPedidoCancelamento.Id = Nothing, "", objPedidoCancelamento.InfPedidoCancelamento.Id), sCNPJ, IIf(objIdentificacaoNfse.CodigoMunicipio = Nothing, "", objIdentificacaoNfse.CodigoMunicipio), _
                      IIf(objIdentificacaoNfse.InscricaoMunicipal = Nothing, "", objIdentificacaoNfse.InscricaoMunicipal), IIf(objIdentificacaoNfse.Numero = Nothing, "", objIdentificacaoNfse.Numero), IIf(objConfirmacaoCancelamento.Id = Nothing, "", objConfirmacaoCancelamento.Id), IIf(objConfirmacaoCancelamento.DataHoraCancelamento Is Nothing, Now.Date, objConfirmacaoCancelamento.DataHoraCancelamento.Date), IIf(objConfirmacaoCancelamento.DataHoraCancelamento Is Nothing, TimeOfDay.ToOADate, objConfirmacaoCancelamento.DataHoraCancelamento.TimeOfDay.TotalDays))
#ElseIf RJ Then
                    iResult = db1.ExecuteCommand("INSERT INTO RPSWEBRetCanc ( NumIntDoc, FilialEmpresa, NumIntNF, versao, tpAmb, CodigoCancelamento, Id,CNPJ, CodigoMunicipio, InscricaoMunicipal, Numero, IdConfirmacao, Data, Hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, (9), {10}, {11}, {12}, {13} )", _
                      lRPSWEBRetCancNumIntDoc, iFilialEmpresa, lNumIntNF, "", iRPSAmbiente, IIf(objPedidoCancelamento.InfPedidoCancelamento.CodigoCancelamento = Nothing, 0, objPedidoCancelamento.InfPedidoCancelamento.CodigoCancelamento), IIf(objPedidoCancelamento.InfPedidoCancelamento.Id = Nothing, "", objPedidoCancelamento.InfPedidoCancelamento.Id), sCNPJ, IIf(objIdentificacaoNfse.CodigoMunicipio = Nothing, "", objIdentificacaoNfse.CodigoMunicipio), _
                      IIf(objIdentificacaoNfse.Numero = Nothing, "", objIdentificacaoNfse.Numero), IIf(objIdentificacaoNfse.InscricaoMunicipal = Nothing, "", objIdentificacaoNfse.InscricaoMunicipal), IIf(objConfirmacaoCancelamento.Id = Nothing, "", objConfirmacaoCancelamento.Id), IIf(objConfirmacaoCancelamento.DataHoraCancelamento = Nothing, Now.Date, objConfirmacaoCancelamento.DataHoraCancelamento.Date), IIf(objConfirmacaoCancelamento.DataHoraCancelamento = Nothing, TimeOfDay.ToOADate, objConfirmacaoCancelamento.DataHoraCancelamento.TimeOfDay.TotalDays))

#Else
                    iResult = db1.ExecuteCommand("INSERT INTO RPSWEBRetCanc ( NumIntDoc, FilialEmpresa, NumIntNF, versao, tpAmb, CodigoCancelamento, Id,CNPJ, CodigoMunicipio, InscricaoMunicipal, Numero, IdConfirmacao, Data, Hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, (9), {10}, {11}, {12}, {13} )", _
                      lRPSWEBRetCancNumIntDoc, iFilialEmpresa, lNumIntNF, "", iRPSAmbiente, IIf(objPedidoCancelamento.InfPedidoCancelamento.CodigoCancelamento = Nothing, 0, objPedidoCancelamento.InfPedidoCancelamento.CodigoCancelamento), IIf(objPedidoCancelamento.InfPedidoCancelamento.Id = Nothing, "", objPedidoCancelamento.InfPedidoCancelamento.Id), sCNPJ, IIf(objIdentificacaoNfse.CodigoMunicipio = Nothing, "", objIdentificacaoNfse.CodigoMunicipio), _
                      IIf(objIdentificacaoNfse.Numero = Nothing, "", objIdentificacaoNfse.Numero), IIf(objIdentificacaoNfse.InscricaoMunicipal = Nothing, "", objIdentificacaoNfse.InscricaoMunicipal), IIf(objConfirmacaoCancelamento.Id = Nothing, "", objConfirmacaoCancelamento.Id), IIf(objConfirmacaoCancelamento.DataHora = Nothing, Now.Date, objConfirmacaoCancelamento.DataHora.Date), IIf(objConfirmacaoCancelamento.DataHora = Nothing, TimeOfDay.ToOADate, objConfirmacaoCancelamento.DataHora.TimeOfDay.TotalDays))

#End If


                    lRPSWEBRetCancNumIntDoc = lRPSWEBRetCancNumIntDoc + 1

                    iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
                    lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, 0, Now.Date, TimeOfDay.ToOADate, "RPS " & objNFeNFiscal.NumNotaFiscal & " cancelada. Numero da NFSE = " & objIdentificacaoNfse.Numero, 0)

                    lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

                    Form1.Msg.Items.Add("RPS " & objNFeNFiscal.NumNotaFiscal & " cancelada. Numero da NFSE = " & objIdentificacaoNfse.Numero)

                    Application.DoEvents()

            End Select
#End If

            db1.Transaction.Commit()

        Catch ex As Exception When objRPSWEBProt Is Nothing

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBRetCanc ( NumIntDoc, FilialEmpresa, NumIntNF, tpAmb, Msg, Data, Hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6} )", _
            lRPSWEBRetCancNumIntDoc, iFilialEmpresa, lNumIntNF, iRPSAmbiente, "a RPS " & objNFeNFiscal.NumNotaFiscal & " nao esta autorizada", Now.Date, TimeOfDay.ToOADate)

            lRPSWEBRetCancNumIntDoc = lRPSWEBRetCancNumIntDoc + 1

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, 0, Now.Date, TimeOfDay.ToOADate, "a RPS " & objNFeNFiscal.NumNotaFiscal & " nao esta autorizada", 0)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, 0, Now.Date, TimeOfDay.ToOADate, "ERRO - Encerrado o cancelamento da RPS " & objNFeNFiscal.NumNotaFiscal, lNumIntNF)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            Form1.Msg.Items.Add("ERRO - o cancelamento da nota fiscal NumIntDoc = " & CStr(lNumIntNF) & " foi encerrado por erro. Erro = a nota nao esta autorizada")

            Application.DoEvents()

            db1.Transaction.Rollback()

        Catch ex As Exception

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, 0, Now.Date, TimeOfDay.ToOADate, Replace(Left(ex.Message, 255), "'", "*"), lNumIntNF)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            iResult = db2.ExecuteCommand("INSERT INTO RPSWEBLoteLog ( NumIntDoc, Ambiente, FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", _
            lRPSWEBLoteLogNumIntDoc, iRPSAmbiente, iFilialEmpresa, 0, Now.Date, TimeOfDay.ToOADate, "ERRO - Encerrado o cancelamento da RPS " & objNFeNFiscal.NumNotaFiscal, lNumIntNF)

            lRPSWEBLoteLogNumIntDoc = lRPSWEBLoteLogNumIntDoc + 1

            Form1.Msg.Items.Add("ERRO - o cancelamento da RPS " & objNFeNFiscal.NumNotaFiscal & " foi encerrado por erro. Erro = " & ex.Message)

            Application.DoEvents()

            db1.Transaction.Rollback()

        Finally


            If lRPSWEBLoteLogNumIntDoc <> 0 Then
                iResult = db2.ExecuteCommand("UPDATE FatConfig Set Conteudo = {0} WHERE Codigo = {1}", lRPSWEBLoteLogNumIntDoc, "NUM_INT_PROX_RPSWEBLOTELOG")
            End If

            If lRPSWEBRetCancNumIntDoc <> 0 Then
                iResult = db2.ExecuteCommand("UPDATE FatConfig Set Conteudo = {0} WHERE Codigo = {1}", lRPSWEBRetCancNumIntDoc, "NUM_INT_PROX_RPSWEBRETCANC")
            End If


            db2.Transaction.Commit()

            db1.Connection.Close()
            db2.Connection.Close()


            db1.Dispose()
            db2.Dispose()

        End Try

#End If
    End Function
End Class
