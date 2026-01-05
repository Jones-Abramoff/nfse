Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.Schema

Public Class AssinaturaDigital

    '//
    '// mensagem de Retorno
    '//
    Private msgResultado As String
    Private XMLDoc As XmlDocument

    Public Function XMLDocAssinado() As XmlDocument
        XMLDocAssinado = XMLDoc
    End Function

    Public Function XMLStringAssinado() As String
        XMLStringAssinado = XMLDoc.OuterXml
    End Function

    Public Function mensagemResultado() As String
        mensagemResultado = msgResultado
    End Function

    Public Function Assinar(ByVal XMLString As String, ByVal RefUri As String, ByVal X509Cert As X509Certificate2, ByVal sCidade As String) As Integer
        '     Entradas:
        '         XMLString: string XML a ser assinada
        '         RefUri   : Referência da URI a ser assinada (Ex. infNFe
        '         X509Cert : certificado digital a ser utilizado na assinatura digital
        ' 
        '     Retornos:
        '         Assinar : 0 - Assinatura realizada com sucesso
        '                   1 - Erro: Problema ao acessar o certificado digital - %exceção%
        '                   2 - Problemas no certificado digital
        '                  3 - XML mal formado + exceção
        '                   4 - A tag de assinatura %RefUri% inexiste
        '                   5 - A tag de assinatura %RefUri% não é unica
        '                   6 - Erro Ao assinar o documento - ID deve ser string %RefUri(Atributo)%
        '                   7 - Erro: Ao assinar o documento - %exceção%
        ' 
        '        XMLStringAssinado : string XML assinada
        ' 
        '         XMLDocAssinado    : XMLDocument do XML assinado
        '

        Dim resultado As Integer = 0
        msgResultado = "Assinatura realizada com sucesso"
        Try
            '   certificado para ser utilizado na assinatura
            '
            Dim _xnome As String = ""
            If (Not X509Cert Is Nothing) Then
                _xnome = X509Cert.Subject.ToString()
            End If

            Dim _X509Cert As X509Certificate2 = New X509Certificate2()
            Dim store As X509Store = New X509Store("MY", StoreLocation.CurrentUser)

            store.Open(OpenFlags.ReadOnly Or OpenFlags.OpenExistingOnly)
            Dim collection As X509Certificate2Collection = store.Certificates
            '(X509Certificate2Collection)g
            Dim collection1 As X509Certificate2Collection = collection.Find(X509FindType.FindBySubjectDistinguishedName, _xnome, False)
            '(X509Certificate2Collection)
            If (collection1.Count = 0) Then
                resultado = 2
                msgResultado = "Problemas no certificado digital"
            Else
                ' certificado ok
                _X509Cert = collection1(0)
                Dim x As String
                x = _X509Cert.GetKeyAlgorithm().ToString()
                ' Create a new XML document.
                Dim doc As XmlDocument = New XmlDocument()

                ' Format the document to ignore white spaces.
                doc.PreserveWhitespace = False

                ' Load the passed XML file using it's name.
                Try

                    doc.LoadXml(XMLString)

                    ' Verifica se a tag a ser assinada existe é única
                    Dim qtdeRefUri As Integer = doc.GetElementsByTagName(RefUri).Count

                    If (qtdeRefUri = 0) Then
                        '  a URI indicada não existe
                        resultado = 4
                        msgResultado = "A tag de assinatura " + RefUri.Trim() + " inexiste"
                        ' Exsiste mais de uma tag a ser assinada
                    Else

                        If (qtdeRefUri > 1) Then
                            ' existe mais de uma URI indicada
                            resultado = 5
                            msgResultado = "A tag de assinatura " + RefUri.Trim() + " não é unica"

                            '//else if (_listaNum.IndexOf(doc.GetElementsByTagName(RefUri).Item(0).Attributes.ToString().Substring(1,1))>0)
                            '//{
                            '//    resultado = 6;
                            '//    msgResultado = "Erro: Ao assinar o documento - ID deve ser string (" + doc.GetElementsByTagName(RefUri).Item(0).Attributes + ")";
                            '//}
                        Else
                            Try

                                ' Create a SignedXml object.
                                Dim SignedXml As SignedXml = New SignedXml(doc)

                                ' Add the key to the SignedXml document 



                                SignedXml.SigningKey = _X509Cert.PrivateKey

                                ' Create a reference to be signed
                                Dim reference As Reference = New Reference()
                                ' pega o uri que deve ser assinada
                                Dim _Uri As XmlAttributeCollection = doc.GetElementsByTagName(RefUri).Item(0).Attributes
                                Dim _atributo As XmlAttribute
                                'novo colocado para corrigir o problema
                                reference.Uri = ""

                                For Each _atributo In _Uri

                                    If UCase(sCidade) = "SALVADOR" Then
                                        If (_atributo.Name = "id") Then
                                            reference.Uri = "#" + _atributo.InnerText
                                        End If

                                    Else
                                        If (_atributo.Name = "Id") Then
                                            reference.Uri = "#" + _atributo.InnerText
                                        End If
                                    End If
                                Next

                                ' Add an enveloped transformation to the reference.
                                Dim env As XmlDsigEnvelopedSignatureTransform = New XmlDsigEnvelopedSignatureTransform()
                                reference.AddTransform(env)

                                Dim c14 As XmlDsigC14NTransform = New XmlDsigC14NTransform()
                                reference.AddTransform(c14)

                                ' Add the reference to the SignedXml object.
                                SignedXml.AddReference(reference)

                                '// Create a new KeyInfo object
                                Dim keyInfo As KeyInfo = New KeyInfo()

                                '// Load the certificate into a KeyInfoX509Data object
                                '// and add it to the KeyInfo object.
                                keyInfo.AddClause(New KeyInfoX509Data(_X509Cert))

                                '// Add the KeyInfo object to the SignedXml object.
                                SignedXml.KeyInfo = keyInfo

                                SignedXml.ComputeSignature()

                                '// Get the XML representation of the signature and save
                                '// it to an XmlElement object.
                                Dim xmlDigitalSignature As XmlElement = SignedXml.GetXml()

                                '// Append the element to the XML document.
                                doc.DocumentElement.AppendChild(doc.ImportNode(xmlDigitalSignature, True))
                                XMLDoc = New XmlDocument()
                                XMLDoc.PreserveWhitespace = False
                                XMLDoc = doc

                            Catch caught As Exception
                                resultado = 7
                                msgResultado = "Erro: Ao assinar o documento - " + caught.Message
                            End Try
                        End If
                    End If
                Catch caught As Exception
                    resultado = 3
                    msgResultado = "Erro: XML mal formado - " + caught.Message
                End Try
            End If
        Catch caught As Exception

            resultado = 1
            msgResultado = "Erro: Problema ao acessar o certificado digital" + caught.Message
        End Try
        Assinar = resultado

    End Function


    ''' <summary>
    ''' Exemplo de como assinar uma mensagem XML com um certificado digital.
    ''' Linguagem: VB.NET
    ''' Framework: 3.5
    ''' </summary>
    ''' <param name="mensagemXML">String contendo a própria mensagem XML</param>
    ''' <param name="certificado">O certificado que será usado para assinar
    ''' a mensagam XML</param>
    ''' <returns>Um objeto do tipo XmlDocument já assinado</returns>
    ''' <remarks></remarks>
    Public Function Assinar2(ByVal mensagemXML As String, _
    ByVal certificado As  _
    System.Security.Cryptography.X509Certificates.X509Certificate2) _
    As XmlDocument
        Dim xmlDoc As New System.Xml.XmlDocument()
        Dim Key As New System.Security.Cryptography.RSACryptoServiceProvider()
        Dim SignedDocument As System.Security.Cryptography.Xml.SignedXml
        Dim keyInfo As New System.Security.Cryptography.Xml.KeyInfo()
        xmlDoc.LoadXml(mensagemXML)
        'Retira chave privada ligada ao certificado
        Key = CType(certificado.PrivateKey,  _
        System.Security.Cryptography.RSACryptoServiceProvider)
        'Adiciona Certificado ao Key Info
        keyInfo.AddClause(New  _
        System.Security.Cryptography.Xml.KeyInfoX509Data(certificado))
        SignedDocument = New System.Security.Cryptography.Xml.SignedXml(xmlDoc)
        'Seta chaves
        SignedDocument.SigningKey = Key
        SignedDocument.KeyInfo = keyInfo
        ' Cria referencia
        Dim reference As New System.Security.Cryptography.Xml.Reference()
        reference.Uri = String.Empty
        ' Adiciona transformacao a referencia
        reference.AddTransform(New  _
        System.Security.Cryptography.Xml.XmlDsigEnvelopedSignatureTransform())
        reference.AddTransform(New  _
        System.Security.Cryptography.Xml.XmlDsigC14NTransform(False))
        ' Adiciona referencia ao xml
        SignedDocument.AddReference(reference)
        ' Calcula Assinatura
        SignedDocument.ComputeSignature()
        ' Pega representação da assinatura
        Dim xmlDigitalSignature As System.Xml.XmlElement = SignedDocument.GetXml()
        ' Adiciona ao doc XML
        xmlDoc.DocumentElement.AppendChild(xmlDoc.ImportNode(xmlDigitalSignature, True))
        Return xmlDoc
    End Function

    ''' <summary>
    ''' Exemplo de como assinar elementos de um documento XML com um certificado digital.
    ''' Linguagem: VB.NET
    ''' Framework: 3.5
    ''' </summary>
    ''' <param name="Document">
    ''' O documento que contem os elementos que devem ser assinados
    ''' </param>
    ''' <param name="ParentElementName">O elemento que contem as tags a serem assinadas
    ''' Ex.:
    ''' Para assinar o rps em um lote o ParentElementName = "Rps"
    ''' Para assinar o pedido do lote de rps o ParentElementName = "EnviarLoteRpsEnvio"
    ''' </param>
    ''' <param name="ElementName">O elemento (tag) que será assinado
    ''' Ex.:
    ''' Para assinar o rps em um lote o ElementName = "InfRps"
    ''' Para assinar o pedido do lote de rps o ElementName = "LoteRps"
    ''' </param>
    ''' <param name="AttributeName">
    ''' O nome do atributo do elemento que será assinado
    ''' Obs.:
    ''' Por padrão o NOME do atributo possui a mesma identificação.
    ''' AtributeName = "Id"
    ''' </param>
    Public Sub AssinarElementos(ByVal Document As XmlDocument, _
    ByVal x509 As X509Certificate2, _
    ByVal ParentElementName As String, _
    ByVal ElementName As String, _
    ByVal AttributeName As String)
        Dim el As XmlElement
        Dim elInf As XmlElement
        Dim elInfID As String
        Dim elSigned As SignedXml
        Dim Key As RSACryptoServiceProvider
        Dim keyInfo As New KeyInfo
        'Retira chave privada ligada ao certificado
        Key = CType(x509.PrivateKey, RSACryptoServiceProvider)
        'Adiciona Certificado ao Key Info
        keyInfo.AddClause(New KeyInfoX509Data(x509))
        For Each el In Document.GetElementsByTagName(ParentElementName)
            elInf = _
            CType(el.GetElementsByTagName(ElementName)(el.GetElementsByTagName(ElementName).Count - 1),  _
            XmlElement)
            elInfID = elInf.Attributes.GetNamedItem(AttributeName).Value
            elSigned = New SignedXml(elInf)
            'Seta chaves
            elSigned.SigningKey = Key
            elSigned.KeyInfo = keyInfo
            ' Cria referencia
            Dim reference As New Reference()
            reference.Uri = "#" & elInfID
            ' Adiciona tranformacao a referencia
            reference.AddTransform(New XmlDsigEnvelopedSignatureTransform())
            reference.AddTransform(New XmlDsigC14NTransform(False))
            ' Adiciona referencia ao xml
            elSigned.AddReference(reference)
            ' Calcula Assinatura
            elSigned.ComputeSignature()
            'Adiciona assinatura
            el.AppendChild(Document.ImportNode(elSigned.GetXml(), True))
        Next
    End Sub
End Class

Public Class Certificado

    Public Function BuscaNome(ByVal Nome As String) As X509Certificate2

        Dim _X509Cert As X509Certificate2 = New X509Certificate2()
        Try

            Dim store As X509Store = New X509Store("MY", StoreLocation.CurrentUser)
            store.Open(OpenFlags.OpenExistingOnly Or OpenFlags.IncludeArchived Or OpenFlags.ReadWrite)
            Dim collection As X509Certificate2Collection = store.Certificates
            Dim collection1 As X509Certificate2Collection = collection.Find(X509FindType.FindByTimeValid, DateTime.Now, False)
            Dim collection2 As X509Certificate2Collection = collection.Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, False)
            If Nome = "" Then
                Dim scollection As X509Certificate2Collection = X509Certificate2UI.SelectFromCollection(collection1, "Certificado(s) Digital(is) disponível(is)", "Selecione o Certificado Digital para uso no aplicativo", X509SelectionFlag.SingleSelection)
                If (scollection.Count = 0) Then
                    _X509Cert.Reset()
                    Console.WriteLine("Nenhum certificado escolhido", "Atenção")
                Else
                    _X509Cert = scollection(0)
                End If
            Else
                Dim scollection As X509Certificate2Collection = collection2.Find(X509FindType.FindBySubjectName, Nome, False)
                If (scollection.Count = 0) Then
                    Console.WriteLine("Nenhum certificado válido foi encontrado com o nome informado: " + Nome, "Atenção")
                    _X509Cert.Reset()
                Else
                    _X509Cert = scollection(0)
                End If
            End If
            store.Close()
            BuscaNome = _X509Cert

        Catch ex As SystemException
            Console.WriteLine(ex.Message)
            BuscaNome = _X509Cert
        End Try
    End Function

    Public Function BuscaNroSerie(ByVal NroSerie As String) As X509Certificate2
        Dim _X509Cert As X509Certificate2 = New X509Certificate2()
        Try

            Dim store As X509Store = New X509Store("My", StoreLocation.CurrentUser)
            store.Open(OpenFlags.ReadOnly Or OpenFlags.OpenExistingOnly)
            Dim collection As X509Certificate2Collection = store.Certificates
            Dim collection1 As X509Certificate2Collection = collection.Find(X509FindType.FindByTimeValid, DateTime.Now, True)
            Dim collection2 As X509Certificate2Collection = collection1.Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, True)
            If (NroSerie = "") Then
                Dim scollection As X509Certificate2Collection = X509Certificate2UI.SelectFromCollection(collection2, "Certificados Digitais", "Selecione o Certificado Digital para uso no aplicativo", X509SelectionFlag.SingleSelection)
                If (scollection.Count = 0) Then
                    _X509Cert.Reset()
                    Console.WriteLine("Nenhum certificado válido foi encontrado com o número de série informado: " + NroSerie, "Atenção")
                Else
                    _X509Cert = scollection(0)
                End If
            Else
                Dim scollection As X509Certificate2Collection = collection2.Find(X509FindType.FindBySerialNumber, NroSerie, True)
                If (scollection.Count = 0) Then
                    _X509Cert.Reset()
                    Console.WriteLine("Nenhum certificado válido foi encontrado com o número de série informado: " + NroSerie, "Atenção")
                Else
                    _X509Cert = scollection(0)
                End If
            End If
            store.Close()
            Return _X509Cert
        Catch ex As System.Exception
            Console.WriteLine(ex.Message)
            Return _X509Cert
        End Try

    End Function

    ''' <summary>
    ''' Exemplo de como assinar uma mensagem XML com um certificado digital.
    ''' Linguagem: VB.NET
    ''' Framework: 3.5
    ''' </summary>
    ''' <param name="mensagemXML">String contendo a própria mensagem XML</param>
    ''' <param name="certificado">O certificado que será usado para assinar
    ''' a mensagam XML</param>
    ''' <returns>Um objeto do tipo XmlDocument já assinado</returns>
    ''' <remarks></remarks>
    Public Function Assinar3(ByVal mensagemXML As String, _
    ByVal certificado As  _
    System.Security.Cryptography.X509Certificates.X509Certificate2) _
    As XmlDocument
        Dim xmlDoc As New System.Xml.XmlDocument()
        Dim Key As New System.Security.Cryptography.RSACryptoServiceProvider()
        Dim SignedDocument As System.Security.Cryptography.Xml.SignedXml
        Dim keyInfo As New System.Security.Cryptography.Xml.KeyInfo()
        xmlDoc.LoadXml(mensagemXML)
        'Retira chave privada ligada ao certificado
        Key = CType(certificado.PrivateKey,  _
        System.Security.Cryptography.RSACryptoServiceProvider)
        'Adiciona Certificado ao Key Info
        keyInfo.AddClause(New  _
        System.Security.Cryptography.Xml.KeyInfoX509Data(certificado))
        SignedDocument = New System.Security.Cryptography.Xml.SignedXml(xmlDoc)
        'Seta chaves
        SignedDocument.SigningKey = Key
        SignedDocument.KeyInfo = keyInfo
        ' Cria referencia
        Dim reference As New System.Security.Cryptography.Xml.Reference()
        reference.Uri = String.Empty
        ' Adiciona transformacao a referencia
        reference.AddTransform(New  _
        System.Security.Cryptography.Xml.XmlDsigEnvelopedSignatureTransform())
        reference.AddTransform(New  _
        System.Security.Cryptography.Xml.XmlDsigC14NTransform(False))
        ' Adiciona referencia ao xml
        SignedDocument.AddReference(reference)
        ' Calcula Assinatura
        SignedDocument.ComputeSignature()
        ' Pega representação da assinatura
        Dim xmlDigitalSignature As System.Xml.XmlElement = SignedDocument.GetXml()
        ' Adiciona ao doc XML
        xmlDoc.DocumentElement.AppendChild(xmlDoc.ImportNode(xmlDigitalSignature, True))
        Return xmlDoc
    End Function

End Class
