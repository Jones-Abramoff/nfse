Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.IO.Compression
Imports System.Security.Cryptography.X509Certificates
Imports System.Security.Cryptography.Xml
Imports System.Text
Imports System.Xml

Public NotInheritable Class XmlUtil
    Private Sub New()
    End Sub

    Public Shared Function GZipToBase64(xmlUtf8 As String) As String
        Dim bytes = Encoding.UTF8.GetBytes(xmlUtf8)
        Using ms As New MemoryStream()
            Using gz As New GZipStream(ms, CompressionMode.Compress, True)
                gz.Write(bytes, 0, bytes.Length)
            End Using
            Return Convert.ToBase64String(ms.ToArray())
        End Using
    End Function

    Public Shared Function Base64GunzipToString(b64 As String) As String
        Dim data = Convert.FromBase64String(b64)
        Using inMs As New MemoryStream(data)
            Using gz As New GZipStream(inMs, CompressionMode.Decompress)
                Using outMs As New MemoryStream()
                    gz.CopyTo(outMs)
                    Return Encoding.UTF8.GetString(outMs.ToArray())
                End Using
            End Using
        End Using
    End Function

    Public Shared Function AssinarDps(xml As String, cert As X509Certificate2) As String

        Dim doc As New XmlDocument()
        doc.PreserveWhitespace = True
        doc.LoadXml(xml)

        ' Correct namespace (this is CRITICAL)
        Dim ns As New XmlNamespaceManager(doc.NameTable)
        ns.AddNamespace("dps", "http://www.sped.fazenda.gov.br/nfse")
        ns.AddNamespace("ds", "http://www.w3.org/2000/09/xmldsig#")

        ' Locate infDPS
        Dim infNode As XmlElement =
        CType(doc.SelectSingleNode("//dps:infDPS", ns), XmlElement)

        If infNode Is Nothing Then
            Throw New Exception("infDPS não encontrado para assinatura.")
        End If

        ' Read Id from XML
        Dim idAttr As XmlAttribute = infNode.GetAttributeNode("Id")
        If idAttr Is Nothing OrElse String.IsNullOrWhiteSpace(idAttr.Value) Then
            Throw New Exception("Atributo Id de infDPS ausente.")
        End If

        Dim idValue As String = idAttr.Value

        ' Tell .NET that this attribute is an XML ID
        infNode.SetAttribute("Id", idValue) ' ensures attribute exists
        doc.DocumentElement.SetAttribute("xmlns:ds", ns.LookupNamespace("ds"))

        ' Create SignedXml
        Dim signedXml As New SignedXml(doc)
        signedXml.SigningKey = cert.GetRSAPrivateKey()

        ' Canonicalization (required by NFSe)
        signedXml.SignedInfo.CanonicalizationMethod =
        SignedXml.XmlDsigC14NTransformUrl

        ' Reference the infDPS by Id
        Dim reference As New Reference()
        reference.Uri = "#" & idValue
        reference.DigestMethod = SignedXml.XmlDsigSHA256Url

        ' Enveloped + C14N transforms
        reference.AddTransform(New XmlDsigEnvelopedSignatureTransform())
        reference.AddTransform(New XmlDsigC14NTransform())

        signedXml.AddReference(reference)

        ' KeyInfo
        Dim keyInfo As New KeyInfo()
        keyInfo.AddClause(New KeyInfoX509Data(cert))
        signedXml.KeyInfo = keyInfo

        ' Compute signature
        signedXml.ComputeSignature()

        ' Append Signature to root <DPS>
        Dim signatureElement As XmlElement = signedXml.GetXml()
        doc.DocumentElement.AppendChild(doc.ImportNode(signatureElement, True))

        Return doc.OuterXml
    End Function

    Public Shared Sub ValidateAgainstXsd(xml As String, xsdPath As String)
        Dim settings As New XmlReaderSettings()
        settings.ValidationType = ValidationType.Schema
        settings.Schemas.Add(Nothing, xsdPath)
        AddHandler settings.ValidationEventHandler, Sub(sender, e)
                                                       Throw New Exception("XSD validation error: " & e.Message)
                                                   End Sub
        Using sr As New StringReader(xml)
            Using xr = XmlReader.Create(sr, settings)
                While xr.Read()
                End While
            End Using
        End Using
    End Sub
End Class
