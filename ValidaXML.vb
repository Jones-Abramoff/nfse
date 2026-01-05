Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.Schema


Public Class ClassValidaXML

    Dim gsMsg As String
    Dim glErro As Long

    Public Function validaXML(ByVal _arquivo As String, ByVal _schema As String, ByVal lLote As Long, ByVal lNumIntNF As Long, ByVal db2 As SGEDadosDataContext, ByVal iFilialEmpresa As Integer) As Long

        ' Create a new validating reader

        Dim reader As XmlValidatingReader = New XmlValidatingReader(New XmlTextReader(New StreamReader(_arquivo)))
        'Dim reader As XmlValidatingReader = New XmlValidatingReader()
        '        Dim reader1 As XmlWriter
        '       reader1.
        Dim schema(2) As System.Xml.Schema.XmlSchema
        Dim iResult As Integer

        '// Create a schema collection, add the xsd to it

        Dim schemaCollection As XmlSchemaSet = New XmlSchemaSet()
        Dim iLinha As Integer

        Try

            glErro = 0

            'schemaCollection.Add("http://www.abrasf.org.br/nfse.xsd", _schema)
            '     schemaCollection.Add("http://www.ginfes.com.br/servico_consultar_lote_rps_resposta_v03.xsd", _schema)
            schemaCollection.Add("http://www.ginfes.com.br/servico_enviar_lote_rps_envio_v03.xsd", _schema)


            schemaCollection.CopyTo(schema, 0)

            '// Add the schema collection to the XmlValidatingReader

            reader.Schemas.Add(schema(0))

            '       Console.Write("Início da validação...\n")

            '    // Wire up the call back.  The ValidationEvent is fired when the
            '    // XmlValidatingReader hits an issue validating a section of the xml

            '            reader. += new ValidationEventHandler(reader_ValidationEventHandler);
            AddHandler reader.ValidationEventHandler, AddressOf reader_ValidationEventHandler

            '            // Iterate through the xml document



            '            while (reader.Read()) {}

            iLinha = 0

            While reader.Read()

                iLinha = iLinha + 1
                If Len(Trim(gsMsg)) > 0 Then

                    iResult = db2.ExecuteCommand("INSERT INTO NFeFedLoteLog ( FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5})", _
                    iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Replace(Left(" Linha = " & iLinha & "  " & gsMsg, 255), "'", "*"), lNumIntNF)

                    Form1.Msg.Items.Add(" Linha = " & iLinha & gsMsg)

                    gsMsg = ""

                End If

            End While


        Catch ex As Exception

            Dim sMsg As String

            If ex.InnerException Is Nothing Then
                sMsg = ""
            Else
                sMsg = " - " & ex.InnerException.Message
            End If

            iResult = db2.ExecuteCommand("INSERT INTO NFeFedLoteLog ( FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5})", _
            iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, Replace(Left(ex.Message & sMsg, 255), "'", "*"), lNumIntNF)

            iResult = db2.ExecuteCommand("INSERT INTO NFeFedLoteLog ( FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5})", _
            iFilialEmpresa, lLote, Now.Date, TimeOfDay.ToOADate, "ERRO - Validação do schema", 0)

            Form1.Msg.Items.Add(ex.Message & sMsg)
            Form1.Msg.Items.Add("ERRO - Validação do schema.")

            glErro = 1

        Finally
            validaXML = glErro
        End Try
        '          Console.WriteLine("\rFim de validação\n");
        'Console.ReadLine();
    End Function

    Sub reader_ValidationEventHandler(ByVal sender As Object, ByVal e As ValidationEventArgs)

        '            // Report back error information to the console...
        '        MessageBox.Show(e.Exception.Message)
        '        Console.WriteLine("\rLinha:{0} Coluna:{1} Erro:{2} Name:[3} Valor:{4}\r", e.Exception.LinePosition, e.Exception.LineNumber, e.Exception.Message, sender.Name, sender.Value)

        gsMsg = e.Exception.Message
        glErro = 1

    End Sub

End Class
