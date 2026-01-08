Imports System.Runtime.Serialization

<DataContract>
Public Class RetornoCriacaoNfse

    <DataMember(Name:="tipoAmbiente")>
    Public Property TipoAmbiente As Integer

    <DataMember(Name:="versaoAplicativo")>
    Public Property VersaoAplicativo As String

    <DataMember(Name:="dataHoraProcessamento")>
    Public Property DataHoraProcessamento As String

    <DataMember(Name:="idDps")>
    Public Property IdDps As String

    <DataMember(Name:="chaveAcesso")>
    Public Property ChaveAcesso As String

    <DataMember(Name:="nfseXmlGZipB64")>
    Public Property NfseXmlGZipB64 As String

    <DataMember(Name:="alertas")>
    Public Property Alertas As List(Of AlertaNfse)

End Class


<DataContract>
Public Class AlertaNfse

    <DataMember(Name:="mensagem")>
    Public Property Mensagem As String

    <DataMember(Name:="codigo")>
    Public Property Codigo As String

    <DataMember(Name:="descricao")>
    Public Property Descricao As String

    <DataMember(Name:="complemento")>
    Public Property Complemento As String

End Class
