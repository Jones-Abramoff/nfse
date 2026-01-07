Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic

' ---------------------------
' Modelos "do banco" (mínimos)
' Ajuste nomes/queries conforme seu SGE.
' ---------------------------
Public Class FilialEmpresaRow
    Public Property FilialEmpresa As Integer
    Public Property CGC As String = ""
    Public Property InscricaoMunicipal As String = ""
    Public Property RazaoSocial As String = ""
    Public Property NomeFantasia As String = ""
    Public Property EnderecoCod As Integer
    Public Property CertificadoA1A3 As String = ""
    Public Property RPSAmbiente As Integer
    Public Property Telefone As String = ""
    Public Property Email As String = ""
End Class

Public Class EnderecoRow
    Public Property Codigo As Integer
    Public Property Logradouro As String = ""
    Public Property Numero As String = ""
    Public Property Complemento As String = ""
    Public Property Bairro As String = ""
    Public Property CEP As String = ""
    Public Property CidadeCodMun As Integer
    Public Property UF As String = ""
End Class

Public Class NFiscalRow
    Public Property NumIntDoc As Long
    Public Property FilialEmpresa As Integer
    Public Property Serie As String = ""
    Public Property NumNotaFiscal As Integer
    Public Property DataEmissao As DateTime?
    Public Property DataCadastro As DateTime?
    Public Property HoraEmissao As Double?
    Public Property ValorServicos As Double?
    Public Property ValorDeducoes As Double?
    Public Property ValorDesconto As Double?
    Public Property AliquotaISS As Double?
    Public Property ValorISS As Double?
    Public Property ISSRetido As Integer?
    Public Property CodVerificacaoNFe As String = ""
End Class

Public Class ItensNFiscalRow
    Public Property NumIntNF As Long
    Public Property Item As Integer
    Public Property Produto As String = ""
    Public Property Quantidade As Double
    Public Property PrecoUnitario As Double
    Public Property ValorDesconto As Double
    Public Property DescricaoItem As String = ""
End Class

Public Class ISSQNRow
    Public Property Codigo As String = ""      ' Ex: 0702 / 7.02 etc (do seu cadastro)
    Public Property CodServNFe As Integer      ' municipal (3 dígitos no XSD)
    Public Property CNAE As String = ""
    Public Property CListServ As String = ""  ' lista LC116 ex "7.02"
End Class

' Tomador (sempre existe no seu fluxo)
Public Class TomadorRow
    Public Property RazaoSocial As String = ""
    Public Property Email As String = ""
    Public Property Telefone As String = ""
    ' Identificação: ou CPF/CNPJ ou cNaoNIF (para sem doc)
    Public Property CpfCnpj As String = "" ' 11 ou 14
    Public Property CNaoNIF As Integer = 1 ' motivo sem doc (ajuste conforme regras)
    Public Property Endereco As EnderecoRow = New EnderecoRow()
End Class

' ---------------------------
' Resultado da emissão
' ---------------------------
Public Class NfseEmissaoResult
    Public Property Sucesso As Boolean
    Public Property TipoAmbiente As Integer
    Public Property VersaoAplicativo As String = ""
    Public Property DataHoraProcessamento As DateTime?
    Public Property IdDps As String = ""
    Public Property ChaveAcesso As String = ""
    Public Property NfseXml As String = ""
    Public Property Alertas As List(Of NfseAlerta) = New List(Of NfseAlerta)()
    Public Property Erro As String = ""
End Class

Public Class NfseAlerta
    Public Property mensagem As String = ""
    Public Property codigo As String = ""
    Public Property descricao As String = ""
    Public Property complemento As String = ""
End Class
