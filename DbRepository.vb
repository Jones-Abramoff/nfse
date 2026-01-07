Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Odbc

' Repositório minimalista via ODBC (como no seu padrão atual).
' Ajuste nomes das tabelas/colunas conforme seu SGE.
Public Class DbRepository
    Private ReadOnly _connStr As String

    Public Sub New(empresa As String)
        _connStr = "DSN=SGEDados" & empresa & ";UID=sa;PWD=SAPWD"
    End Sub

    Public Function GetFilial(filialEmpresa As Integer) As FilialEmpresaRow
        Using cn As New OdbcConnection(_connStr)
            cn.Open()
            Using cmd As New OdbcCommand("SELECT TOP 1 FilialEmpresa, CGC, InscricaoMunicipal, RazaoSocial, NomeFantasia, Endereco, CertificadoA1A3, RPSAmbiente, Telefone, Email FROM FiliaisEmpresa WHERE FilialEmpresa=?", cn)
                cmd.Parameters.Add("p1", OdbcType.SmallInt).Value = filialEmpresa
                Using rd = cmd.ExecuteReader()
                    If Not rd.Read() Then Throw New Exception("FilialEmpresa não encontrada: " & filialEmpresa)
                    Return New FilialEmpresaRow With {
                        .FilialEmpresa = Convert.ToInt32(rd("FilialEmpresa")),
                        .CGC = Nz(rd("CGC")),
                        .InscricaoMunicipal = Nz(rd("InscricaoMunicipal")),
                        .RazaoSocial = Nz(rd("RazaoSocial")),
                        .NomeFantasia = Nz(rd("NomeFantasia")),
                        .EnderecoCod = Convert.ToInt32(rd("Endereco")),
                        .CertificadoA1A3 = Nz(rd("CertificadoA1A3")),
                        .RPSAmbiente = Convert.ToInt32(rd("RPSAmbiente")),
                        .Telefone = Nz(rd("Telefone")),
                        .Email = Nz(rd("Email"))
                    }
                End Using
            End Using
        End Using
    End Function

    Public Function GetEndereco(cod As Integer) As EnderecoRow
        Using cn As New OdbcConnection(_connStr)
            cn.Open()
            Using cmd As New OdbcCommand("SELECT TOP 1 Codigo, Endereco as Logradouro, Numero, Complemento, Bairro, CEP, CodMun as CidadeCodMun, UF FROM Enderecos WHERE Codigo=?", cn)
                cmd.Parameters.Add("p1", OdbcType.Int).Value = cod
                Using rd = cmd.ExecuteReader()
                    If Not rd.Read() Then Throw New Exception("Endereço não encontrado: " & cod)
                    Return New EnderecoRow With {
                        .Codigo = Convert.ToInt32(rd("Codigo")),
                        .Logradouro = Nz(rd("Logradouro")),
                        .Numero = Nz(rd("Numero")),
                        .Complemento = Nz(rd("Complemento")),
                        .Bairro = Nz(rd("Bairro")),
                        .CEP = SoNumeros(Nz(rd("CEP"))),
                        .CidadeCodMun = Convert.ToInt32(rd("CidadeCodMun")),
                        .UF = Nz(rd("UF"))
                    }
                End Using
            End Using
        End Using
    End Function

    Public Function GetNFiscal(numIntDoc As Long) As NFiscalRow
        Using cn As New OdbcConnection(_connStr)
            cn.Open()
            Using cmd As New OdbcCommand("SELECT TOP 1 NumIntDoc, FilialEmpresa, Serie, NumNotaFiscal, DataEmissao, DataCadastro, HoraEmissao, ValorServicos, ValorDeducoes, ValorDesconto, Aliquota as AliquotaISS, ValorISS, ISSRetido FROM RPSWEBProt WHERE NumIntDoc=? ORDER BY DataEmissao DESC", cn)
                cmd.Parameters.Add("p1", OdbcType.Int).Value = numIntDoc
                Using rd = cmd.ExecuteReader()
                    If rd.Read() Then
                        ' Se você usa outra tabela para os dados base, ajuste aqui.
                        Return New NFiscalRow With {
                            .NumIntDoc = Convert.ToInt64(rd("NumIntDoc")),
                            .FilialEmpresa = Convert.ToInt32(rd("FilialEmpresa")),
                            .Serie = Nz(rd("Serie")),
                            .NumNotaFiscal = If(IsDBNull(rd("NumNotaFiscal")), 0, Convert.ToInt32(rd("NumNotaFiscal"))),
                            .DataEmissao = If(IsDBNull(rd("DataEmissao")), CType(Nothing, DateTime?), Convert.ToDateTime(rd("DataEmissao"))),
                            .DataCadastro = If(IsDBNull(rd("DataCadastro")), CType(Nothing, DateTime?), Convert.ToDateTime(rd("DataCadastro"))),
                            .HoraEmissao = If(IsDBNull(rd("HoraEmissao")), CType(Nothing, Double?), Convert.ToDouble(rd("HoraEmissao"))),
                            .ValorServicos = If(IsDBNull(rd("ValorServicos")), CType(Nothing, Double?), Convert.ToDouble(rd("ValorServicos"))),
                            .ValorDeducoes = If(IsDBNull(rd("ValorDeducoes")), CType(Nothing, Double?), Convert.ToDouble(rd("ValorDeducoes"))),
                            .ValorDesconto = If(IsDBNull(rd("ValorDesconto")), CType(Nothing, Double?), Convert.ToDouble(rd("ValorDesconto"))),
                            .AliquotaISS = If(IsDBNull(rd("AliquotaISS")), CType(Nothing, Double?), Convert.ToDouble(rd("AliquotaISS"))),
                            .ValorISS = If(IsDBNull(rd("ValorISS")), CType(Nothing, Double?), Convert.ToDouble(rd("ValorISS"))),
                            .ISSRetido = If(IsDBNull(rd("ISSRetido")), CType(Nothing, Integer?), Convert.ToInt32(rd("ISSRetido")))
                        }
                    End If
                End Using
            End Using
        End Using

        ' fallback na NFiscal (caso não esteja em RPSWEBProt)
        Using cn As New OdbcConnection(_connStr)
            cn.Open()
            Using cmd As New OdbcCommand("SELECT TOP 1 NumIntDoc, FilialEmpresa, Serie, NumNotaFiscal, DataEmissao, DataCadastro, HoraEmissao, ValorTotal as ValorServicos, ValorDeducoes, ValorDesconto, 0 as AliquotaISS, 0 as ValorISS, 0 as ISSRetido FROM NFiscal WHERE NumIntDoc=?", cn)
                cmd.Parameters.Add("p1", OdbcType.Int).Value = numIntDoc
                Using rd = cmd.ExecuteReader()
                    If Not rd.Read() Then Throw New Exception("NFiscal não encontrada: " & numIntDoc)
                    Return New NFiscalRow With {
                        .NumIntDoc = Convert.ToInt64(rd("NumIntDoc")),
                        .FilialEmpresa = Convert.ToInt32(rd("FilialEmpresa")),
                        .Serie = Nz(rd("Serie")),
                        .NumNotaFiscal = If(IsDBNull(rd("NumNotaFiscal")), 0, Convert.ToInt32(rd("NumNotaFiscal"))),
                        .DataEmissao = If(IsDBNull(rd("DataEmissao")), CType(Nothing, DateTime?), Convert.ToDateTime(rd("DataEmissao"))),
                        .DataCadastro = If(IsDBNull(rd("DataCadastro")), CType(Nothing, DateTime?), Convert.ToDateTime(rd("DataCadastro"))),
                        .HoraEmissao = If(IsDBNull(rd("HoraEmissao")), CType(Nothing, Double?), Convert.ToDouble(rd("HoraEmissao"))),
                        .ValorServicos = If(IsDBNull(rd("ValorServicos")), CType(Nothing, Double?), Convert.ToDouble(rd("ValorServicos"))),
                        .ValorDeducoes = If(IsDBNull(rd("ValorDeducoes")), CType(Nothing, Double?), Convert.ToDouble(rd("ValorDeducoes"))),
                        .ValorDesconto = If(IsDBNull(rd("ValorDesconto")), CType(Nothing, Double?), Convert.ToDouble(rd("ValorDesconto"))),
                        .AliquotaISS = If(IsDBNull(rd("AliquotaISS")), CType(Nothing, Double?), Convert.ToDouble(rd("AliquotaISS"))),
                        .ValorISS = If(IsDBNull(rd("ValorISS")), CType(Nothing, Double?), Convert.ToDouble(rd("ValorISS"))),
                        .ISSRetido = If(IsDBNull(rd("ISSRetido")), CType(Nothing, Integer?), Convert.ToInt32(rd("ISSRetido")))
                    }
                End Using
            End Using
        End Using
    End Function

    Public Function GetItens(numIntNF As Long) As List(Of ItensNFiscalRow)
        Dim list As New List(Of ItensNFiscalRow)
        Using cn As New OdbcConnection(_connStr)
            cn.Open()
            Using cmd As New OdbcCommand("SELECT NumIntNF, Item, Produto, Quantidade, PrecoUnitario, ValorDesconto, DescricaoItem FROM ItensNFiscal WHERE NumIntNF=? ORDER BY Item", cn)
                cmd.Parameters.Add("p1", OdbcType.Int).Value = numIntNF
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        list.Add(New ItensNFiscalRow With {
                            .NumIntNF = Convert.ToInt64(rd("NumIntNF")),
                            .Item = Convert.ToInt32(rd("Item")),
                            .Produto = Nz(rd("Produto")),
                            .Quantidade = If(IsDBNull(rd("Quantidade")), 0R, Convert.ToDouble(rd("Quantidade"))),
                            .PrecoUnitario = If(IsDBNull(rd("PrecoUnitario")), 0R, Convert.ToDouble(rd("PrecoUnitario"))),
                            .ValorDesconto = If(IsDBNull(rd("ValorDesconto")), 0R, Convert.ToDouble(rd("ValorDesconto"))),
                            .DescricaoItem = Nz(rd("DescricaoItem"))
                        })
                    End While
                End Using
            End Using
        End Using
        Return list
    End Function

    ' Tomador: adapte para sua origem real. Aqui deixo um exemplo "placeholder".
    Public Function GetTomador(numIntDoc As Long) As TomadorRow
        ' Se você já monta o tomador no ABRASF, replique a mesma query e mapeie aqui.
        Return New TomadorRow With {
            .RazaoSocial = "TOMADOR SEM DOC",
            .Email = "",
            .Telefone = "",
            .CpfCnpj = "",
            .CNaoNIF = 1,
            .Endereco = New EnderecoRow With {
                .Logradouro = "SEM ENDEREÇO",
                .Numero = "S/N",
                .Complemento = "",
                .Bairro = "CENTRO",
                .CEP = "00000000",
                .CidadeCodMun = 3550308,
                .UF = "SP"
            }
        }
    End Function

    Public Sub GravaEmissaoOk(numIntDoc As Long,
                             filialEmpresa As Integer,
                             ambiente As Integer,
                             idDps As String,
                             chaveAcesso As String,
                             nfseXml As String,
                             protocolo As String,
                             codigoVerificacao As String)

        Using cn As New OdbcConnection(_connStr)
            cn.Open()

            ' Ajuste conforme sua tabela destino.
            Using cmd As New OdbcCommand("UPDATE RPSWEBProt SET Protocolo=?, CodigoVerificacao=?, Id=?, Numero=?, Data=?, Hora=? WHERE NumIntDoc=? AND FilialEmpresa=? AND Ambiente=?", cn)
                cmd.Parameters.Add("p1", OdbcType.VarChar, 50).Value = Nz(protocolo)
                cmd.Parameters.Add("p2", OdbcType.VarChar, 9).Value = Nz(codigoVerificacao)
                cmd.Parameters.Add("p3", OdbcType.VarChar, 255).Value = Nz(idDps)
                cmd.Parameters.Add("p4", OdbcType.VarChar, 15).Value = Nz(chaveAcesso)
                cmd.Parameters.Add("p5", OdbcType.DateTime).Value = DateTime.Now.Date
                cmd.Parameters.Add("p6", OdbcType.Double).Value = DateTime.Now.TimeOfDay.TotalDays
                cmd.Parameters.Add("p7", OdbcType.Int).Value = Convert.ToInt32(numIntDoc)
                cmd.Parameters.Add("p8", OdbcType.SmallInt).Value = filialEmpresa
                cmd.Parameters.Add("p9", OdbcType.SmallInt).Value = ambiente
                cmd.ExecuteNonQuery()
            End Using

            Using cmd2 As New OdbcCommand("INSERT INTO RPSWEBLoteLog (NumIntDoc, FilialEmpresa, Lote, Data, Hora, status, NumIntNF, Ambiente) VALUES (?,?,?,?,?,?,?,?)", cn)
                cmd2.Parameters.Add("p1", OdbcType.Int).Value = Convert.ToInt32(numIntDoc) ' ou seu NumInt sequencial de log
                cmd2.Parameters.Add("p2", OdbcType.SmallInt).Value = filialEmpresa
                cmd2.Parameters.Add("p3", OdbcType.Int).Value = 0
                cmd2.Parameters.Add("p4", OdbcType.DateTime).Value = DateTime.Now.Date
                cmd2.Parameters.Add("p5", OdbcType.Double).Value = DateTime.Now.TimeOfDay.TotalDays
                cmd2.Parameters.Add("p6", OdbcType.VarChar, 255).Value = Left("NFSe Nacional emitida. Chave=" & chaveAcesso, 255)
                cmd2.Parameters.Add("p7", OdbcType.Int).Value = Convert.ToInt32(numIntDoc)
                cmd2.Parameters.Add("p8", OdbcType.SmallInt).Value = ambiente
                cmd2.ExecuteNonQuery()
            End Using

            ' Se quiser gravar o XML da NFSe em tabela própria, inclua aqui.
        End Using
    End Sub

    Private Shared Function Nz(o As Object) As String
        If o Is Nothing OrElse IsDBNull(o) Then Return ""
        Return Convert.ToString(o)
    End Function

    Private Shared Function SoNumeros(s As String) As String
        If String.IsNullOrEmpty(s) Then Return ""
        Dim sb As New System.Text.StringBuilder()
        For Each ch In s
            If Char.IsDigit(ch) Then sb.Append(ch)
        Next
        Return sb.ToString()
    End Function
End Class
