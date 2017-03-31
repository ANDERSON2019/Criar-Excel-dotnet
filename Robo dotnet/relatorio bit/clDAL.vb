Imports System.Data.SqlClient
Imports System.Data.Common
Imports System.Text

Public Class ParameterBuilder
    Implements IDisposable

    Private _ParamName As String
    Private _Value As Object
    Private _DbType As DbType = Data.DbType.String
    Private _Direction As ParameterDirection = ParameterDirection.Input
    Private _Size As Integer = 0
    Private _SourceColumn As String = String.Empty
    Private _SourceVersion As DataRowVersion = DataRowVersion.Current
    Private _SourceColumnNullMapping As Boolean = False

    Public Enum TipoExecucao
        ExecutaQuery = 0
        RetornaParam = 1
        RetornaTabela = 2
    End Enum


    Public Sub New(ByVal ParameterName As String,
                 ByVal dbType As DbType,
                 ByVal Value As Object,
                 ByVal Direction As ParameterDirection,
                 ByVal Size As Integer)
        Me._ParamName = ParameterName
        Me._DbType = dbType
        Me._Value = Value
        Me._Direction = Direction
        Me._Size = Size
    End Sub


    Public Property ParamName() As String
        Get
            Return Me._ParamName
        End Get
        Set(ByVal value As String)
            Me._ParamName = value
        End Set
    End Property

    Public Property Value() As Object
        Get
            Return _Value
        End Get
        Set(ByVal value As Object)
            _Value = value
        End Set
    End Property

    Public Property DbType() As DbType
        Get
            Return _DbType
        End Get
        Set(ByVal value As DbType)
            _DbType = value
        End Set
    End Property

    Public Property Direction() As ParameterDirection
        Get
            Return _Direction
        End Get
        Set(ByVal value As ParameterDirection)
            _Direction = value
        End Set
    End Property

    Public Property Size() As Integer
        Get
            Return Me._Size
        End Get
        Set(ByVal value As Integer)
            Me._Size = value
        End Set
    End Property


#Region "IDisposable Support"
    Private disposedValue As Boolean

    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                _ParamName = Nothing
                _Value = Nothing
                _DbType = Nothing
                _Direction = Nothing
                _Size = Nothing
                _SourceColumn = Nothing
                _SourceVersion = Nothing
                _SourceColumnNullMapping = Nothing
            End If
        End If
        Me.disposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

Public Class clDALSQL
    Implements IDisposable

    Dim Conn As New SqlClient.SqlConnection
    Dim Command As New SqlClient.SqlCommand
    Dim Adapt As New SqlClient.SqlDataAdapter
    Dim Dt As DataTable
    'Public Parametros As New Microsoft.VisualBasic.Collection From {SqlParameter}
    Public Parametros As New List(Of SqlParameter)

    Public Enum AmbienteExecucao As Integer
        Desenvolvimento = 0
        Producao = 1
        Homologacao = 2
        DW = 3
    End Enum

    Public Enum InstanciaExecucao As Integer
        Normal = 0
        Alog = 1
    End Enum



    Private mAmbienteExecucao As AmbienteExecucao = AmbienteExecucao.Producao

    Public Property Ambiente() As AmbienteExecucao
        Get
            Return mAmbienteExecucao
        End Get
        Set(ByVal value As AmbienteExecucao)
            mAmbienteExecucao = value
        End Set
    End Property

    Private Function RetornaStrConexao(ByVal NomeBanco As String, Optional ByVal Instancia As InstanciaExecucao = InstanciaExecucao.Normal) As String
        RetornaStrConexao = ""
        If mAmbienteExecucao = AmbienteExecucao.Producao Then
            If Instancia = InstanciaExecucao.Normal Then
                RetornaStrConexao = "Password=SENHA;Persist Security Info=True;User ID=lOGIN;Initial Catalog=" & NomeBanco & ";Data Source=IP"
            ElseIf Instancia = InstanciaExecucao.Alog Then
                RetornaStrConexao = "Password=SENHA;Persist Security Info=True;User ID=lOGIN;Initial Catalog=" & NomeBanco & ";Data Source=IP"
            End If
        ElseIf mAmbienteExecucao = AmbienteExecucao.DW Then
            RetornaStrConexao = "Password=SENHA;Persist Security Info=True;User ID=lOGIN;Initial Catalog=" & NomeBanco & ";Data Source=IP"
        Else
            RetornaStrConexao = "Password=SENHA;Persist Security Info=True;User ID=lOGIN;Initial Catalog=" & NomeBanco & ";Data Source=IP"
        End If

    End Function

    Private Function RetornaStrConexao_PostGre(ByVal NomeBanco As String) As String
        If mAmbienteExecucao = AmbienteExecucao.Producao Then
            RetornaStrConexao_PostGre = "Server=IP;Port=Port;UserId=integracao;Password=;Database=" & NomeBanco
        Else
            RetornaStrConexao_PostGre = "Server=IP;Port=Port;UserId=integracao;Password=;Database=" & NomeBanco
        End If
    End Function

    Public Sub ClearParameters()
        Parametros.Clear()
    End Sub

    Public Sub ExecutaProcedure(ByVal NomeProcedure As String, ByVal NomeBanco As String, Optional ByVal ExecutionTimeOut As Integer = 0, Optional ByVal Instancia As InstanciaExecucao = InstanciaExecucao.Normal)
        Dim obParam As SqlParameter

        If Conn.State = ConnectionState.Open Then
            Conn.Close()
        End If
        Conn.ConnectionString = RetornaStrConexao(NomeBanco, Instancia)
        Conn.Open()
        Command.Connection = Conn
        Command.CommandText = NomeProcedure
        Command.Parameters.Clear()
        Command.CommandType = CommandType.StoredProcedure
        For Each obParam In Parametros
            Command.Parameters.Add(obParam)
        Next
        If ExecutionTimeOut > 0 Then
            Command.CommandTimeout = ExecutionTimeOut
        Else
            Command.CommandTimeout = 0
        End If
        Command.ExecuteNonQuery()

        'APOS RETORNAR A CONSULTA ATUALIZAR OS PARAMETROS
        Parametros.Clear()

        For Each obParam In Command.Parameters
            Parametros.Add(obParam)
        Next

        If Conn.State = ConnectionState.Open Then
            Conn.Close()
        End If

    End Sub

    Public Sub AddParameters(ByVal Nome As String, ByVal Valor As Object, ByVal Type As SqlDbType, Optional ByVal Direcao As ParameterDirection = ParameterDirection.Input, Optional ByVal Tamanho As Integer = 0)
        Dim obParam As SqlParameter
        obParam = New SqlParameter(Nome, Valor)
        obParam.Direction = Direcao
        If Tamanho > 0 Then
            obParam.Size = Tamanho
        ElseIf Tamanho = -1 Then
            obParam.DbType = SqlDbType.VarChar
            obParam.Size = -1
        Else
            obParam.Size = 0
        End If
        Parametros.Add(obParam)
    End Sub

    Public Function RetornaTabela(ByVal NomeProcedure As String, ByVal NomeBanco As String, Optional ByVal ExecutionTimeOut As Integer = 0, Optional ByVal Instancia As InstanciaExecucao = InstanciaExecucao.Normal) As DataTable
        Dim obDT As New DataTable
        Try

            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.ConnectionString = RetornaStrConexao(NomeBanco, Instancia)
            Conn.Open()
            Command.Connection = Conn
            Command.CommandText = NomeProcedure
            Command.Parameters.Clear()
            Command.CommandType = CommandType.StoredProcedure
            For Each obParam In Parametros
                Command.Parameters.Add(obParam)
            Next
            If ExecutionTimeOut > 0 Then
                Command.CommandTimeout = ExecutionTimeOut
            Else
                Command.CommandTimeout = 0
            End If
            Adapt.SelectCommand = Command

            Adapt.Fill(obDT)
        Catch ex As Exception


        Finally
            RetornaTabela = obDT
        End Try


    End Function


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.Dispose()
            Command.Dispose()
            Adapt.Dispose()
            Parametros.Clear()
            Parametros = Nothing
            Dt = Nothing
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class



