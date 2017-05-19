#Region "Imports"

Imports System.Data
Imports System.Data.OleDb
'Imports system.windows.forms

#End Region

Public Class LibDataConnection

#Region "Variables"

    Private Shared p_ConnectionString As String

    Private p_OLE_DataAdapter As New OleDbDataAdapter
    Private p_OLE_Connection As New OleDbConnection()
    Private p_OLE_Command As OleDbCommand

#End Region

#Region "Properties"
    Public Property ConnectionString() As String
        Get
            ConnectionString = p_ConnectionString
        End Get
        Set(ByVal value As String)
            p_ConnectionString = value
        End Set
    End Property

    Public Property OLE_DataAdapter() As OleDbDataAdapter
        Get
            OLE_DataAdapter = p_OLE_DataAdapter
        End Get
        Set(ByVal value As OleDbDataAdapter)
            p_OLE_DataAdapter = value
        End Set
    End Property

    Public Property OLE_Connection() As OleDbConnection
        Get
            OLE_Connection = p_OLE_Connection
        End Get
        Set(ByVal value As OleDbConnection)
            p_OLE_Connection = value
        End Set
    End Property

    Public Property OLE_Command() As OleDbCommand
        Get
            OLE_Command = p_OLE_Command
        End Get
        Set(ByVal value As OleDbCommand)
            p_OLE_Command = value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        If p_ConnectionString <> "" Then        'Open a connection
            p_OLE_Connection = New OleDbConnection(p_ConnectionString)
            Try
                p_OLE_Connection.Open()
            Catch ex As Exception
                'messsagebox.show("Error connecting to database: " & vbnewline & vbnewline & ex.message, "Database Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try

        End If
    End Sub

    Public Sub New(useNewConStr As Boolean, thisConnStr As String)
        If useNewConStr Then
            p_ConnectionString = thisConnStr
        End If
        p_OLE_Connection = New OleDbConnection(p_ConnectionString)
        Try
            p_OLE_Connection.Open()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub New(ByVal thisSQL As String)
        Me.new()
        p_OLE_Command = New OleDbCommand(thisSQL, p_OLE_Connection)
    End Sub

    Public Sub New(ByVal prepareCommand As Boolean)
        Me.New()
        p_OLE_Command = New OleDbCommand
        p_OLE_Command.Connection = p_OLE_Connection
    End Sub
#End Region

#Region "Public Methods"

    'Test Connection (open and close)
    '
    Public Function TestConnection() As Boolean
        Dim flg As Boolean = False
        Try
            flg = OpenConnection()
            CloseConnection()
            Return flg
        Catch ex As Exception
            Return False
        End Try
    End Function

    'Open new Connection
    '
    Public Function OpenConnection(Optional ByRef thisConn As OleDbConnection = Nothing) As Boolean
        If thisConn Is Nothing Then
            If p_OLE_Connection Is Nothing Then
                p_OLE_Connection = New OleDbConnection(p_ConnectionString)
            Else
                If p_OLE_Connection.ConnectionString = "" Then p_OLE_Connection = New OleDbConnection(p_ConnectionString)
            End If
            thisConn = p_OLE_Connection
        Else
            p_OLE_Connection = thisConn
        End If

        Try
            If p_OLE_Connection.State = ConnectionState.Closed Then
                p_OLE_Connection.Open()
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub CloseConnection(ByRef thisConn As OleDbConnection)
        thisConn.Close()
        thisConn.Dispose()
        thisConn = Nothing
    End Sub

    Public Sub CloseConnection()
        p_OLE_Connection.Close()
        p_OLE_Connection.Dispose()
        p_OLE_Connection = Nothing
    End Sub
#End Region

End Class
