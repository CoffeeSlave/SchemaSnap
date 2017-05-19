Imports System.Data

Public Class clsSQLServerDetails

#Region "declarations"
    Private p_ServerProductVersion As String
    Private p_ServerVersion As String
    Private p_ServerEdition As String
    Private p_ServerName As String
    Private p_ServerUserName As String
    Private p_ServerPassword As String

    Private p_ConnectionString As String
    Private p_ServerDetailsDictionary As New Dictionary(Of String, String)

    Private p_DatabaseName As String
    Private p_DatabaseCollection As New Collection
    'Private p_DatabaseDetails As New clsDatabaseDetails

#End Region

#Region "properties"

    Public Property ServerProductVersion As String
        Get
            Return p_ServerProductVersion
        End Get
        Set(value As String)
            p_ServerProductVersion = value
        End Set
    End Property

    Public Property ServerVersion As String
        Get
            Return p_ServerVersion
        End Get
        Set(value As String)
            p_ServerVersion = value
        End Set
    End Property

    Public Property ServerEdition As String
        Get
            Return p_ServerEdition
        End Get
        Set(value As String)
            p_ServerEdition = value
        End Set
    End Property

    Public Property ServerName As String
        Get
            Return p_ServerName
        End Get
        Set(value As String)
            p_ServerName = value
        End Set
    End Property

    Public Property ServerUserName As String
        Get
            Return p_ServerUserName
        End Get
        Set(value As String)
            p_ServerUserName = value
        End Set
    End Property

    Public Property ServerPassword As String
        Get
            Return p_ServerPassword
        End Get
        Set(value As String)
            p_ServerPassword = value
        End Set
    End Property

    Public Property ConnectionString As String
        Get
            Return p_ConnectionString
        End Get
        Set(value As String)
            p_ConnectionString = value
        End Set
    End Property

    Public Property DatabaseCollection As Collection
        Get
            If p_DatabaseCollection Is Nothing Then p_DatabaseCollection = New Collection
            Return p_DatabaseCollection
        End Get
        Set(value As Collection)
            p_DatabaseCollection = value
        End Set
    End Property

    'Public Property DatabaseDetails As clsDatabaseDetails
    '    Get
    '        If p_DatabaseDetails Is Nothing Then p_DatabaseDetails = New clsDatabaseDetails
    '        Return p_DatabaseDetails
    '    End Get
    '    Set(value As clsDatabaseDetails)
    '        p_DatabaseDetails = value
    '    End Set
    'End Property

    Public Property ServerDetails As Dictionary(Of String, String)
        Get
            If p_ServerDetailsDictionary Is Nothing Then p_ServerDetailsDictionary = New Dictionary(Of String, String)
            Return p_ServerDetailsDictionary
        End Get
        Set(value As Dictionary(Of String, String))
            p_ServerDetailsDictionary = value
        End Set
    End Property

    Public Property DatabaseName As String
        Get
            Return p_DatabaseName
        End Get
        Set(value As String)
            p_DatabaseName = value
        End Set
    End Property

#End Region

#Region "constructors"

    Public Function InitialiseMe(thisServerName As String) As Boolean

        Return InitialiseMe(thisServerName, "", "")

    End Function

    Public Function InitialiseMe(thisSQLSrvName As String, thisUsrNme As String, thisPWrd As String) As Boolean
        Dim isValidCon As String

        ClearClass()
        p_DatabaseName = "master"
        p_ServerName = thisSQLSrvName
        p_ServerUserName = thisUsrNme
        p_ServerPassword = thisPWrd

        isValidCon = initialiseMe_SetLibCon()
        If isValidCon Then serverDetails_Initialise()

        Return isValidCon


    End Function

    Private Function initialiseMe_SetLibCon() As Boolean

        InitialiseConnStr(p_DatabaseName, p_ServerName, p_ServerUserName, p_ServerPassword)
        Dim libCon As New LibDataConnection(True, p_ConnectionString)
        'libCon.ConnectionString = p_ConnectionString

        Return libCon.TestConnection()

    End Function

    Public Sub InitialiseConnStr(thisDatabaseName As String, Optional thisServerName As String = "", Optional thisUsrNme As String = "", Optional thisPWrd As String = "")
        Dim securityTag As String = ""

        If thisServerName = "" Then thisServerName = p_ServerName
        p_ServerUserName = thisUsrNme
        p_ServerPassword = thisPWrd

        If thisUsrNme <> "" Then
            securityTag = ";Trusted_Connection=False;User Id=" & p_ServerUserName & ";Password=" & p_ServerPassword & ";"
        Else
            securityTag = ";Integrated Security=SSPI"
        End If
        p_ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & thisServerName.Trim & ";Initial Catalog=" & thisDatabaseName.Trim & securityTag

    End Sub

#End Region

#Region "methods"

    Public Sub ClearClass()

        p_ServerProductVersion = ""
        p_ServerVersion = ""
        p_ServerEdition = ""
        p_ServerName = ""

        ServerDetails = New Dictionary(Of String, String)
        p_DatabaseCollection = New Collection
        'p_DatabaseDetails = New clsDatabaseDetails

    End Sub

    Public Function ConnectionString_Test() As Boolean
        Dim libCon As New LibDataConnection()
        libCon.ConnectionString = p_ConnectionString

        ConnectionString_Test = libCon.TestConnection

    End Function

    Private Sub serverDetails_Initialise()

        p_ServerProductVersion = serverDetails_Get("SELECT serverproperty('ProductVersion') AS ProductVersion ", "ProductVersion")
        p_ServerVersion = serverDetails_Get("SELECT @@Version AS Version", "Version")
        p_ServerEdition = serverDetails_Get("SELECT convert(char(30), serverproperty('Edition')) AS Edition ", "Edition")

        p_ServerDetailsDictionary.Add("Server_Name", p_ServerName)
        p_ServerDetailsDictionary.Add("Product_Version", p_ServerProductVersion)
        p_ServerDetailsDictionary.Add("Server_Version", p_ServerVersion)
        p_ServerDetailsDictionary.Add("Server_Edition", p_ServerEdition)

        DatabaseList_Get()

    End Sub

    Public Function DatabaseList_Get(Optional showSystemDatabases As Boolean = False) As Boolean
        Dim sql As String = "SELECT [name] FROM master.dbo.sysdatabases " & IIf(Not showSystemDatabases, " WHERE name NOT IN ('master', 'tempdb', 'model', 'msdb', 'ReportServer$')", "") & " ORDER BY Name "

        Try
            p_DatabaseCollection = New Collection
            Dim dReader As OleDb.OleDbDataReader
            Dim libCon As New LibDataConnection
            Dim idx As Integer = -1
            libCon.ConnectionString = p_ConnectionString
            Dim cn As New LibDataConnection(sql)
            dReader = cn.OLE_Command.ExecuteReader
            If dReader.HasRows Then
                While dReader.Read
                    If idx < 0 Then idx = getIndex(dReader, "name")
                    databaseColl_AddDatabase(dReader.GetString(idx))
                End While
            End If
            dReader.Close()
            cn.CloseConnection()
        Catch
            '
        End Try

        Return True

    End Function

    Private Sub databaseColl_AddDatabase(thisDBName As String)
        Dim db As New clsDatabaseDetails

        db.DatabaseName = thisDBName
        p_DatabaseCollection.Add(db)

    End Sub

    Private Function serverDetails_Get(thisSQL As String, propertyFieldToGet As String) As String
        Dim result As String = ""

        Try
            Dim dReader As OleDb.OleDbDataReader
            Dim cn As New LibDataConnection(thisSQL)
            dReader = cn.OLE_Command.ExecuteReader
            If dReader.Read Then
                Dim idx As Integer
                idx = getIndex(dReader, propertyFieldToGet)
                If idx > -1 Then result = dReader.GetString(idx).Trim
                dReader.Close()
                cn.CloseConnection()
            End If
        Catch

        End Try

        Return result

    End Function

    Private Function getIndex(ByVal thisReader As OleDb.OleDbDataReader, ByVal thisfield As String) As Integer
        Try
            Dim idx As Integer
            If thisReader Is Nothing Then Return -1
            idx = thisReader.GetOrdinal(thisfield)
            If thisReader.IsDBNull(idx) Then
                Return -1
            Else
                Return idx
            End If
        Catch ex As Exception
            Return -1
        End Try
    End Function

#End Region

End Class
