Public Class clsDatabaseDetails

#Region "properties"

    Private p_DatabaseName As String
    Private p_FieldsCollection As New Collection
#End Region

#Region "properties"

    Public Property DatabaseName As String
        Get
            Return p_DatabaseName
        End Get
        Set(value As String)
            p_DatabaseName = value
        End Set
    End Property
#End Region

    Public Function TableList_Get() As Boolean
        Dim sql As String = "SELECT [TableName] = so.name FROM sysobjects so WHERE so.xtype = 'U' GROUP BY so.name"



        Return True

    End Function

End Class
