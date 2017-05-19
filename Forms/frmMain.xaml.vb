Imports System.Data.SqlClient
Imports System.Data
Imports System.Printing
Imports System.IO.Packaging
Imports System.IO
Imports System.Windows.Xps.Packaging
Imports System.Windows.Xps.Serialization
Imports System.Data.Sql

Class frmMain
    Private cnn1 As SqlConnection
    Private cnn2 As SqlConnection
    Private AnalyseTables As Dictionary(Of String, AnalyseTable)

    Private Sub btnAnalyse_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnAnalyse.Click
        Try
            Dim cnnString As String = "Initial Catalog=" & Me.cboSQLDBList1.Text & ";Connect Timeout=60;Data Source=" & Me.cboSvrName1.Text & ";Trusted_Connection=yes;"
            cnn1 = New SqlConnection(cnnString)
            cnn1.Open()
        Catch ex As Exception
            MessageBox.Show("Connection to " & Me.cboSQLDBList1.Text & " Failed - " & ex.Message, "SQL Connection", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End Try

        Try
            Dim cnnString As String = "Initial Catalog=" & Me.cboSQLDBList2.Text & ";Connect Timeout=60;Data Source=" & Me.cboSvrName2.Text & ";Trusted_Connection=yes;"
            cnn2 = New SqlConnection(cnnString)
            cnn2.Open()
        Catch ex As Exception
            MessageBox.Show("Connection to " & Me.cboSQLDBList2.Text & " Failed - " & ex.Message, "SQL Connection", MessageBoxButton.OK, MessageBoxImage.Error)
            cnn1.Close()
            Exit Sub
        End Try
        Analyse()
        Try
            cnn1.Close()
            cnn2.Close()
        Catch ex As Exception
            MessageBox.Show("Connection to " & Me.cboSQLDBList2.Text & " Failed - " & ex.Message, "SQL Connection", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try

    End Sub

#Region "Analyse Databases"
    Private Sub Analyse()

        AnalyseTables = New Dictionary(Of String, AnalyseTable)
        CheckTables()       ' Look for diffences in the tables defined in each database
        CheckColumns()      ' Look for differences in each tables columns
        GenerateReport()    ' Generate analysis report and display

    End Sub

    Private Sub CheckTables()
 
        Dim ds1 As New DataSet
        Dim ds2 As New DataSet


        '   Get database 1's table list


        Try
            Dim cmd1 As New SqlCommand("select id,Name  from sysobjects where xType='U'", cnn1)
            Dim da1 As New SqlDataAdapter(cmd1)
            da1.Fill(ds1)
        Catch ex As Exception
            MessageBox.Show("Reading tables from Database 1 Failed - " & ex.Message, "SQL Connection", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End Try


        '   Get database 2's table list


        Try
            Dim cmd2 As New SqlCommand("select id,Name  from sysobjects where xType='U'", cnn2)
            Dim da2 As New SqlDataAdapter(cmd2)
            da2.Fill(ds2)
        Catch ex As Exception
            MessageBox.Show("Reading tables from Database 2 Failed - " & ex.Message, "SQL Connection", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End Try


        '   For each table in database 1 check if it exists in database 2
        '   and add the result to the tables collection


        For Each dr1 As DataRow In ds1.Tables(0).Rows
            Dim ExistsInDatabase2 As Boolean = False
            Dim Database2ID As Integer = 0
            For Each dr2 As DataRow In ds2.Tables(0).Rows
                If dr2("Name") = dr1("Name") Then
                    ExistsInDatabase2 = True
                    Database2ID = dr2("ID")
                    Exit For
                End If
            Next
            AnalyseTables.Add(dr1("Name"), New AnalyseTable(dr1("ID"), Database2ID, True, ExistsInDatabase2))
        Next


        '   For each table in database 2 check if it exists in the tables collection
        '   If it doesn't we need to add an item for this table to the tables collection


        For Each dr2 As DataRow In ds2.Tables(0).Rows
            If AnalyseTables.ContainsKey(dr2("Name")) = False Then
                AnalyseTables.Add(dr2("Name"), New AnalyseTable(0, dr2("ID"), False, True))
            End If
        Next
    End Sub
    Private Sub CheckColumns()

        '   If the table exists in both databases we need to compare the fields for each


        '   Pass through each table

        For Each TableName As String In AnalyseTables.Keys

            '   Look to see if the table exists in both databases and if so

            Dim ds1 As New DataSet
            Dim ds2 As New DataSet
            If AnalyseTables(TableName).ExistsInDatabase1 = True And AnalyseTables(TableName).ExistsInDatabase2 = True Then

                '   Get list of columns for the table from database 1

                Try
                    Dim cmd1 As New SqlCommand("select name,xtype,length from syscolumns where id=" & AnalyseTables(TableName).Database1ID, cnn1)
                    Dim da1 As New SqlDataAdapter(cmd1)
                    da1.Fill(ds1)
                Catch ex As Exception
                    MessageBox.Show("Reading table columns from " & Me.cboSQLDBList2.Text & " Failed - " & ex.Message, "SQL Connection", MessageBoxButton.OK, MessageBoxImage.Error)
                    Exit Sub
                End Try


                '    Get list of columns for table from database 2

                Try
                    Dim cmd2 As New SqlCommand("select name,xtype,length from syscolumns where id=" & AnalyseTables(TableName).Database2ID, cnn2)
                    Dim da2 As New SqlDataAdapter(cmd2)
                    da2.Fill(ds2)
                Catch ex As Exception
                    MessageBox.Show("Reading table columns from " & Me.cboSQLDBList2.Text & " Failed - " & ex.Message, "SQL Connection", MessageBoxButton.OK, MessageBoxImage.Error)
                    Exit Sub
                End Try


                '   For each column in database1 table check if it exists in database 2 tables
                '   and add the result to the tables columns collection collection

                For Each dr1 As DataRow In ds1.Tables(0).Rows
                    Dim Difference As String = ""
                    Dim ExistsInDatabase2 As Boolean = False
                    For Each dr2 As DataRow In ds2.Tables(0).Rows
                        If dr2("Name") = dr1("Name") Then
                            If dr2("xtype") <> dr1("xtype") Then
                                Difference = "Type is Different -  " & Me.cboSQLDBList1.Text & "  has type of " & dr1("xtype") & " " & Me.cboSQLDBList2.Text & " has type of " & dr2("xtype")
                            End If
                            If dr2("length") <> dr1("length") Then
                                Difference = "Length is Different -  " & Me.cboSQLDBList1.Text & "  has length of " & dr1("length") & " " & Me.cboSQLDBList2.Text & " has length of " & dr2("length")
                            End If
                            ExistsInDatabase2 = True
                            Exit For
                        End If
                    Next
                    If ExistsInDatabase2 = False Then
                        Difference = "Does not exists in " & Me.cboSQLDBList2.Text & " "
                    End If
                    If Difference <> "" Then
                        AnalyseTables(TableName).AnalyseColumns.Add(dr1("Name"), New AnalyseColumn(Difference))
                    End If
                Next

                '   For each column in database2 table check if it exists in database 1 table
                '   If it doesn't we need to add it to the tables columns collection

                For Each dr2 As DataRow In ds2.Tables(0).Rows
                    Dim ExistsInDatabase1 As Boolean = False
                    For Each dr1 As DataRow In ds1.Tables(0).Rows
                        If dr2("Name") = dr1("Name") Then
                            ExistsInDatabase1 = True
                            Exit For
                        End If
                    Next
                    If ExistsInDatabase1 = False Then
                        AnalyseTables(TableName).AnalyseColumns.Add(dr2("Name"), New AnalyseColumn("Does not exist in  " & Me.cboSQLDBList1.Text & " "))
                    End If
                Next
            End If
        Next
    End Sub

    Private Sub GenerateReport()
        '
        '   Produce a Flow Document containing info on the differences found
        '

        Dim MemStream As New System.IO.MemoryStream
        Dim xpsPackage As Package = Package.Open(MemStream, FileMode.CreateNew)
        Dim FlowDocument As New FlowDocument
        Dim Section As New Section
        Dim Para As Paragraph

        '
        '   Show the databases we have analysed
        '
        Para = New Paragraph
        Section.Blocks.Add(Para)
        Para.FontSize = 12
        Para.Inlines.Add("Database Compare results.")

        Para = New Paragraph
        Section.Blocks.Add(Para)
        Para.FontSize = 12
        Para.Inlines.Add("Database 1:")

        Dim Table As Table
        Dim currentRow As TableRow

        Table = New Table
        Table.Columns.Add(New TableColumn)
        Table.Columns.Add(New TableColumn)
        Table.Columns(0).Width = New GridLength(50)
        Table.FontSize = 10
        Table.RowGroups.Add(New TableRowGroup())

        currentRow = New TableRow()
        Table.RowGroups(0).Rows.Add(currentRow)
        currentRow.Cells.Add(New TableCell(New Paragraph(New Run("Server"))))
        currentRow.Cells.Add(New TableCell(New Paragraph(New Run(Me.cboSvrName1.Text))))

        currentRow = New TableRow()
        Table.RowGroups(0).Rows.Add(currentRow)
        currentRow.Cells.Add(New TableCell(New Paragraph(New Run("Database"))))
        currentRow.Cells.Add(New TableCell(New Paragraph(New Run(Me.cboSQLDBList1.Text))))

        Section.Blocks.Add(Table)

        Para = New Paragraph
        Section.Blocks.Add(Para)
        Para.FontSize = 12
        Para.Inlines.Add("Database 2:")
        Table = New Table
        Table.Columns.Add(New TableColumn)
        Table.Columns.Add(New TableColumn)
        Table.Columns(0).Width = New GridLength(50)
        Table.FontSize = 10
        Table.RowGroups.Add(New TableRowGroup())

        currentRow = New TableRow()
        Table.RowGroups(0).Rows.Add(currentRow)
        currentRow.Cells.Add(New TableCell(New Paragraph(New Run("Server"))))
        currentRow.Cells.Add(New TableCell(New Paragraph(New Run(Me.cboSvrName2.Text))))

        currentRow = New TableRow()
        Table.RowGroups(0).Rows.Add(currentRow)
        currentRow.Cells.Add(New TableCell(New Paragraph(New Run("Database"))))
        currentRow.Cells.Add(New TableCell(New Paragraph(New Run(Me.cboSQLDBList2.Text))))

        Section.Blocks.Add(Table)

        Para = New Paragraph
        Section.Blocks.Add(Para)
        Para.FontSize = 12
        Para.Inlines.Add("The following tables produced differences")

        '
        '   Pass through the table collection and print details of the differences
        '
        Dim ChangesExists As Boolean = False
        For Each TableName As String In AnalyseTables.Keys
            Dim AnalyseTable As AnalyseTable = AnalyseTables(TableName)
            If AnalyseTable.ExistsInDatabase1 <> AnalyseTable.ExistsInDatabase2 Or AnalyseTable.AnalyseColumns.Count Then
                ChangesExists = True
                Para = New Paragraph
                Section.Blocks.Add(Para)
                Para.FontSize = 14
                Para.Inlines.Add(TableName)

                If AnalyseTable.ExistsInDatabase1 = False Then
                    Para = New Paragraph
                    Para.FontSize = 10
                    Para.Foreground = Brushes.DarkBlue
                    Section.Blocks.Add(Para)
                    Para.Inlines.Add("    " & "This table does not exits in " & Me.cboSQLDBList1.Text)
                End If
                If AnalyseTable.ExistsInDatabase2 = False Then
                    Para = New Paragraph
                    Para.FontSize = 10
                    Para.Foreground = Brushes.DarkBlue
                    Section.Blocks.Add(Para)
                    Para.Inlines.Add("    " & "This table does not exits in " & Me.cboSQLDBList1.Text)
                End If
                For Each ColumnName As String In AnalyseTable.AnalyseColumns.Keys
                    Para = New Paragraph
                    Section.Blocks.Add(Para)
                    Para.FontSize = 10
                    Para.Foreground = Brushes.Maroon
                    Para.Inlines.Add("    " & ColumnName & " " & AnalyseTable.AnalyseColumns(ColumnName).Difference)
                Next
            End If
        Next

        If ChangesExists = False Then
            Para = New Paragraph
            Section.Blocks.Add(Para)
            Para.FontSize = 12
            Para.Inlines.Add("No differences found")
        End If

        FlowDocument.Blocks.Add(Section)

        '
        '   Convert Flowdocument to Fixed Page
        '
        Dim xpsDocument As New XpsDocument(xpsPackage, CompressionOption.Maximum)
        Dim paginator As DocumentPaginator = CType(FlowDocument, IDocumentPaginatorSource).DocumentPaginator
        Dim rsm As New XpsSerializationManager(New XpsPackagingPolicy(xpsDocument), False)
        paginator = New DocumentPaginatorWrapper(paginator, New Size(768, 1056), New Size(48, 48))
        rsm.SaveAsXaml(paginator)
        xpsDocument.Close()
        xpsPackage.Close()
        Dim DisplayReport As New DisplayReport
        DisplayReport.OpenStream(MemStream)
        DisplayReport.Show()

    End Sub

#End Region
#Region "Private Classes"

    Private Class AnalyseTable
        Friend ExistsInDatabase1 As Boolean
        Friend ExistsInDatabase2 As Boolean
        Friend Database1ID As Integer
        Friend Database2ID As Integer
        Friend AnalyseColumns As New Dictionary(Of String, AnalyseColumn)

        Friend Sub New(ByVal Database1ID As Integer, ByVal Database2ID As Integer, ByVal ExistsInDatabase1 As Boolean, ByVal ExistsInDatabase2 As Boolean)
            Me.Database1ID = Database1ID
            Me.Database2ID = Database2ID
            Me.ExistsInDatabase1 = ExistsInDatabase1
            Me.ExistsInDatabase2 = ExistsInDatabase2
        End Sub
    End Class
    Private Class AnalyseColumn
        Friend Difference As String

        Friend Sub New(ByVal Difference As String)
            Me.Difference = Difference
        End Sub
    End Class

#End Region

    Private Sub btnFind1_Click(sender As Object, e As RoutedEventArgs) Handles btnFind1.Click
        findSQLInstances(1)
        findSQLInstances(2)
    End Sub

    Private Sub findSQLInstances(ByVal type As Integer) 'this bit takes FOREVER
        Me.Cursor = Cursors.Wait
        Dim instance As SqlDataSourceEnumerator = SqlDataSourceEnumerator.Instance
        Dim table As System.Data.DataTable = instance.GetDataSources()
        sqlServerList_Load(table, type)        ' Display the contents of the table.
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub sqlServerList_Load(ByVal table As DataTable, ByVal type As Integer)
        Dim thisServer As String

        If type = 1 Then
            Me.cboSvrName1.Items.Clear()
        Else
            Me.cboSvrName2.Items.Clear()
        End If
        For Each row As DataRow In table.Rows
            thisServer = ""
            For Each col As DataColumn In table.Columns
                If col.ColumnName = "ServerName" Then thisServer = row(col).trim
                If col.ColumnName = "InstanceName" Then
                    If Not row(col) Is DBNull.Value Then
                        If row(col).trim <> "" Then
                            thisServer &= "\" & row(col).trim
                        End If
                    End If
                End If
            Next

            If type = 1 Then
                Me.cboSvrName1.Items.Add(thisServer)
            Else
                Me.cboSvrName2.Items.Add(thisServer)
            End If
        Next

    End Sub

    Private Sub cboSvrName1_LostFocus(sender As Object, e As System.EventArgs) Handles cboSvrName1.LostFocus
        loadDatabaseList(1)
    End Sub

    Private Sub cboSvrName2_LostFocus(sender As Object, e As System.EventArgs) Handles cboSvrName2.LostFocus
        loadDatabaseList(2)
    End Sub

    Private Sub loadDatabaseList(ByVal type As Integer)
        Dim loadSvrDetails As Boolean
        Me.Cursor = Cursors.Wait
        gSQLServerDetails.ClearClass()
        If type = 1 Then
            Me.cboSQLDBList1.Items.Clear()
            If Me.cboSvrName1.Text <> "" Then
                loadSvrDetails = gSQLServerDetails.InitialiseMe(Me.cboSvrName1.Text)
                If loadSvrDetails Then
                    ' load databases...
                    gSQLServerDetails.DatabaseList_Get(False)
                    Me.cboSQLDBList1.Items.Clear()
                    Me.cboSQLDBList1.Items.Add("Master")
                    For i As Integer = 1 To gSQLServerDetails.DatabaseCollection.Count
                        Me.cboSQLDBList1.Items.Add(DirectCast(gSQLServerDetails.DatabaseCollection(i), clsDatabaseDetails).DatabaseName.Trim)
                    Next
                End If
            End If
        Else
            Me.cboSQLDBList2.Items.Clear()
            If Me.cboSvrName2.Text <> "" Then
                loadSvrDetails = gSQLServerDetails.InitialiseMe(Me.cboSvrName2.Text)
                If loadSvrDetails Then
                    ' load databases...
                    gSQLServerDetails.DatabaseList_Get(False)
                    Me.cboSQLDBList2.Items.Clear()
                    Me.cboSQLDBList2.Items.Add("Master")
                    For i As Integer = 1 To gSQLServerDetails.DatabaseCollection.Count
                        Me.cboSQLDBList2.Items.Add(DirectCast(gSQLServerDetails.DatabaseCollection(i), clsDatabaseDetails).DatabaseName.Trim)
                    Next
                End If
            End If
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class
