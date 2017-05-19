Imports System
Imports System.Net
Imports System.IO
Imports System.IO.Packaging
Imports System.Printing
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Markup
Imports System.Windows.Xps
Imports System.Windows.Xps.Packaging

Partial Public Class DisplayReport
    Private m_xpsDocumentPath As String             ' XPS document path and filename.
    Private m_packageUri As Uri                     ' XPS document path and filename URI.
    Private m_xpsPackage As Package = Nothing       ' XPS document package.
    Private m_xpsDocument As XpsDocument            ' XPS document within the package.
    Private m_xpsFixedDocumentSequence As FixedDocumentSequence

    Private ReadOnly _fixedDocumentSequenceContentType As String = "application/vnd.ms-package.xps-fixeddocumentsequence+xml"

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    ''' <summary>
    ''' Prints the Document
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Print(ByVal PrintTicket As PrintTicket)
        Dim dlg As New PrintDialog()

        ' Show the printer dialog.  If the return is "true",
        ' the user made a valid selection and clicked "Ok".
        dlg.PrintTicket = PrintTicket
        dlg.PrintQueue.DefaultPrintTicket = PrintTicket
        Select Case dlg.ShowDialog()
            Case True
                Dim xpsdw As XpsDocumentWriter = PrintQueue.CreateXpsDocumentWriter(dlg.PrintQueue)
                xpsdw.Write(m_xpsFixedDocumentSequence, dlg.PrintTicket)
        End Select

    End Sub

    ''' ------------------- GetFixedDocumentSequenceUri --------------------
    ''' <summary>
    '''   Returns the part URI of first FixedDocumentSequence
    '''   contained in the package.</summary>
    ''' <returns>
    '''   The URI of first FixedDocumentSequence in the package,
    '''   or null if no FixedDocumentSequence is found.</returns>
    Private Function GetFixedDocumentSequenceUri() As Uri
        ' Iterate through the package parts
        ' to find first FixedDocumentSequence.
        For Each part As PackagePart In m_xpsPackage.GetParts()
            If part.ContentType = _fixedDocumentSequenceContentType Then
                Return part.Uri
            End If
        Next
        ' Return null if a FixedDocumentSequence isn't found.
        Return Nothing
    End Function

    ''' --------------------------- GetPackage -----------------------------
    ''' <summary>
    '''   Returns the XPS package contained within a given file.</summary>
    ''' <param name="filename">
    '''   The path and name of the file containing the package.</param>
    ''' <returns>
    '''   The package contained within the specifed file.</returns>
    Private Function GetPackage(ByVal filename As String) As Package
        Dim inputPackage As Package = Nothing

        ' "filename" should be the full path and name of the file.
        Dim webRequest As WebRequest = System.Net.WebRequest.Create(filename)
        If webRequest IsNot Nothing Then
            Dim webResponse As WebResponse = webRequest.GetResponse()
            If webResponse IsNot Nothing Then
                Dim inputPackageStream As Stream = webResponse.GetResponseStream()
                If inputPackageStream IsNot Nothing Then
                    ' Retreive the Package from that stream.
                    inputPackage = Package.Open(inputPackageStream)
                End If
            End If
        End If

        Return inputPackage
    End Function

    ''' <summary>
    '''   Loads, displays, and enables user annotations
    '''   for a given XPS document stream.</summary>
    ''' <param name="xpsStream">
    '''   The memory stream of the XPS document
    '''   to load, display, and annotate.</param>
    ''' <returns>
    '''   true if the document loads successfully; otherwise false.</returns>
    Public Function OpenStream(ByVal xpsStream As MemoryStream) As Boolean

        If m_xpsPackage IsNot Nothing Then
            ' The user clicked OK, close the current document and continue.
            CloseDocument()
        End If

        m_xpsPackage = Package.Open(xpsStream)
        Dim packageUriString As String = "memorystream://inMemory" & Guid.NewGuid.ToString & ".xps"
        With m_xpsPackage

            Dim m_packageUri As Uri = New Uri(packageUriString)         ' Remember to create URI for the package
            PackageStore.AddPackage(m_packageUri, m_xpsPackage)         ' Need to add the Package to the PackageStore
            m_xpsDocument = New XpsDocument(m_xpsPackage, CompressionOption.SuperFast, packageUriString)
            '                                                           ' Create instance of XpsDocument 
            m_xpsFixedDocumentSequence = m_xpsDocument.GetFixedDocumentSequence
            Me.DocViewer.Document = m_xpsFixedDocumentSequence
        End With
        Return True
    End Function


    ''' ------------------------ GetFixedDocument --------------------------
    ''' <summary>
    '''   Returns the fixed document at a given URI within
    '''   the currently open XPS package.</summary>
    ''' <param name="fixedDocUri">
    '''   The URI of the fixed document to return.</param>
    ''' <returns>
    '''   The fixed document at a given URI
    '''   within the current XPS package.</returns>
    Private Function GetFixedDocument(ByVal fixedDocUri As Uri) As FixedDocument
        Dim fixedDocument As FixedDocument = Nothing

        ' Create the URI for the fixed document within the package. The URI
        ' is used to set the Parser context so fonts & other items can load.
        Dim tempUri As New Uri(m_xpsDocumentPath, UriKind.RelativeOrAbsolute)
        Dim parserContext As New ParserContext()
        parserContext.BaseUri = PackUriHelper.Create(tempUri)

        ' Retreive the fixed document.
        Dim fixedDocPart As PackagePart = m_xpsPackage.GetPart(fixedDocUri)
        If fixedDocPart IsNot Nothing Then
            Dim fixedObject As Object = XamlReader.Load(fixedDocPart.GetStream(), parserContext)
            If fixedObject <> Nothing Then
                fixedDocument = TryCast(fixedObject, FixedDocument)
            End If
        End If

        Return fixedDocument
    End Function

    Private Sub docViewer_Unloaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles DocViewer.Unloaded
        CloseDocument()
    End Sub

    ''' <summary>
    '''   Closes the document currently displayed in
    '''   the DocumentViewer control.</summary>
    Public Sub CloseDocument()

        ' Remove the document from the DocumentViewer control.
        DocViewer.Document = Nothing

        ' If the package is open, close it.
        If m_xpsPackage IsNot Nothing Then
            m_xpsPackage.Close()
            m_xpsPackage = Nothing
        End If

        If m_packageUri <> Nothing Then
            PackageStore.RemovePackage(m_packageUri)
        End If

    End Sub

End Class
