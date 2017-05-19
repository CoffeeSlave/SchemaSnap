Imports System.Windows.Documents
Imports System.Windows
Imports System.Windows.Media

Public Class DocumentPaginatorWrapper
    Inherits DocumentPaginator

    Private m_PageSize As Size
    Private m_Margin As Size
    Private m_Paginator As DocumentPaginator
    Private m_Typeface As Typeface

    Public Sub New(ByVal paginator As DocumentPaginator, ByVal pageSize As Size, ByVal margin As Size)
        m_PageSize = pageSize
        m_Margin = margin
        m_Paginator = paginator
        m_Paginator.PageSize = New Size(m_PageSize.Width - margin.Width * 2, m_PageSize.Height - margin.Height * 2)
    End Sub
    Function Move(ByVal rect As Rect) As Rect
        If rect.IsEmpty Then
            Return rect
        Else
            Return New Rect(rect.Left + m_Margin.Width, rect.Top + m_Margin.Height, rect.Width, rect.Height)
        End If
    End Function
    Public Overloads Overrides Function GetPage(ByVal pageNumber As Integer) As DocumentPage
        Dim page As DocumentPage = m_Paginator.GetPage(pageNumber)
        ' Create a wrapper visual for transformation and add extras
        Dim newpage As New ContainerVisual()
        Dim title As New DrawingVisual()
        Using ctx As DrawingContext = title.RenderOpen()
            If m_Typeface Is Nothing Then
                m_Typeface = New Typeface("Times New Roman")
            End If
            Dim text As FormattedText

            text = New FormattedText("Page " + CStr((pageNumber + 1)), System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight, m_Typeface, 14, Brushes.Black)
            ctx.DrawText(text, New Point(0, page.ContentBox.Height + 96 / 4))

            text = New FormattedText("Robs Database Compare Analysis", System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight, m_Typeface, 14, Brushes.Black)
            ctx.DrawText(text, New Point(page.ContentBox.Width / 2 - 50, -96 / 4))
        End Using

        ' Scale down page and center
        Dim smallerPage As New ContainerVisual()
        smallerPage.Children.Add(page.Visual)
        smallerPage.Transform = New MatrixTransform(0.95, 0, 0, 0.95, 0.025 * page.ContentBox.Width, 0.025 * page.ContentBox.Height)
        newpage.Children.Add(smallerPage)
        newpage.Children.Add(title)
        newpage.Transform = New TranslateTransform(m_Margin.Width, m_Margin.Height)
        Return New DocumentPage(newpage, m_PageSize, Move(page.BleedBox), Move(page.ContentBox))

    End Function
    Public Overloads Overrides ReadOnly Property IsPageCountValid() As Boolean
        Get
            Return m_Paginator.IsPageCountValid
        End Get
    End Property
    Public Overloads Overrides ReadOnly Property PageCount() As Integer
        Get
            Return m_Paginator.PageCount
        End Get
    End Property
    Public Overloads Overrides Property PageSize() As Size
        Get
            Return m_Paginator.PageSize
        End Get
        Set(ByVal value As Size)
            m_Paginator.PageSize = value
        End Set
    End Property
    Public Overloads Overrides ReadOnly Property Source() As IDocumentPaginatorSource
        Get
            Return m_Paginator.Source
        End Get
    End Property

End Class

