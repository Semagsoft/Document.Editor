Public Class FlowDocumenttoOpenDocument

    Public Sub New()

    End Sub

    Public Function Convert(ByVal flowdoc As FlowDocument) As AODL.Document.TextDocuments.TextDocument
        Dim textdoc As New AODL.Document.TextDocuments.TextDocument
        textdoc.[New]()
        For Each p As Paragraph In flowdoc.Blocks
            Dim par As New AODL.Document.Content.Text.Paragraph(textdoc)
            Dim t As New TextRange(p.ElementStart, p.ElementEnd)
            par.TextContent.Add(New AODL.Document.Content.Text.SimpleText(textdoc, t.Text))
        Next
        Return textdoc
    End Function
End Class