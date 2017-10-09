Public Class OpenDocumenttoFlowDocument

    Public Sub New()

    End Sub

    Public Function Convert(ByVal filename As String) As FlowDocument
        Dim opendocument As New AODL.Document.TextDocuments.TextDocument
        opendocument.Load(filename)
        Dim flowdocument As New FlowDocument
        For Each p As AODL.Document.Content.Text.Paragraph In opendocument.Content
            Dim par As New Paragraph()
            For Each t As AODL.Document.Content.Text.SimpleText In p.TextContent
                'TODO: (20xx.xx) Add OpenDocument Format support
                par.ElementStart.InsertTextInRun(t.Text)
            Next
            flowdocument.Blocks.Add(par)
        Next
        Return flowdocument
    End Function
End Class