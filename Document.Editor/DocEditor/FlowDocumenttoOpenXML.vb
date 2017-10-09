Imports DocumentFormat.OpenXml, DocumentFormat.OpenXml.Packaging, DocumentFormat.OpenXml.Wordprocessing
Public Class FlowDocumenttoOpenXML
    Private flow As FlowDocument = Nothing

    Public Sub New()

    End Sub

    Public Sub Close()
        flow = Nothing
    End Sub

    Public Function Convert(ByVal flowdoc As FlowDocument, ByVal fileloc As String)
        Using myDoc As WordprocessingDocument = WordprocessingDocument.Create(fileloc, WordprocessingDocumentType.Document)
            ' Add a new main document part. 
            Dim mainPart As MainDocumentPart = myDoc.AddMainDocumentPart()
            'Create Document tree for simple document. 
            mainPart.Document = New Wordprocessing.Document()
            'Create Body (this element contains
            'other elements that we want to include 
            Dim body As New Body()
            'Create paragraph
            For Each par As Documents.Paragraph In flowdoc.Blocks
                Dim paragraph As New Paragraph()
                Dim run_paragraph As New Run()
                ' we want to put that text into the output document
                Dim t As New TextRange(par.ElementStart, par.ElementEnd)
                Dim text_paragraph As New Text(t.Text)
                'Append elements appropriately.
                run_paragraph.Append(text_paragraph)
                paragraph.Append(run_paragraph)
                body.Append(paragraph)
            Next
            mainPart.Document.Append(body)
            ' Save changes to the main document part.
            mainPart.Document.Save()
        End Using
    End Function
End Class