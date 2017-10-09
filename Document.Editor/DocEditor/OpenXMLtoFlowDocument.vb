Imports DocumentFormat.OpenXml, DocumentFormat.OpenXml.Packaging, DocumentFormat.OpenXml.Wordprocessing
Public Class OpenXMLtoFlowDocument
    Private OpenXMLFile As WordprocessingDocument = Nothing

    Public Sub New(ByVal openxmlfilename As String)
        If My.Computer.FileSystem.FileExists(openxmlfilename) Then
            Dim validator As New Validation.OpenXmlValidator()
            Dim count As Integer = 0
            'For Each [error] As Validation.ValidationErrorInfo In validator.Validate(WordprocessingDocument.Open(openxmlfilename, True))
            '    count += 1
            '    MessageBox.Show([error].Description)
            'Next
            If count = 0 Then
                OpenXMLFile = WordprocessingDocument.Open(openxmlfilename, True)
            Else
                Dim m As New MessageBoxDialog("OpenXML File is invalid!", "Error", Nothing, Nothing)
                m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
                m.Owner = My.Windows.MainWindow
                m.ShowDialog()
            End If
        Else
            Dim m As New MessageBoxDialog("File Not Found!", "Error", Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = My.Windows.MainWindow
            m.ShowDialog()
            Throw New Exception
        End If
    End Sub

    Public Sub Close()
        OpenXMLFile.Dispose()
    End Sub

    Public Function Convert()
        Dim flowdoc As New FlowDocument
        'For Each p As OpenXmlPart In OpenXMLFile.MainDocumentPart.
        'Next
        flowdoc.PageWidth = 816
        flowdoc.PageHeight = 1056
        Dim body As DocumentFormat.OpenXml.Wordprocessing.Body = OpenXMLFile.MainDocumentPart.Document.Body
        For Each oxmlobject As DocumentFormat.OpenXml.OpenXmlElement In body.ChildElements
            Dim par As DocumentFormat.OpenXml.Wordprocessing.Paragraph = TryCast(oxmlobject, DocumentFormat.OpenXml.Wordprocessing.Paragraph)
            If par IsNot Nothing Then
                Dim p As New Windows.Documents.Paragraph
                flowdoc.Blocks.Add(p)
                For Each paroxmlobject As DocumentFormat.OpenXml.OpenXmlElement In par.ChildElements
                    Dim ru As DocumentFormat.OpenXml.Wordprocessing.Run = TryCast(paroxmlobject, DocumentFormat.OpenXml.Wordprocessing.Run)
                    If ru IsNot Nothing Then
                        Dim r As New Windows.Documents.Run
                        r.Text = ru.InnerText
                        p.Inlines.Add(r)
                    End If
                Next
                'flowdoc.ContentStart.InsertTextInRun(par.InnerText)
            End If
        Next
        Return flowdoc
    End Function
End Class