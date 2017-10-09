Imports System.IO

Public Class DocumentEditor
    Inherits RichTextBox

    Public FileChanged As Boolean = False, DocumentName As String = Nothing, ZoomLevel As Double = 1
    Public docpadding As New Thickness(96, 96, 96, 96)
    Public SelectedTableCell As TableCell = Nothing, SelectedImage As Image = Nothing, SelectedVideo As MediaElement = Nothing, SelectedShape As Shape = Nothing, SelectedObject As UIElement = Nothing

    Public SelectedLineNumber As Integer = 0, LineCount As Integer = 0, SelectedColumnNumber As Integer = 0, ColumnCount As Integer = 0, WordCount As Integer = 0

    Public Sub New()
        Me.IsDocumentEnabled = True
        Document.PageWidth = 816
        Document.PageHeight = 1056
        CaretPosition.Paragraph.LineHeight = 1.15
        FontFamily = My.Settings.Options_DefaultFont
        FontSize = My.Settings.Options_DefaultFontSize
        AcceptsTab = True
        'Margin = New Thickness(-2, -3, 0, 0)
        SetPageMargins(docpadding)
        If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath + "\spellcheck_ignorelist.lex") Then
            Dim dictionaries As IList = SpellCheck.GetCustomDictionaries(Me)
            dictionaries.Add(New Uri(My.Application.Info.DirectoryPath + "\spellcheck_ignorelist.lex"))
        End If
    End Sub

    Public Sub SetupMarginsManager(document As FlowDocument, margin As Thickness)
        docpadding = margin
        Dim dpd As ComponentModel.DependencyPropertyDescriptor = ComponentModel.DependencyPropertyDescriptor.FromProperty(FlowDocument.PagePaddingProperty, GetType(FlowDocument))
        dpd.AddValueChanged(document, New EventHandler(AddressOf SetPadding))
    End Sub

    Public Sub SetPageMargins(thickness As Thickness)
        docpadding = thickness
        Document.PagePadding = docpadding
    End Sub

    Private Sub SetPadding(sender As Object, e As EventArgs)
        Dim fd As FlowDocument = DirectCast(sender, FlowDocument)
        If fd.PagePadding <> New Thickness(docpadding.Left, docpadding.Top, docpadding.Right, docpadding.Bottom) Then
            fd.PagePadding = docpadding
        End If
    End Sub

#Region "LoadDocument"

    Public Sub LoadDocument(ByVal filename As String)
        Dim f As New FileInfo(filename), tr As New TextRange(Document.ContentStart, Document.ContentEnd)
        Dim fs As FileStream = Nothing
        Dim isreadonlyfile As Boolean = False
        If f.IsReadOnly Then
            isreadonlyfile = True
        End If
        If f.Extension.ToLower = ".xamlpackage" Then
            Dim t As New TextRange(Document.ContentStart, Document.ContentEnd)
            Dim file As New FileStream(filename, FileMode.Open)
            t.Load(file, DataFormats.XamlPackage)
            file.Close()
        ElseIf f.Extension.ToLower = ".xaml" Then
            fs = File.Open(f.FullName, FileMode.Open, FileAccess.Read)
            Dim content As FlowDocument = TryCast(Markup.XamlReader.Load(fs), FlowDocument)
            Dim thi As Thickness = content.PagePadding
            Try
                Dim leftmargin As Integer = thi.Left, topmargin As Integer = thi.Top, rightmargin As Integer = thi.Right, bottommargin As Integer = thi.Bottom
                SetPageMargins(New Thickness(leftmargin, topmargin, rightmargin, bottommargin))
            Catch ex As Exception
                SetPageMargins(New Thickness(0, 0, 0, 0))
            End Try
            Document = content
        ElseIf f.Extension.ToLower = ".docx" Then
            ReadDocx(f.FullName)
            Document.PageWidth = 816
            Document.PageHeight = 1056
            FileChanged = False
        ElseIf f.Extension.ToLower = ".odt" Then
            'TODO: add odt format support
            'fs.Close()
            'fs = Nothing
            'Dim converter As New OpenDocumenttoFlowDocument
            'Document = converter.Convert(f.FullName)
        ElseIf f.Extension.ToLower = ".html" OrElse f.Extension.ToLower = ".htm" Then
                Try
                    Dim content As FlowDocument = TryCast(Markup.XamlReader.Parse(HTMLConverter.HtmlToXamlConverter.ConvertHtmlToXaml(My.Computer.FileSystem.ReadAllText(filename), True)), FlowDocument)
                    Document = content
                    Document.PageWidth = 816
                Catch ex As Exception
                    Dim m As New MessageBoxDialog("Error loading html document", "Error", Nothing, Nothing)
                    m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
                    m.Owner = My.Windows.MainWindow
                    m.ShowDialog()
                End Try
            ElseIf f.Extension.ToLower = ".rtf" Then
                If isreadonlyfile Then
                fs = File.Open(f.FullName, FileMode.Open, FileAccess.Read)
            Else
                fs = File.Open(f.FullName, FileMode.Open)
            End If
            tr.Load(fs, DataFormats.Rtf)
        Else
                If isreadonlyfile Then
                fs = File.Open(f.FullName, FileMode.Open, FileAccess.Read)
            Else
                fs = File.Open(f.FullName, FileMode.Open)
            End If
            tr.Load(fs, DataFormats.Text)
        End If
        If fs IsNot Nothing Then
            fs.Close()
            fs = Nothing
        End If
        If f.IsReadOnly Then
            IsReadOnly = True
        End If
        DocumentName = filename
        Dim helper As New Semagsoft.HyperlinkHelper
        helper.SubscribeToAllHyperlinks(Document)
        'Dim helper2 As New Semagsoft.ImageHelper
        'helper2.AddImageResizers(Me)
        'p.SetDocumentTitle(f.Name)
        FileChanged = False
    End Sub

    Private Sub ReadDocx(path__1 As String)
        Using stream = File.Open(path__1, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim flowDocumentConverter = New DocxReaderApplication.DocxToFlowDocumentConverter(stream)
            flowDocumentConverter.Read()
            Document = flowDocumentConverter.Document
        End Using
    End Sub

#End Region

#Region "SaveDocument"

    Public Sub SaveDocument(ByVal filename As String)
        If My.Computer.FileSystem.FileExists(filename) Then
            My.Computer.FileSystem.DeleteFile(filename)
        End If
        Dim file As New FileInfo(filename), fs As FileStream = IO.File.Open(filename, FileMode.OpenOrCreate),
            tr As New TextRange(Document.ContentStart, Document.ContentEnd)
        If file.Extension.ToLower = ".xamlpackage" Then
            fs.Close()
            fs = Nothing
            Dim range As TextRange
            Dim fStream As FileStream
            range = New TextRange(Document.ContentStart, Document.ContentEnd)
            fStream = New FileStream(filename, FileMode.Create)
            range.Save(fStream, DataFormats.XamlPackage, True)
            fStream.Close()
        ElseIf file.Extension.ToLower = ".xaml" Then
            Markup.XamlWriter.Save(Document, fs)
            fs.Close()
            fs = Nothing
            Dim xd As New Xml.XmlDocument()
            xd.LoadXml(My.Computer.FileSystem.ReadAllText(filename))
            Dim sb As New Text.StringBuilder()
            Dim sw As New StringWriter(sb)
            Dim xtw As Xml.XmlTextWriter = Nothing
            Try
                xtw = New Xml.XmlTextWriter(sw)
                xtw.Formatting = Xml.Formatting.Indented
                xd.WriteTo(xtw)
            Finally
                If xtw IsNot Nothing Then
                    xtw.Close()
                End If
            End Try
            Dim tex As String = sb.ToString
            'Dim final As String
            'Try
            '    final = tex.Remove(tex.IndexOf("</FlowDocument>"), tex.Length)
            'Catch ex As Exception
            'End Try
            My.Computer.FileSystem.WriteAllText(filename, tex, False)
        ElseIf file.Extension.ToLower = ".docx" Then
            fs.Close()
            fs = Nothing
            Dim converter As New FlowDocumenttoOpenXML
            converter.Convert(Document, filename)
            converter.Close()
        ElseIf file.Extension.ToLower = ".odt" Then
            'TODO: add odt format support
            '    fs.Close()
            '    fs = Nothing
            '    Dim converter As New FlowDocumenttoOpenDocument
            '    Dim opendoc As AODL.Document.TextDocuments.TextDocument = converter.Convert(Document)
            '    opendoc.SaveTo(filename)
        ElseIf file.Extension.ToLower = ".html" OrElse file.Extension.ToLower = ".htm" Then
            fs.Close()
            fs = Nothing
            Dim s As String = Markup.XamlWriter.Save(Document)
            Try
                My.Computer.FileSystem.WriteAllText(filename, HTMLConverter.HtmlFromXamlConverter.ConvertXamlToHtml(s), False)
            Catch ex As Exception
                Dim m As New MessageBoxDialog("Error saving document", "Error", Nothing, Nothing)
                m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
                m.Owner = My.Windows.MainWindow
                m.ShowDialog()
            End Try
        ElseIf file.Extension.ToLower = ".rtf" Then
            tr.Save(fs, DataFormats.Rtf)
        Else
            tr.Save(fs, DataFormats.Text)
        End If
        If fs IsNot Nothing Then
            fs.Close()
            fs = Nothing
        End If
        'doctab.SetDocumentTitle(file.Name)
        DocumentName = filename
        FileChanged = False
    End Sub

#End Region

    Private Sub DocumentEditor_LayoutUpdated(sender As Object, e As EventArgs) Handles Me.LayoutUpdated
        SetPageMargins(docpadding)
    End Sub

    Private Sub DocumentEditor_SelectionChanged(sender As Object, e As RoutedEventArgs) Handles Me.SelectionChanged
        Dim ls As TextPointer = CaretPosition.GetLineStartPosition(0), p As TextPointer = Document.ContentStart.GetLineStartPosition(0),
            int As Integer = 1, int2 As Integer = 1
        While True
            If ls.CompareTo(p) < 1 Then
                Exit While
            End If
            Dim r As Integer
            p = p.GetLineStartPosition(1, r)
            If r = 0 Then
                Exit While
            End If
            int += 1
        End While
        Dim ls2 As TextPointer = Document.ContentStart.DocumentEnd.GetLineStartPosition(0), p2 As TextPointer = Document.ContentEnd.DocumentStart.GetLineStartPosition(0)
        While True
            If ls2.CompareTo(p2) < 1 Then
                Exit While
            End If
            Dim r As Integer
            p2 = p2.GetLineStartPosition(1, r)
            If r = 0 Then
                Exit While
            End If
            int2 += 1
        End While
        SelectedLineNumber = int
        LineCount = int2
        Dim t As New TextRange(Document.ContentStart, Document.ContentEnd)
        Dim caretPos As TextPointer = CaretPosition, poi As TextPointer = CaretPosition.GetLineStartPosition(0)
        Dim currentColumnNumber As Integer = Math.Max(p.GetOffsetToPosition(caretPos) - 1, 0) + 1, currentColumnCount As Integer = currentColumnNumber
        currentColumnCount += CaretPosition.GetTextRunLength(LogicalDirection.Forward)
        SelectedColumnNumber = currentColumnNumber
        ColumnCount = currentColumnCount
    End Sub

    Private Sub DocumentEditor_TextChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs) Handles Me.TextChanged
        If FileChanged = False Then
            FileChanged = True
        End If
        Dim helper As New Semagsoft.HyperlinkHelper
        helper.SubscribeToAllHyperlinks(Document)
        WordCount = GetWordCount()
    End Sub

#Region "Find"

    Public Function FindWordFromPosition(ByVal position As TextPointer, ByVal word As String) As TextRange
        While position IsNot Nothing
            If position.GetPointerContext(LogicalDirection.Forward) = TextPointerContext.Text Then
                Dim textRun As String = position.GetTextInRun(LogicalDirection.Forward)
                ' Find the starting index of any substring that matches "word".
                Dim indexInRun As Integer = textRun.IndexOf(word)
                If indexInRun >= 0 Then
                    Dim start As TextPointer = position.GetPositionAtOffset(indexInRun)
                    Dim [end] As TextPointer = start.GetPositionAtOffset(word.Length)
                    Return New TextRange(start, [end])
                End If
            End If
            position = position.GetNextContextPosition(LogicalDirection.Forward)
        End While
        ' position will be null if "word" is not found.
        Return Nothing
    End Function

#End Region

#Region "Subscript/Superscript"

    Public Sub ToggleSubscript()
        Dim currentAlignment = Me.Selection.GetPropertyValue(Inline.BaselineAlignmentProperty)
        Dim newAlignment As BaselineAlignment = If((DirectCast(currentAlignment, BaselineAlignment) = BaselineAlignment.Subscript), BaselineAlignment.Baseline, BaselineAlignment.Subscript)
        Me.Selection.ApplyPropertyValue(Inline.BaselineAlignmentProperty, newAlignment)
    End Sub

    Public Sub ToggleSuperscript()
        Dim currentAlignment = Me.Selection.GetPropertyValue(Inline.BaselineAlignmentProperty)
        Dim newAlignment As BaselineAlignment = If((DirectCast(currentAlignment, BaselineAlignment) = BaselineAlignment.Superscript), BaselineAlignment.Baseline, BaselineAlignment.Superscript)
        Me.Selection.ApplyPropertyValue(Inline.BaselineAlignmentProperty, newAlignment)
    End Sub

#End Region

#Region "Strikethrough"

    Public Sub ToggleStrikethrough()
        Dim range As New TextRange(Selection.Start, Selection.End)
        Dim t As TextDecorationCollection = DirectCast(Selection.GetPropertyValue(Inline.TextDecorationsProperty), TextDecorationCollection)
        If t Is Nothing OrElse Not t.Equals(TextDecorations.Strikethrough) Then
            t = TextDecorations.Strikethrough
        Else
            t = New TextDecorationCollection()
        End If
        range.ApplyPropertyValue(Inline.TextDecorationsProperty, t)
    End Sub

#End Region

#Region "GetWordCount"

    Public Function GetWordCount() As Integer
        Dim SpacePos As Integer, X As Integer = 1, WordCount As Integer = 0, NoMore As Boolean = False
        Dim CharValue As Integer, tr As New TextRange(Document.ContentStart, Document.ContentEnd)
        Dim content As String = tr.Text
        content = content.Replace(vbCr, " ")
        content = content.Replace(vbLf, " ")
        If content.Trim.Length > 0 Then
            Do While NoMore = False
                SpacePos = InStr(X, Trim(content), " ")
                If SpacePos > 0 Then
                    CharValue = Asc(content.Substring(X - 1, 1))
                    If CharValue > 64 AndAlso CharValue < 91 OrElse CharValue > 96 AndAlso CharValue < 123 OrElse CharValue > 47 AndAlso CharValue < 58 Then
                        WordCount += 1
                    End If
                    X = SpacePos + 1
                    Do While InStr(X, (content.Substring(X - 1, 1)), " ") > 0
                        X += 1
                    Loop
                Else
                    If X <= content.Length Then
                        CharValue = Asc(content.Substring(X - 1, 1))
                        If CharValue > 64 AndAlso CharValue < 91 OrElse CharValue > 96 AndAlso CharValue < 123 OrElse CharValue > 47 AndAlso CharValue < 58 Then
                            WordCount += 1
                        End If
                    End If
                    NoMore = True
                End If
            Loop
        End If
        Return WordCount
    End Function

#End Region

#Region "GoToLine"

    Public Sub GoToLine(ByVal linenumber As Integer)
        If linenumber = 1 Then
            Me.CaretPosition = Document.ContentStart.DocumentStart.GetLineStartPosition(0)
        Else
            Dim ls As TextPointer = Document.ContentStart.DocumentStart.GetLineStartPosition(0)
            Dim p As TextPointer = Document.ContentStart.GetLineStartPosition(0), int As Integer = 2
            While True
                Dim r As Integer
                p = p.GetLineStartPosition(1, r)
                If r = 0 Then
                    Me.CaretPosition = p
                    Exit While
                End If
                If linenumber = int Then
                    Me.CaretPosition = p
                    Exit While
                End If
                int += 1
            End While
        End If
    End Sub

#End Region

End Class