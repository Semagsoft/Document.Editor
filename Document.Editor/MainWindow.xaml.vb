Class MainWindow
    Private SelectedDocument As DocumentTab = Nothing, Speech As New Speech.Synthesis.SpeechSynthesizer, IsJumpListAdded As Boolean = False
    Private WithEvents myJumpList As New Shell.JumpList, DocumentPreview As New DocumentEditor, DocPreviewScrollViewer As New ScrollViewer

#Region "Reuseable Code"

#Region "Update"

    Public Function PathExists(ByVal path As String, ByVal timeout As Integer) As Boolean
        Dim exists As Boolean = True
        Dim t As New System.Threading.Thread(DirectCast(Function() CheckPathFunction(path), System.Threading.ThreadStart))
        t.Start()
        Dim completed As Boolean = t.Join(timeout)
        If Not completed Then
            exists = False
            t.Abort()
        End If
        Return exists
    End Function

    Public Function CheckPathFunction(ByVal path As String) As Boolean
        Return System.IO.File.Exists(path)
    End Function

    Private WithEvents CheckForUpdateWorker As New System.ComponentModel.BackgroundWorker

    Private Sub CheckForUpdateWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles CheckForUpdateWorker.DoWork
        Try
            If PathExists("http://semagsoft.com/software/updates/document.editor.update", 5000) Then
                Dim fileReader As New Net.WebClient
                My.Computer.FileSystem.CreateDirectory(My.Computer.FileSystem.SpecialDirectories.Temp + "\Semagsoft\DocumentEditor\")
                Dim filename As String = My.Computer.FileSystem.SpecialDirectories.Temp + "\Semagsoft\DocumentEditor\updatechecker.ini"
                If IO.File.Exists(filename) Then
                    IO.File.Delete(filename)
                End If
                fileReader.DownloadFile(New Uri("http://semagsoft.com/software/updates/document.editor.update"), filename)
                e.Result = True
            Else
                e.Result = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ex.ToString, MessageBoxButton.OK, MessageBoxImage.Error)
            e.Result = False
        End Try
    End Sub

    Private Sub CheckForUpdateWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles CheckForUpdateWorker.RunWorkerCompleted
        If e.Result = True Then
            Dim textreader As IO.TextReader = IO.File.OpenText(My.Computer.FileSystem.SpecialDirectories.Temp + "\Semagsoft\DocumentEditor\updatechecker.ini")
            Dim version As String = textreader.ReadLine
            Dim versionyear As Integer = Convert.ToInt16(version.Substring(0, 4))
            Dim versionnumber As Integer = Convert.ToInt16(version.Substring(5))
            textreader.Close()
            If versionyear >= My.Application.Info.Version.Major AndAlso versionnumber > My.Application.Info.Version.Minor Then
                Dim a As New AboutDialog
                a.IsCheckingForUpdates = True
                a.ShowDialog()
            End If
        Else
        End If
    End Sub

#End Region

    Private Sub UpdateDocumentPreview()
        Dim range As New TextRange(SelectedDocument.Editor.Document.ContentStart, SelectedDocument.Editor.Document.ContentEnd), stream As New IO.MemoryStream
        Markup.XamlWriter.Save(range, stream)
        range.Save(stream, DataFormats.XamlPackage, True)
        Dim previewdoc As New FlowDocument
        Dim range2 As New TextRange(previewdoc.ContentEnd, previewdoc.ContentEnd)
        range2.Load(stream, DataFormats.XamlPackage)
        'TODO: (20xx.xx) set background color for preview document
        DocumentPreview.Document.PageWidth = SelectedDocument.Editor.Document.PageWidth
        DocumentPreview.Document.PageHeight = SelectedDocument.Editor.Document.PageHeight
        DocumentPreview.Width = DocumentPreview.Document.PageWidth
        DocumentPreview.Height = DocumentPreview.Document.PageHeight
        DocumentPreview.Document = previewdoc
        DocumentPreview.Document.PagePadding = SelectedDocument.Editor.docpadding
        DocumentPreview.InvalidateVisual()
        DocumentPreview.UpdateLayout()
        Dim c As Canvas = TryCast(DocPreviewScrollViewer.Content, Canvas)
        If c.Children.Count = 0 Then
            c.Children.Add(DocumentPreview)
        End If
        DocumentPreview.InvalidateVisual()
        DocumentPreview.UpdateLayout()
        DocPreviewScrollViewer.Content = c
    End Sub

#Region "Document"

    Public Sub NewDocument(ByVal title As String)
        If TabCell.Children.Count > 0 Then
            SelectedDocument.IsSelected = False
        End If
        Dim tb As New DocumentTab(title, Brushes.Transparent)
        tb.Ruler.Background = Background
        TabCell.Children.Add(tb)
        UpdateUI()
        tb.IsSelected = True
        SelectedDocument.Editor.Focus()
        SelectedDocument.Editor.FileChanged = False
        UpdateButtons()
    End Sub

    Private Sub CloseDocument(ByVal file As DocumentTab)
        TabCell.Children.Remove(file)
        UpdateUI()
    End Sub

#End Region

    Private Sub RunPlugin(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim p As Fluent.Button = e.Source, plugins As New Plugins, i As Object = plugins.Build(p.Header, My.Computer.FileSystem.ReadAllText(p.Tag))
        If i.GetType.Name.ToString = "String" Then
            SelectedDocument.Editor.CaretPosition.InsertTextInRun(i)
        End If
    End Sub

#Region "Themes"

    Private Enum Theme
        Office2010
        Office2013
        Windows8
    End Enum

    Private currentTheme As Nullable(Of Theme)

    Private Sub SetRibbonTheme_Windows8(sender As Object, e As RoutedEventArgs)
        Me.ChangeTheme(Theme.Windows8, "pack://application:,,,/Fluent;component/Themes/Windows8/Generic.xaml")
    End Sub

    Private Sub SetRibbonTheme_Office2013(sender As Object, e As RoutedEventArgs)
        Me.ChangeTheme(Theme.Office2013, "pack://application:,,,/Fluent;component/Themes/Office2013/Generic.xaml")
    End Sub

    Private Sub SetRibbonTheme__Office2010Silver(sender As Object, e As RoutedEventArgs)
        Me.ChangeTheme(Theme.Office2010, "pack://application:,,,/Fluent;component/Themes/Office2010/Silver.xaml")
    End Sub

    Private Sub SetRibbonTheme_Office2010Black(sender As Object, e As RoutedEventArgs)
        Me.ChangeTheme(Theme.Office2010, "pack://application:,,,/Fluent;component/Themes/Office2010/Black.xaml")
    End Sub

    Private Sub SetRibbonTheme_Office2010Blue(sender As Object, e As RoutedEventArgs)
        Me.ChangeTheme(Theme.Office2010, "pack://application:,,,/Fluent;component/Themes/Office2010/Blue.xaml")
    End Sub

    Private Sub ChangeTheme(theme_1 As Theme, color As String)
        Me.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ApplicationIdle, DirectCast(Function()

                                                                                                              Dim owner = Window.GetWindow(Me)
                                                                                                              If owner IsNot Nothing Then
                                                                                                                  owner.Resources.BeginInit()

                                                                                                                  If owner.Resources.MergedDictionaries.Count > 0 Then
                                                                                                                      owner.Resources.MergedDictionaries.RemoveAt(0)
                                                                                                                  End If

                                                                                                                  If String.IsNullOrEmpty(color) = False Then
                                                                                                                      owner.Resources.MergedDictionaries.Add(New ResourceDictionary() With {
                                                                                                                          .Source = New Uri(color)
                                                                                                                      })
                                                                                                                  End If

                                                                                                                  owner.Resources.EndInit()
                                                                                                              End If

                                                                                                              If Me.currentTheme <> theme_1 Then
                                                                                                                  Application.Current.Resources.BeginInit()
                                                                                                                  Select Case theme_1
                                                                                                                      Case Theme.Office2010
                                                                                                                          Application.Current.Resources.MergedDictionaries.Add(New ResourceDictionary() With {
                                                                                                                              .Source = New Uri("pack://application:,,,/Fluent;component/Themes/Generic.xaml")
                                                                                                                          })
                                                                                                                          Application.Current.Resources.MergedDictionaries.RemoveAt(0)
                                                                                                                          Exit Select
                                                                                                                      Case Theme.Office2013
                                                                                                                          Application.Current.Resources.MergedDictionaries.Add(New ResourceDictionary() With {
                                                                                                                              .Source = New Uri("pack://application:,,,/Fluent;component/Themes/Office2013/Generic.xaml")
                                                                                                                          })
                                                                                                                          Application.Current.Resources.MergedDictionaries.RemoveAt(0)
                                                                                                                          Exit Select
                                                                                                                      Case Theme.Windows8
                                                                                                                          Application.Current.Resources.MergedDictionaries.Add(New ResourceDictionary() With {
                                                                                                                              .Source = New Uri("pack://application:,,,/Fluent;component/Themes/Windows8/Generic.xaml")
                                                                                                                          })
                                                                                                                          Application.Current.Resources.MergedDictionaries.RemoveAt(0)
                                                                                                                          Exit Select
                                                                                                                  End Select

                                                                                                                  Me.currentTheme = theme_1
                                                                                                                  Application.Current.Resources.EndInit()

                                                                                                                  If owner IsNot Nothing Then
                                                                                                                      owner.Style = Nothing
                                                                                                                      owner.Style = TryCast(owner.FindResource("RibbonWindowStyle"), Style)
                                                                                                                      owner.Style = Nothing
                                                                                                                  End If
                                                                                                              End If

                                                                                                          End Function, Threading.ThreadStart))
    End Sub

#End Region

#Region "RibbonBar"

    Private Sub UpdateSelected()
        UpdateContextualTabs()
        If SelectedDocument.Editor.FileChanged Then
            Title = SelectedDocument.DocName + " * - " + My.Application.Info.ProductName
        Else
            Title = SelectedDocument.DocName + " - " + My.Application.Info.ProductName
        End If
        MainBar.Title = Title
        If My.Computer.FileSystem.FileExists(SelectedDocument.Editor.DocumentName) Then
            Dim kb As Double = My.Computer.FileSystem.GetFileInfo(SelectedDocument.Editor.DocumentName).Length / 1024, inter As Integer = Convert.ToInt32(kb)
            FileSizeTextBlock.Text = "Size: " + inter.ToString + " KB"
        Else
            FileSizeTextBlock.Text = "Size: " + "0 KB"
        End If
        LinesTextBlock.Text = "Line: " + SelectedDocument.Editor.SelectedLineNumber.ToString + " of " + SelectedDocument.Editor.LineCount.ToString
        ColumnsTextBlock.Text = "Column: " + SelectedDocument.Editor.SelectedColumnNumber.ToString + " of " + SelectedDocument.Editor.ColumnCount.ToString
        WordCountTextBlock.Text = "Word Count: " + SelectedDocument.Editor.WordCount.ToString
        If SelectedDocument.Editor.CanUndo Then
            UndoButton.IsEnabled = True
        Else
            UndoButton.IsEnabled = False
        End If
        If SelectedDocument.Editor.CanRedo Then
            RedoButton.IsEnabled = True
        Else
            RedoButton.IsEnabled = False
        End If
        If My.Computer.Clipboard.ContainsText OrElse My.Computer.Clipboard.ContainsImage Then
            PasteButton.IsEnabled = True
            If My.Computer.Clipboard.ContainsText Then
                PasteTextButton.IsEnabled = True
            Else
                PasteTextButton.IsEnabled = False
            End If
            If My.Computer.Clipboard.ContainsImage Then
                PasteImageMenuItem.IsEnabled = True
            Else
                PasteImageMenuItem.IsEnabled = False
            End If
        Else
            PasteButton.IsEnabled = False
            PasteTextButton.IsEnabled = False
            PasteImageMenuItem.IsEnabled = False
        End If
        If SelectedDocument.Editor.FileChanged Then
            SaveMenuItem.IsEnabled = True
        Else
            SaveMenuItem.IsEnabled = False
        End If
        UpdatingFont = True
        If FontComboBox.IsLoaded Then
            Dim value As Object = SelectedDocument.Editor.Selection.GetPropertyValue(TextElement.FontFamilyProperty)
            Dim currentfontfamily As FontFamily = TryCast(value, FontFamily)
            If currentfontfamily IsNot Nothing Then
                FontComboBox.SelectedItem = currentfontfamily
            End If
        End If
        If FontSizeComboBox.IsLoaded Then
            Try
                Dim sizevalue As Double = SelectedDocument.Editor.Selection.GetPropertyValue(TextElement.FontSizeProperty)
                FontSizeComboBox.SelectedValue = sizevalue
            Catch ex As Exception
            End Try
        End If
        UpdatingFont = False
        Dim r As Run = TryCast(SelectedDocument.Editor.CaretPosition.Parent, Run)
        If r IsNot Nothing Then
            If r.FontWeight = FontWeights.Bold Then
                BoldButton.IsChecked = True
            Else
                BoldButton.IsChecked = False
            End If
            If r.FontStyle = FontStyles.Italic Then
                ItalicButton.IsChecked = True
            Else
                ItalicButton.IsChecked = False
            End If
            Dim td As TextDecorationCollection = TryCast(SelectedDocument.Editor.Selection.GetPropertyValue(Inline.TextDecorationsProperty), TextDecorationCollection)
            UnderlineButton.IsChecked = False
            StrikethroughButton.IsChecked = False
            If td Is TextDecorations.Underline Then
                UnderlineButton.IsChecked = True
            End If
            If td Is TextDecorations.Strikethrough Then
                StrikethroughButton.IsChecked = True
            End If
            If r.BaselineAlignment = BaselineAlignment.Subscript Then
                SubscriptButton.IsChecked = True
                SuperscriptButton.IsChecked = False
            ElseIf r.BaselineAlignment = BaselineAlignment.Superscript Then
                SubscriptButton.IsChecked = False
                SuperscriptButton.IsChecked = True
            Else
                SubscriptButton.IsChecked = False
                SuperscriptButton.IsChecked = False
            End If
            Dim runparent As Paragraph = TryCast(r.Parent, Paragraph)
            If runparent IsNot Nothing Then
                If runparent.TextAlignment = TextAlignment.Left Then
                    AlignLeftButton.IsChecked = True
                    AlignCenterButton.IsChecked = False
                    AlignRightButton.IsChecked = False
                    AlignJustifyButton.IsChecked = False
                ElseIf runparent.TextAlignment = TextAlignment.Center Then
                    AlignLeftButton.IsChecked = False
                    AlignCenterButton.IsChecked = True
                    AlignRightButton.IsChecked = False
                    AlignJustifyButton.IsChecked = False
                ElseIf runparent.TextAlignment = TextAlignment.Right Then
                    AlignLeftButton.IsChecked = False
                    AlignCenterButton.IsChecked = False
                    AlignRightButton.IsChecked = True
                    AlignJustifyButton.IsChecked = False
                ElseIf runparent.TextAlignment = TextAlignment.Justify Then
                    AlignLeftButton.IsChecked = False
                    AlignCenterButton.IsChecked = False
                    AlignRightButton.IsChecked = False
                    AlignJustifyButton.IsChecked = True
                End If
                Dim listitem As ListItem = TryCast(runparent.Parent, ListItem)
                If listitem IsNot Nothing Then
                    Dim list As List = TryCast(listitem.Parent, List)
                    If list IsNot Nothing Then
                        If list.MarkerStyle = TextMarkerStyle.Disc OrElse list.MarkerStyle = TextMarkerStyle.Circle OrElse list.MarkerStyle = TextMarkerStyle.Box OrElse list.MarkerStyle = TextMarkerStyle.Square Then
                            BulletListButton.IsChecked = True
                        Else
                            BulletListButton.IsChecked = False
                        End If
                        If list.MarkerStyle = TextMarkerStyle.Decimal OrElse list.MarkerStyle = TextMarkerStyle.UpperLatin OrElse list.MarkerStyle = TextMarkerStyle.LowerLatin OrElse list.MarkerStyle = TextMarkerStyle.UpperRoman OrElse list.MarkerStyle = TextMarkerStyle.LowerRoman Then
                            NumberListButton.IsChecked = True
                        Else
                            NumberListButton.IsChecked = False
                        End If
                    Else
                        BulletListButton.IsChecked = False
                        NumberListButton.IsChecked = False
                    End If
                Else
                    BulletListButton.IsChecked = False
                    NumberListButton.IsChecked = False
                End If
            End If
        Else
            Dim p As Paragraph = TryCast(SelectedDocument.Editor.CaretPosition.Parent, Paragraph)
            If p IsNot Nothing Then
                If p.FontWeight = FontWeights.Bold Then
                    BoldButton.IsChecked = True
                Else
                    BoldButton.IsChecked = False
                End If
                If p.FontStyle = FontStyles.Italic Then
                    ItalicButton.IsChecked = True
                Else
                    ItalicButton.IsChecked = False
                End If
                Dim td As TextDecorationCollection = TryCast(SelectedDocument.Editor.Selection.GetPropertyValue(Inline.TextDecorationsProperty), TextDecorationCollection)
                UnderlineButton.IsChecked = False
                StrikethroughButton.IsChecked = False
                If td Is TextDecorations.Underline Then
                    UnderlineButton.IsChecked = True
                End If
                If td Is TextDecorations.Strikethrough Then
                    StrikethroughButton.IsChecked = True
                End If
                Dim runparent As Paragraph = TryCast(p, Paragraph)
                If runparent IsNot Nothing Then
                    If runparent.TextAlignment = TextAlignment.Left Then
                        AlignLeftButton.IsChecked = True
                        AlignCenterButton.IsChecked = False
                        AlignRightButton.IsChecked = False
                        AlignJustifyButton.IsChecked = False
                    ElseIf runparent.TextAlignment = TextAlignment.Center Then
                        AlignLeftButton.IsChecked = False
                        AlignCenterButton.IsChecked = True
                        AlignRightButton.IsChecked = False
                        AlignJustifyButton.IsChecked = False
                    ElseIf runparent.TextAlignment = TextAlignment.Right Then
                        AlignLeftButton.IsChecked = False
                        AlignCenterButton.IsChecked = False
                        AlignRightButton.IsChecked = True
                        AlignJustifyButton.IsChecked = False
                    ElseIf runparent.TextAlignment = TextAlignment.Justify Then
                        AlignLeftButton.IsChecked = False
                        AlignCenterButton.IsChecked = False
                        AlignRightButton.IsChecked = False
                        AlignJustifyButton.IsChecked = True
                    End If
                End If
                Dim listitem As ListItem = TryCast(p.Parent, ListItem)
                If listitem IsNot Nothing Then
                    Dim list As List = TryCast(listitem.Parent, List)
                    If list IsNot Nothing Then
                        If list.MarkerStyle = TextMarkerStyle.Disc OrElse list.MarkerStyle = TextMarkerStyle.Circle OrElse list.MarkerStyle = TextMarkerStyle.Box OrElse list.MarkerStyle = TextMarkerStyle.Square Then
                            BulletListButton.IsChecked = True
                        Else
                            BulletListButton.IsChecked = False
                        End If
                        If list.MarkerStyle = TextMarkerStyle.Decimal OrElse list.MarkerStyle = TextMarkerStyle.UpperLatin OrElse list.MarkerStyle = TextMarkerStyle.LowerLatin OrElse list.MarkerStyle = TextMarkerStyle.UpperRoman OrElse list.MarkerStyle = TextMarkerStyle.LowerRoman Then
                            NumberListButton.IsChecked = True
                        Else
                            NumberListButton.IsChecked = False
                        End If
                    Else
                        BulletListButton.IsChecked = False
                        NumberListButton.IsChecked = False
                    End If
                Else
                    BulletListButton.IsChecked = False
                    NumberListButton.IsChecked = False
                End If
            Else
                BulletListButton.IsChecked = False
                NumberListButton.IsChecked = False
            End If
        End If
        If SelectedDocument.Editor.CaretPosition.Paragraph IsNot Nothing Then
            LineSpacing1Point0.IsChecked = False
            LineSpacing1Point15.IsChecked = False
            LineSpacing1Point5.IsChecked = False
            LineSpacing2Point0.IsChecked = False
            LineSpacing2Point5.IsChecked = False
            LineSpacing3Point0.IsChecked = False
            If SelectedDocument.Editor.CaretPosition.Paragraph.LineHeight = 1.0 Then
                LineSpacing1Point0.IsChecked = True
            ElseIf SelectedDocument.Editor.CaretPosition.Paragraph.LineHeight = 1.15 Then
                LineSpacing1Point15.IsChecked = True
            ElseIf SelectedDocument.Editor.CaretPosition.Paragraph.LineHeight = 1.5 Then
                LineSpacing1Point5.IsChecked = True
            ElseIf SelectedDocument.Editor.CaretPosition.Paragraph.LineHeight = 2.0 Then
                LineSpacing2Point0.IsChecked = True
            ElseIf SelectedDocument.Editor.CaretPosition.Paragraph.LineHeight = 2.5 Then
                LineSpacing2Point5.IsChecked = True
            ElseIf SelectedDocument.Editor.CaretPosition.Paragraph.LineHeight = 3.0 Then
                LineSpacing3Point0.IsChecked = True
            End If
        End If
    End Sub

    Private Sub UpdateButtons()
        If SelectedDocument IsNot Nothing Then
            If SelectedDocument.Editor.Document.Blocks.Count = 0 Then
                SelectedDocument.Editor.Document.Blocks.Add(New Paragraph)
            End If
            UpdateSelected()
            SearchPanel_StackPanel.IsEnabled = True
            If SelectedDocument.Editor.DocumentName IsNot Nothing AndAlso SelectedDocument.Editor.FileChanged Then
                RevertMenuItem.IsEnabled = True
            Else
                RevertMenuItem.IsEnabled = False
            End If
            CloseMenuItem.IsEnabled = True
            CloseAllButThisMenuItem.IsEnabled = True
            CloseAllMenuItem.IsEnabled = True
            SaveAsMenuItem.IsEnabled = True
            SaveCopyMenuItem.IsEnabled = True
            SaveAllMenuItem.IsEnabled = True
            ExportMenuItem.IsEnabled = True
            ExportXPSMenuItem.IsEnabled = True
            ExportArchiveMenuItem.IsEnabled = True
            ExportImageMenuItem.IsEnabled = True
            ExportSoundMenuItem.IsEnabled = True
            PropertiesMenuItem.IsEnabled = True
            PrintTab.IsEnabled = True
            PrintMenuItem.IsEnabled = True
            PageSetupMenuItem.IsEnabled = True
            ClipboardGroup.IsEnabled = True
            zoomSlider.Value = SelectedDocument.Editor.ZoomLevel * 100
            ZoomGroup.IsEnabled = True
            If zoomSlider.Value > 10 Then
                ZoomOutButton.IsEnabled = True
            Else
                ZoomOutButton.IsEnabled = False
            End If
            If zoomSlider.Value < 500 Then
                ZoomInButton.IsEnabled = True
            Else
                ZoomInButton.IsEnabled = False
            End If
            zoomSlider.IsEnabled = True
            LeftMarginBox.Value = SelectedDocument.Editor.docpadding.Left
            TopMarginBox.Value = SelectedDocument.Editor.docpadding.Top
            RightMarginBox.Value = SelectedDocument.Editor.docpadding.Right
            BottomMarginBox.Value = SelectedDocument.Editor.docpadding.Bottom
            PageHeightBox.Value = SelectedDocument.Editor.Document.PageHeight
            PageWidthBox.Value = SelectedDocument.Editor.Document.PageWidth
            ClipboardGroup.IsEnabled = True
            FontGroup.IsEnabled = True
            ParagraphGroup.IsEnabled = True
            StylesGroup.IsEnabled = True
            EditingGroup.IsEnabled = True
            Insert_TablesGroup.IsEnabled = True
            Insert_DateAndTimeGroup.IsEnabled = True
            Insert_IllustrationGroup.IsEnabled = True
            Insert_MiscGroup.IsEnabled = True
            Design_PageBackgroundGroup.IsEnabled = True
            PageLayout_PageSetupGroup.IsEnabled = True
            Navigation_LineGroup.IsEnabled = True
            Navigation_PageGroup.IsEnabled = True
            Navigation_DocumentGroup.IsEnabled = True
            Review_Proofing.IsEnabled = True
            Review_Language.IsEnabled = True
            If StatusbarButton.IsChecked Then
                StatusBar.Visibility = Visibility.Visible
            End If
        ElseIf dockingManager.Layout.ActiveContent Is Nothing Then
            SelectedDocument = Nothing
            SearchPanel_StackPanel.IsEnabled = False
            Title = My.Application.Info.ProductName
            MainBar.Title = My.Application.Info.ProductName
            BackStage.SelectedIndex = 0
            MainBar.SelectedTabIndex = 0
            RevertMenuItem.IsEnabled = False
            CloseMenuItem.IsEnabled = False
            CloseAllButThisMenuItem.IsEnabled = False
            CloseAllMenuItem.IsEnabled = False
            SaveMenuItem.IsEnabled = False
            SaveAsMenuItem.IsEnabled = False
            SaveCopyMenuItem.IsEnabled = False
            SaveAllMenuItem.IsEnabled = False
            ExportMenuItem.IsEnabled = False
            ExportXPSMenuItem.IsEnabled = False
            ExportArchiveMenuItem.IsEnabled = False
            ExportImageMenuItem.IsEnabled = False
            ExportSoundMenuItem.IsEnabled = False
            PropertiesMenuItem.IsEnabled = False
            PrintTab.IsEnabled = False
            PrintMenuItem.IsEnabled = False
            PageSetupMenuItem.IsEnabled = False
            UndoButton.IsEnabled = False
            RedoButton.IsEnabled = False
            ClipboardGroup.IsEnabled = False
            ClipboardGroup.IsEnabled = False
            FontGroup.IsEnabled = False
            ParagraphGroup.IsEnabled = False
            BulletListButton.IsChecked = False
            NumberListButton.IsChecked = False
            AlignLeftButton.IsChecked = False
            AlignCenterButton.IsChecked = False
            AlignRightButton.IsChecked = False
            AlignJustifyButton.IsChecked = False
            StylesGroup.IsEnabled = False
            EditingGroup.IsEnabled = False
            Insert_TablesGroup.IsEnabled = False
            Insert_DateAndTimeGroup.IsEnabled = False
            Insert_IllustrationGroup.IsEnabled = False
            Insert_MiscGroup.IsEnabled = False
            Design_PageBackgroundGroup.IsEnabled = False
            PageLayout_PageSetupGroup.IsEnabled = False
            Navigation_LineGroup.IsEnabled = False
            Navigation_PageGroup.IsEnabled = False
            Navigation_DocumentGroup.IsEnabled = False
            Review_Proofing.IsEnabled = False
            Review_Language.IsEnabled = False
            ZoomGroup.IsEnabled = False
            zoomSlider.IsEnabled = False
            EditTableTab.Visibility = Visibility.Collapsed
            EditTableCellTab.Visibility = Visibility.Collapsed
            EditListTab.Visibility = Visibility.Collapsed
            EditImageTab.Visibility = Visibility.Collapsed
            EditVideoTab.Visibility = Visibility.Collapsed
            EditObjectTab.Visibility = Visibility.Collapsed
            StatusBar.Visibility = Visibility.Collapsed
        End If
    End Sub

#End Region

    Public Sub UpdateUI()
        If SelectedDocument IsNot Nothing Then
            If My.Settings.MainWindow_ShowRuler Then
                If TabCell.Children.Count >= 2 Then
                    For Each t As DocumentTab In TabCell.Children
                        t.Ruler.Margin = New Thickness(-23, 2, 0, 0)
                        t.VSV.Margin = New Thickness(-6, 26, 0, 0)
                    Next
                Else
                    For Each t As DocumentTab In TabCell.Children
                        t.Ruler.Margin = New Thickness(-23, 0, 0, 0)
                        t.VSV.Margin = New Thickness(-6, 24, 0, 0)
                    Next
                End If
            Else
                If TabCell.Children.Count >= 2 Then
                    For Each t As DocumentTab In TabCell.Children
                        t.VSV.Margin = New Thickness(-6, 0, 0, 0)
                    Next
                Else
                    For Each t As DocumentTab In TabCell.Children
                        t.VSV.Margin = New Thickness(-6, -1, 0, 0)
                    Next
                End If
            End If
        End If
    End Sub

#End Region

#Region "MainWindow"

#Region "Add Handlers"

    Private Sub addhandlers()
        Me.AddHandler(DocumentTab.UpdateSelected, New RoutedEventHandler(AddressOf UpdateButtons))
        'Me.AddHandler(DocumentTab.InsertObjectEvent, New RoutedEventHandler(AddressOf ObjectButton_Click))
        Me.AddHandler(DocumentTab.InsertShapeEvent, New RoutedEventHandler(AddressOf ShapeButton_Click))
        Me.AddHandler(DocumentTab.InsertImageEvent, New RoutedEventHandler(AddressOf ImageButton_Click))
        Me.AddHandler(DocumentTab.InsertLinkEvent, New RoutedEventHandler(AddressOf LinkMenuItem_Click))
        Me.AddHandler(DocumentTab.InsertFlowDocumentEvent, New RoutedEventHandler(AddressOf InsertFlowDocumentButton_Click))
        Me.AddHandler(DocumentTab.InsertRichTextFileEvent, New RoutedEventHandler(AddressOf InsertRichTextDocumentButton_Click))
        Me.AddHandler(DocumentTab.InsertTextFileEvent, New RoutedEventHandler(AddressOf TextFileButton_Click))
        Me.AddHandler(DocumentTab.InsertSymbolEvent, New RoutedEventHandler(AddressOf SymbolContextMenuItem_Click))
        Me.AddHandler(DocumentTab.InsertTableEvent, New RoutedEventHandler(AddressOf TableMenuItem_Click))
        Me.AddHandler(DocumentTab.InsertVideoEvent, New RoutedEventHandler(AddressOf VideoButton_Click))
        Me.AddHandler(DocumentTab.InsertHorizontalLineEvent, New RoutedEventHandler(AddressOf HorizontalLineButton_Click))
        Me.AddHandler(DocumentTab.InsertHeaderEvent, New RoutedEventHandler(AddressOf HeaderButton_Click))
        Me.AddHandler(DocumentTab.InsertFooterEvent, New RoutedEventHandler(AddressOf FooterButton_Click))
        Me.AddHandler(DocumentTab.InsertDateEvent, New RoutedEventHandler(AddressOf DateMenuItem_Click))
        Me.AddHandler(DocumentTab.InsertTimeEvent, New RoutedEventHandler(AddressOf TimeMenuItem_Click))
        Me.AddHandler(DocumentTab.ClearFormattingEvent, New RoutedEventHandler(AddressOf ClearFormattingButton_Click))
        Me.AddHandler(DocumentTab.FontEvent, New RoutedEventHandler(AddressOf FontContextMenuItem_Click))
        Me.AddHandler(DocumentTab.FontSizeEvent, New RoutedEventHandler(AddressOf FontSizeContextMenuItem_Click))
        Me.AddHandler(DocumentTab.FontColorEvent, New RoutedEventHandler(AddressOf FontColorContextMenuItem_Click))
        Me.AddHandler(DocumentTab.HighlightColorEvent, New RoutedEventHandler(AddressOf FontHighlightColorContextMenuItem_Click))
        Me.AddHandler(DocumentTab.BoldEvent, New RoutedEventHandler(AddressOf BoldMenuItem_Click))
        Me.AddHandler(DocumentTab.ItalicEvent, New RoutedEventHandler(AddressOf ItalicMenuItem_Click))
        Me.AddHandler(DocumentTab.UnderlineEvent, New RoutedEventHandler(AddressOf UnderlineMenuItem_Click))
        Me.AddHandler(DocumentTab.StrikethroughEvent, New RoutedEventHandler(AddressOf StrikethroughButton_Click))
        Me.AddHandler(DocumentTab.SubscriptEvent, New RoutedEventHandler(AddressOf SubscriptButton_Click))
        Me.AddHandler(DocumentTab.SuperscriptEvent, New RoutedEventHandler(AddressOf SuperscriptButton_Click))
        Me.AddHandler(DocumentTab.IndentMoreEvent, New RoutedEventHandler(AddressOf IndentMoreButton_Click))
        Me.AddHandler(DocumentTab.IndentLessEvent, New RoutedEventHandler(AddressOf IndentLessButton_Click))
        Me.AddHandler(DocumentTab.BulletListEvent, New RoutedEventHandler(AddressOf BulletListMenuItem_Click))
        Me.AddHandler(DocumentTab.NumberListEvent, New RoutedEventHandler(AddressOf NumberListMenuItem_Click))
        Me.AddHandler(DocumentTab.AlignLeftEvent, New RoutedEventHandler(AddressOf AlignLeftMenuItem_Click))
        Me.AddHandler(DocumentTab.AlignCenterEvent, New RoutedEventHandler(AddressOf AlignCenterMenuItem_Click))
        Me.AddHandler(DocumentTab.AlignRightEvent, New RoutedEventHandler(AddressOf AlignRightMenuItem_Click))
        Me.AddHandler(DocumentTab.AlignJustifyEvent, New RoutedEventHandler(AddressOf AlignJustifyMenuItem_Click))
        Me.AddHandler(DocumentTab.LineSpacingEvent, New RoutedEventHandler(AddressOf CustomLineSpacingMenuItem_Click))
        Me.AddHandler(DocumentTab.LeftToRightEvent, New RoutedEventHandler(AddressOf LefttoRightButton_Click))
        Me.AddHandler(DocumentTab.RightToLeftEvent, New RoutedEventHandler(AddressOf RighttoLeftButton_Click))
        Me.AddHandler(DocumentTab.UndoEvent, New RoutedEventHandler(AddressOf UndoMenuItem_Click))
        Me.AddHandler(DocumentTab.RedoEvent, New RoutedEventHandler(AddressOf RedoMenuItem_Click))
        Me.AddHandler(DocumentTab.CutEvent, New RoutedEventHandler(AddressOf CutMenuItem_Click))
        Me.AddHandler(DocumentTab.CopyEvent, New RoutedEventHandler(AddressOf CopyMenuItem_Click))
        Me.AddHandler(DocumentTab.PasteEvent, New RoutedEventHandler(AddressOf PasteMenuItem_Click))
        Me.AddHandler(DocumentTab.DeleteEvent, New RoutedEventHandler(AddressOf DeleteMenuItem_Click))
        Me.AddHandler(DocumentTab.SelectAllEvent, New RoutedEventHandler(AddressOf SelectAllMenuItem_Click))
        Me.AddHandler(DocumentTab.FindEvent, New RoutedEventHandler(AddressOf FindButton_Click))
        Me.AddHandler(DocumentTab.ReplaceEvent, New RoutedEventHandler(AddressOf ReplaceButton_Click))
        Me.AddHandler(DocumentTab.GoToEvent, New RoutedEventHandler(AddressOf GoToButton_Click))
    End Sub

    Private Sub SymbolContextMenuItem_Click(sender As Object, e As EventArgs)
        MainBar.SelectedTabItem = InsertTab
        InsertSymbolButton.IsDropDownOpen = True
    End Sub

    Private Sub FontContextMenuItem_Click(sender As Object, e As EventArgs)
        MainBar.SelectedTabItem = HomeTabItem
        FontComboBox.IsDropDownOpen = True
    End Sub

    Private Sub FontSizeContextMenuItem_Click(sender As Object, e As EventArgs)
        MainBar.SelectedTabItem = HomeTabItem
        FontSizeComboBox.IsDropDownOpen = True
    End Sub

    Private Sub FontColorContextMenuItem_Click(sender As Object, e As EventArgs)
        MainBar.SelectedTabItem = HomeTabItem
        FontColorButton.IsDropDownOpen = True
    End Sub

    Private Sub FontHighlightColorContextMenuItem_Click(sender As Object, e As EventArgs)
        MainBar.SelectedTabItem = HomeTabItem
        HighlightColorButton.IsDropDownOpen = True
    End Sub

#End Region

#Region "Activated"

    Private Sub MainWindow_Activated(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Activated
        If My.Computer.Info.OSVersion >= "6.1" Then
            If Not IsJumpListAdded Then
                Shell.JumpList.SetJumpList(Application.Current, myJumpList)
                Dim OnlineHelpJumpTask As New Shell.JumpTask
                With OnlineHelpJumpTask
                    .CustomCategory = "Links"
                    .IconResourceIndex = 0
                    .Title = "Online Help"
                    .Description = "Get Help Online"
                    .ApplicationPath = "http://documenteditor.net/documentation/"
                    .Arguments = Nothing
                End With
                Dim PluginsJumpTask As New Shell.JumpTask
                With PluginsJumpTask
                    .CustomCategory = "Links"
                    .IconResourceIndex = 0
                    .Title = "Get Plugins"
                    .Description = "Get Plugins for Document.Editor"
                    .ApplicationPath = "http://documenteditor.net/plugins/"
                    .Arguments = Nothing
                End With
                Dim WebsiteJumpTask As New Shell.JumpTask
                With WebsiteJumpTask
                    .CustomCategory = "Links"
                    .IconResourceIndex = 0
                    .Title = "Website"
                    .Description = "Visit Website"
                    .ApplicationPath = "http://documenteditor.net/"
                    .Arguments = Nothing
                End With
                myJumpList.JumpItems.Add(OnlineHelpJumpTask)
                myJumpList.JumpItems.Add(PluginsJumpTask)
                myJumpList.JumpItems.Add(WebsiteJumpTask)
                myJumpList.Apply()
                IsJumpListAdded = True
            End If
        End If
    End Sub

#End Region

#Region "Key Down"

    Private Sub MainWindow_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles Me.KeyDown
        If TabCell.SelectedContent IsNot Nothing Then
            If Keyboard.IsKeyDown(Key.LeftCtrl) OrElse Keyboard.IsKeyDown(Key.RightCtrl) AndAlso Not Keyboard.IsKeyDown(Key.LeftShift) Then
                If e.Key = Key.F Then
                    e.Handled = True
                    FindButton_Click(Nothing, Nothing)
                ElseIf e.Key = Key.G Then
                    e.Handled = True
                    GoToButton_Click(Nothing, Nothing)
                ElseIf e.Key = Key.H Then
                    e.Handled = True
                    ReplaceButton_Click(Nothing, Nothing)
                ElseIf e.Key = Key.N Then
                    NewMenuItem_Click(Nothing, Nothing)
                ElseIf e.Key = Key.O Then
                    e.Handled = True
                    OpenMenuItem_Click(Nothing, Nothing)
                ElseIf e.Key = Key.P Then
                    e.Handled = True
                    PrintMenuItem_Click(Nothing, Nothing)
                ElseIf e.Key = Key.S Then
                    e.Handled = True
                    If SaveMenuItem.IsEnabled Then
                        SaveMenuItem_Click(Nothing, Nothing)
                    End If
                ElseIf e.Key = Key.W Then
                    e.Handled = True
                    CloseMenuItem_Click(Nothing, Nothing)
                    'ElseIf Keyboard.IsKeyDown(Key.OemPlus) Then
                    '    'TODO: Add zoom in shotcut key
                    '    e.Handled = True
                    '    ZoomInButton_Click(Nothing, Nothing)
                    'ElseIf e.Key = Key.OemMinus Then
                    '    e.Handled = True
                    '    If ZoomOutButton.IsEnabled Then
                    '        ZoomOutButton_Click(Nothing, Nothing)
                    '    End If
                End If
            End If
            If Keyboard.IsKeyDown(Key.Insert) Then
                If e.Key = Key.O Then
                    e.Handled = True
                    'TODO: ObjectButton_Click(Nothing, Nothing)
                ElseIf e.Key = Key.H Then
                    e.Handled = True
                    HorizontalLineButton_Click(Nothing, Nothing)
                ElseIf e.Key = Key.I Then
                    e.Handled = True
                    ImageButton_Click(Nothing, Nothing)
                ElseIf e.Key = Key.L Then
                    e.Handled = True
                    LinkMenuItem_Click(Nothing, Nothing)
                ElseIf e.Key = Key.S Then
                    e.Handled = True
                    ShapeButton_Click(Nothing, Nothing)
                ElseIf e.Key = Key.T Then
                    e.Handled = True
                    TableMenuItem_Click(Nothing, Nothing)
                ElseIf e.Key = Key.V Then
                    e.Handled = True
                    VideoButton_Click(Nothing, Nothing)
                ElseIf e.Key = Key.X Then
                    e.Handled = True
                    TextFileButton_Click(Nothing, Nothing)
                End If
            End If
        Else
            If Keyboard.IsKeyDown(Key.LeftCtrl) OrElse Keyboard.IsKeyDown(Key.RightCtrl) AndAlso Not Keyboard.IsKeyDown(Key.LeftShift) Then
                If e.Key = Key.N Then
                    NewMenuItem_Click(Nothing, Nothing)
                ElseIf e.Key = Key.O Then
                    e.Handled = True
                    OpenMenuItem_Click(Nothing, Nothing)
                End If
            End If
        End If
    End Sub

#End Region

#Region "Initialized"

    Private Sub MainWindow_Initialized(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Initialized
        Dim t As New Threading.Thread(AddressOf addhandlers), plugins As New Plugins
        t.Start()
        Me.Show()
        If My.Settings.Options_TemplatesFolder Is "" Then
            If My.Computer.FileSystem.DirectoryExists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Semagsoft\Document.Editor\Templates\") Then
                My.Settings.Options_TemplatesFolder = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Semagsoft\Document.Editor\Templates\"
            End If
        End If
        If My.Computer.FileSystem.DirectoryExists(My.Settings.Options_TemplatesFolder) Then
            For Each file As String In My.Computer.FileSystem.GetFiles(My.Settings.Options_TemplatesFolder)
                Dim template As New IO.FileInfo(file)
                If template.Extension = ".xaml" Then
                    Dim item As New Fluent.Button
                    If My.Settings.Options_Theme = 2 Then
                        item.Foreground = ReadOnlyMenuItem.Foreground
                    End If
                    item.CanAddToQuickAccessToolBar = False
                    Dim tname As String = template.Name.Remove(template.Name.Length - 5)
                    If My.Computer.FileSystem.FileExists(My.Settings.Options_TemplatesFolder + tname.ToLower + ".png") Then
                        Dim bitmap As New BitmapImage(New Uri(My.Settings.Options_TemplatesFolder + tname.ToLower + ".png"))
                        item.Icon = bitmap
                    End If
                    item.Header = tname
                    item.Tag = file
                    item.Foreground = ReadOnlyMenuItem.Foreground
                    AddHandler item.Click, New RoutedEventHandler(AddressOf NewFromTemplate)
                    NewMenuItem.Children.Add(item)
                End If
            Next
        End If
        If My.Settings.Options_EnablePlugins Then
            For Each file As String In My.Computer.FileSystem.GetFiles(My.Application.Info.DirectoryPath + "\Plugins\")
                Try
                    Dim plugin As New IO.FileInfo(file), pluginname As String = plugin.Name.Remove(plugin.Name.Length - 3)
                    Dim IsStartupPlugin As Boolean = plugins.IsStartupPlugin(pluginname, My.Computer.FileSystem.ReadAllText(file))
                    If IsStartupPlugin Then
                        Dim i As Object = plugins.Build(pluginname, My.Computer.FileSystem.ReadAllText(file))
                        If i.GetType.Name.ToString = "String" Then
                            SelectedDocument.Editor.CaretPosition.InsertTextInRun(i)
                        End If
                    End If
                    Dim IsEventPlugin As Boolean = plugins.IsEventPlugin(pluginname, My.Computer.FileSystem.ReadAllText(file))
                    If IsEventPlugin Then
                        Dim ribbonbutton As New Fluent.Button
                        ribbonbutton.Header = pluginname
                        If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath + "\Plugins\" + pluginname + "32.png") Then
                            Dim bitmap As New BitmapImage(New Uri(My.Application.Info.DirectoryPath + "\Plugins\" + pluginname + "32.png"))
                            ribbonbutton.LargeIcon = bitmap
                        Else
                            Dim bitmap As New BitmapImage(New Uri("pack://application:,,,/Images/Tools/plugins32.png"))
                            ribbonbutton.LargeIcon = bitmap
                        End If
                        Dim tip As New Fluent.ScreenTip
                        tip.Title = pluginname
                        tip.Image = New BitmapImage(New Uri("pack://application:,,,/Images/Tools/plugins48.png"))
                        ribbonbutton.ToolTip = tip
                        ribbonbutton.Tag = file
                        AddHandler ribbonbutton.Click, New RoutedEventHandler(AddressOf RunPlugin)
                        PluginsGroup.Items.Add(ribbonbutton)
                    End If
                Catch ex As Exception
                    Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
                    m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
                    m.Owner = Me
                    m.ShowDialog()
                End Try
            Next
        End If
        If PluginsGroup.Items.Count = 0 Then
            PluginsGroup.Visibility = Visibility.Collapsed
        End If
    End Sub

#End Region

#Region "Loaded"

    Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles Me.Loaded
        If My.Settings.MainWindow_IsMax Then
            Me.WindowState = WindowState.Maximized
        End If
        SearchPanel.IsVisible = False
        DocPreviewScrollViewer.Content = New Canvas
        MainGrid.Children.Add(DocPreviewScrollViewer)
        FontComboBox.SelectedItem = My.Settings.Options_DefaultFont
        With FontSizeComboBox
            .Items.Add(Convert.ToDouble(8))
            .Items.Add(Convert.ToDouble(9))
            .Items.Add(Convert.ToDouble(10))
            .Items.Add(Convert.ToDouble(11))
            For i = 12 To 28 Step 2
                .Items.Add(Convert.ToDouble(i))
            Next i
            .Items.Add(Convert.ToDouble(36))
            .Items.Add(Convert.ToDouble(48))
            .Items.Add(Convert.ToDouble(72))
            .SelectedItem = My.Settings.Options_DefaultFontSize
        End With
        If My.Settings.Options_Theme = 0 Then
            SetRibbonTheme_Office2010Blue(Nothing, Nothing)
            dockingManager.Theme = New Xceed.Wpf.AvalonDock.Themes.AeroTheme
            dockingManagerGrid.Margin = New Thickness(0, -4, 0, -30)
            dockingManager.Margin = New Thickness(-8, 0, -5, 0)
        ElseIf My.Settings.Options_Theme = 1 Then
            SetRibbonTheme__Office2010Silver(Nothing, Nothing)
            dockingManager.Theme = New Xceed.Wpf.AvalonDock.Themes.ExpressionLightTheme
            dockingManagerGrid.Margin = New Thickness(0, -4, 0, -30)
            dockingManager.Margin = New Thickness(-8, 0, -5, 0)
        ElseIf My.Settings.Options_Theme = 2 Then
            SetRibbonTheme_Office2010Black(Nothing, Nothing)
            dockingManager.Theme = New Xceed.Wpf.AvalonDock.Themes.ExpressionDarkTheme
            dockingManagerGrid.Margin = New Thickness(0, -4, 0, -30)
            dockingManager.Margin = New Thickness(-8, 0, -5, 0)
        ElseIf My.Settings.Options_Theme = 3 Then
            SetRibbonTheme_Office2013(Nothing, Nothing)
            dockingManager.Theme = New Xceed.Wpf.AvalonDock.Themes.MetroTheme
            MainGrid.ClipToBounds = True
            dockingManagerGrid.Margin = New Thickness(0, -8, 0, -30)
            dockingManager.Margin = New Thickness(-8, 0, -5, 0)

            LinesTextBlock.Foreground = Brushes.LightGray
            ColumnsTextBlock.Foreground = Brushes.LightGray
            WordCountTextBlock.Foreground = Brushes.LightGray
            FileSizeTextBlock.Foreground = Brushes.LightGray
            ZoomTextBlock.Foreground = Brushes.LightGray
        ElseIf My.Settings.Options_Theme = 4 Then
            SetRibbonTheme_Windows8(Nothing, Nothing)
            dockingManager.Theme = New Xceed.Wpf.AvalonDock.Themes.MetroTheme
            MainGrid.ClipToBounds = True
            dockingManagerGrid.Margin = New Thickness(0, -8, 0, -30)
            dockingManager.Margin = New Thickness(-8, 0, -5, 0)
        End If
        NewButton.Foreground = ReadOnlyMenuItem.Foreground
        ImportFTPConnect.Foreground = ReadOnlyMenuItem.Foreground
        ImportFTPImportButton.Foreground = ReadOnlyMenuItem.Foreground
        ImportArchiveMenuItem.Foreground = ReadOnlyMenuItem.Foreground
        ImportImageMenuItem.Foreground = ReadOnlyMenuItem.Foreground
        ExportWordpressExportButton.Foreground = ReadOnlyMenuItem.Foreground
        ExportEmailExportButton.Foreground = ReadOnlyMenuItem.Foreground
        ExportFTPExportButton.Foreground = ReadOnlyMenuItem.Foreground
        ExportXPSMenuItem.Foreground = ReadOnlyMenuItem.Foreground
        ExportArchiveMenuItem.Foreground = ReadOnlyMenuItem.Foreground
        ExportImageMenuItem.Foreground = ReadOnlyMenuItem.Foreground
        ExportSoundMenuItem.Foreground = ReadOnlyMenuItem.Foreground
        PrintMenuItem.Foreground = ReadOnlyMenuItem.Foreground
        PageSetupMenuItem.Foreground = ReadOnlyMenuItem.Foreground
        If My.Settings.Options_ShowRecentDocuments AndAlso My.Settings.Options_RecentFiles.Count > 0 Then
            For Each s As String In My.Settings.Options_RecentFiles
                If My.Computer.FileSystem.FileExists(s) Then
                    Dim i2 As New Fluent.Button, f As New IO.FileInfo(s)
                    Dim ContMenu As New ContextMenu
                    Dim removemenuitem As New MenuItem
                    removemenuitem.Header = "Remove"
                    removemenuitem.ToolTip = f.FullName
                    ContMenu.Items.Add(removemenuitem)
                    i2.ContextMenu = ContMenu
                    Dim img As New Image
                    If f.Extension.ToLower = ".xamlpackage" Then
                        img.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Tab/xaml16.png"))
                    ElseIf f.Extension.ToLower = ".xaml" Then
                        img.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Tab/xaml16.png"))
                    ElseIf f.Extension.ToLower = ".rtf" Then
                        img.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Tab/rtf16.png"))
                    ElseIf f.Extension.ToLower = ".html" OrElse f.Extension.ToLower = ".htm" Then
                        img.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Tab/html16.png"))
                    Else
                        img.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Tab/txt16.png"))
                    End If
                    i2.Icon = img
                    i2.Header = f.Name
                    i2.Tag = s
                    i2.Foreground = ReadOnlyMenuItem.Foreground
                    RecentFilesList.Children.Add(i2)
                    AddHandler(i2.Click), New RoutedEventHandler(AddressOf recentfile_click)
                    AddHandler removemenuitem.Click, New RoutedEventHandler(AddressOf recentfileremove_click)
                End If
            Next
        Else
            BackStage.Items.Remove(RecentFilesTabItem)
        End If
        If My.Settings.Options_TabPlacement = 0 Then
            'TabCell.TabStripPlacement = Dock.Top
        ElseIf My.Settings.Options_TabPlacement = 1 Then
            'TabCell.TabStripPlacement = Dock.Left
        ElseIf My.Settings.Options_TabPlacement = 2 Then
            'TabCell.TabStripPlacement = Dock.Right
        ElseIf My.Settings.Options_TabPlacement = 3 Then
            'TabCell.TabStripPlacement = Dock.Bottom
        End If
        If My.Settings.MainWindow_ShowRuler Then
            HRulerButton.IsChecked = True
        End If
        dockingManager.Margin = New Thickness(dockingManager.Margin.Left, dockingManager.Margin.Top, dockingManager.Margin.Right, 21)
        If My.Settings.MainWindow_ShowStatusBar Then
            StatusbarButton.IsChecked = True
        Else
            MainGrid.RowDefinitions.Item(2).Height = New GridLength(0)
            StatusbarButton.IsChecked = False
        End If
        Try
            If My.Application.StartUpFileNames.Count > 0 Then
                If My.Settings.Options_StartupMode = 1 AndAlso My.Settings.Options_PreviousDocuments.Count > 0 Then
                    CloseMenuItem_Click(Me, Nothing)
                    For Each s As String In My.Settings.Options_PreviousDocuments
                        If My.Computer.FileSystem.FileExists(s) Then
                            Dim f As New IO.FileInfo(s)
                            NewDocument(f.Name)
                            SelectedDocument.Editor.LoadDocument(f.FullName)
                            SelectedDocument.Editor.Height = SelectedDocument.Editor.Document.PageHeight
                            SelectedDocument.Editor.Width = SelectedDocument.Editor.Document.PageWidth
                            SelectedDocument.Ruler.Width = SelectedDocument.Editor.Width
                            Dim ch As Semagsoft.DocRuler.Ruler = SelectedDocument.Ruler.Children.Item(0)
                            ch.Width = SelectedDocument.Editor.Width
                        End If
                    Next
                End If
                For Each s As String In My.Application.StartUpFileNames
                    Dim f As New IO.FileInfo(s)
                    NewDocument(f.Name)
                    SelectedDocument.Editor.LoadDocument(f.FullName)
                    SelectedDocument.Editor.Height = SelectedDocument.Editor.Document.PageHeight
                    SelectedDocument.Editor.Width = SelectedDocument.Editor.Document.PageWidth
                    SelectedDocument.Ruler.Width = SelectedDocument.Editor.Width
                    Dim ch As Semagsoft.DocRuler.Ruler = SelectedDocument.Ruler.Children.Item(0)
                    ch.Width = SelectedDocument.Editor.Width
                Next
            Else
                NewMenuItem_Click(Me, Nothing)
                If My.Settings.Options_StartupMode = 1 AndAlso My.Settings.Options_PreviousDocuments.Count > 0 Then
                    CloseMenuItem_Click(Me, Nothing)
                    For Each s As String In My.Settings.Options_PreviousDocuments
                        If My.Computer.FileSystem.FileExists(s) Then
                            Dim f As New IO.FileInfo(s)
                            NewDocument(f.Name)
                            SelectedDocument.Editor.LoadDocument(f.FullName)
                            SelectedDocument.Editor.Height = SelectedDocument.Editor.Document.PageHeight
                            SelectedDocument.Editor.Width = SelectedDocument.Editor.Document.PageWidth
                            SelectedDocument.Ruler.Width = SelectedDocument.Editor.Width
                            Dim ch As Semagsoft.DocRuler.Ruler = SelectedDocument.Ruler.Children.Item(0)
                            ch.Width = SelectedDocument.Editor.Width
                        End If
                    Next
                ElseIf My.Settings.Options_StartupMode = 2 Then
                    CloseMenuItem_Click(Me, Nothing)
                    OpenMenuItem_Click(Me, Nothing)
                ElseIf My.Settings.Options_StartupMode = 3 Then
                    CloseMenuItem_Click(Me, Nothing)
                End If
            End If
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
        If My.Settings.Options_CheckForUpdatesOnStartup Then
            CheckForUpdateWorker.RunWorkerAsync()
        End If
        If My.Settings.Options_ShowStartupDialog Then
            Dim startup As New StartupDialog
            startup.Owner = Me
            startup.Show()
        End If
    End Sub

#End Region

#Region "Closing"

    Private Sub MainWindow_Closing(ByVal sender As Object, ByVal e As ComponentModel.CancelEventArgs) Handles Me.Closing
        If Me.WindowState = WindowState.Maximized Then
            My.Settings.MainWindow_IsMax = True
        ElseIf Me.WindowState = WindowState.Normal Then
            My.Settings.MainWindow_IsMax = False
            Dim wi As Integer, hi As Integer, le As Integer, ti As Integer
            wi = Convert.ToInt32(Me.ActualWidth)
            hi = Convert.ToInt32(Me.ActualHeight)
            le = Convert.ToInt32(Left)
            ti = Convert.ToInt32(Top)
            Dim s As New System.Windows.Size(wi, hi)
            My.Settings.MainWindow_Left = Left
            My.Settings.MainWindow_Top = Top
            My.Settings.MainWindow_Height = Height
            My.Settings.MainWindow_Width = Width
        End If
        My.Settings.MainWindow_ShowRuler = HRulerButton.IsChecked
        My.Settings.MainWindow_ShowStatusBar = StatusbarButton.IsChecked
        If My.Settings.Options_StartupMode = 1 AndAlso My.Settings.Options_PreviousDocuments.Count > 0 Then
            My.Settings.Options_PreviousDocuments.Clear()
        End If
        While TabCell.Children.Count > 0
            If SelectedDocument.Editor.FileChanged Then
                Dim save As New SaveFileDialog
                save.Owner = Me
                save.SetFileInfo(SelectedDocument.DocName, SelectedDocument.Editor)
                Dim fs As IO.FileStream = IO.File.Open(My.Computer.FileSystem.SpecialDirectories.Temp + "\TVPre.xaml", IO.FileMode.Create)
                Dim tr As New TextRange(SelectedDocument.Editor.Document.ContentStart, SelectedDocument.Editor.Document.ContentEnd)
                Markup.XamlWriter.Save(SelectedDocument.Editor.Document, fs)
                fs.Close()
                My.Settings.Options_PreviousDocuments.Clear()
                save.ShowDialog()
                If save.Res = "Yes" Then
                    If SelectedDocument.Editor.DocumentName IsNot Nothing Then
                        SelectedDocument.Editor.SaveDocument(SelectedDocument.Editor.DocumentName)
                        My.Settings.Options_PreviousDocuments.Add(SelectedDocument.Editor.DocumentName)
                        CloseDocument(SelectedDocument)
                    Else
                        SaveMenuItem_Click(Nothing, Nothing)
                        If SelectedDocument.Editor.DocumentName IsNot Nothing Then
                            My.Settings.Options_PreviousDocuments.Add(SelectedDocument.Editor.DocumentName)
                        End If
                    End If
                ElseIf save.Res = "No" Then
                    If SelectedDocument.Editor.DocumentName IsNot Nothing Then
                        My.Settings.Options_PreviousDocuments.Add(SelectedDocument.Editor.DocumentName)
                    End If
                    CloseDocument(SelectedDocument)
                Else
                    e.Cancel = True
                    Exit While
                End If
            Else
                If SelectedDocument.Editor.DocumentName IsNot Nothing Then
                    My.Settings.Options_PreviousDocuments.Add(SelectedDocument.Editor.DocumentName)
                End If
                CloseDocument(SelectedDocument)
            End If
        End While
        My.Settings.Save()
    End Sub

#End Region

    Private Sub MainWindow_StateChanged(sender As Object, e As EventArgs) Handles Me.StateChanged
        If IsFullscreen AndAlso WindowState = WindowState.Normal Then
            FullscreenMenuItem.IsChecked = False
        End If
    End Sub

#End Region

#Region "Ribbon"

#Region "--DocumentMenu"

    Private Sub BackStageMenu_IsOpenChanged(sender As Object, e As DependencyPropertyChangedEventArgs) Handles BackStageMenu.IsOpenChanged
        If BackStageMenu.IsOpen AndAlso SelectedDocument IsNot Nothing Then
            Dim tr As New TextRange(SelectedDocument.Editor.Document.ContentStart, SelectedDocument.Editor.Document.ContentEnd)
            Dim lineint As Integer, ls2 As TextPointer = SelectedDocument.Editor.Document.ContentStart.DocumentEnd.GetLineStartPosition(0), p2 As TextPointer = SelectedDocument.Editor.Document.ContentEnd.DocumentStart.GetLineStartPosition(0)
            While True
                If ls2.CompareTo(p2) < 1 Then
                    Exit While
                End If
                Dim r As Integer
                p2 = p2.GetLineStartPosition(1, r)
                If r = 0 Then
                    Exit While
                End If
                lineint += 1
            End While
            Dim p As Integer = 0
            For Each b As Block In SelectedDocument.Editor.Document.Blocks
                p += 1
            Next
            Dim CC As Int32 = tr.Text.Length - 2
            Dim LC As Int32 = lineint + 1
            Dim PC As Int32 = p
            StatisticsCharacterCountTextBlock.Text = "Character Count: " + CC.ToString
            StatisticsWordCountTextBlock.Text = "Word Count: " + SelectedDocument.Editor.GetWordCount.ToString
            StatisticsLineCountTextBlock.Text = "Line Count: " + LC.ToString
            StatisticsParagraphCountTextBlock.Text = "Paragraph Count: " + PC.ToString
            UpdateDocumentPreview()
            Dim s As String = My.Computer.FileSystem.SpecialDirectories.Temp + "\" + System.IO.Path.GetRandomFileName
            Dim doc As New Xps.Packaging.XpsDocument(s, IO.FileAccess.ReadWrite)
            Dim xw As Xps.XpsDocumentWriter = Xps.Packaging.XpsDocument.CreateXpsDocumentWriter(doc)
            xw.Write(DocumentPreview)
            doc.Close()
            Dim xpsdoc As New Xps.Packaging.XpsDocument(s, IO.FileAccess.Read)
            Dim fds As FixedDocumentSequence = xpsdoc.GetFixedDocumentSequence
            PrintPreview.Document = fds
            xpsdoc.Close()
        End If
    End Sub

#Region "New"

    Private Sub NewMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles NewButton.Click
        NewDocument("New Document")
        'SelectedDocument.HeaderContent.FileTypeImage.ToolTip = ".xaml"
    End Sub

    Private Sub NewFromTemplate(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim t As Fluent.Button = e.Source, template As New IO.FileInfo(t.Tag),
            fs As New IO.FileStream(template.FullName, IO.FileMode.Open), flow As New FlowDocument
        flow = TryCast(Markup.XamlReader.Load(fs), FlowDocument)
        fs.Close()
        NewDocument("New Document")
        'SelectedDocument.HeaderContent.FileTypeImage.ToolTip = ".xaml"
        Dim thi As Thickness = flow.PagePadding
        SelectedDocument.Editor.Document = flow
        SelectedDocument.Editor.Focus()
        SelectedDocument.Editor.DocumentName = Nothing
        SelectedDocument.Editor.FileChanged = False
        SelectedDocument.Editor.Height = SelectedDocument.Editor.Document.PageHeight
        SelectedDocument.Editor.Width = SelectedDocument.Editor.Document.PageWidth
        SelectedDocument.Ruler.Width = SelectedDocument.Editor.Width
        Dim ch As Semagsoft.DocRuler.Ruler = SelectedDocument.Ruler.Children.Item(0)
        ch.Width = SelectedDocument.Editor.Width
        Try
            Dim leftmargin As Integer = thi.Left, topmargin As Integer = thi.Top, rightmargin As Integer = thi.Right, bottommargin As Integer = thi.Bottom
            SelectedDocument.Editor.SetPageMargins(New Thickness(leftmargin, topmargin, rightmargin, bottommargin))
        Catch ex As Exception
            SelectedDocument.Editor.SetPageMargins(New Thickness(0, 0, 0, 0))
        End Try
    End Sub

#End Region

#Region "Open"

    Private Sub OpenMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles OpenMenuItem.Click
        Dim openfile As New Microsoft.Win32.OpenFileDialog
        openfile.Multiselect = True
        openfile.Filter = "Supported Documents(*.xamlpackage;*.xaml,*.docx,*.html,*.htm,*.rtf,*.txt)|*.xamlpackage;*.xaml;*.docx;*.html;*.htm;*.rtf;*.txt|XAML Packages(*.xamlpackage)|*.xamlpackage|FlowDocuments(*.xaml)|*.xaml|OpenXML Documents(*.docx)|*.docx|HTML Documents(*.html;*.htm)|*.html;*.htm|Rich Text Documents(*.rtf)|*.rtf|Plan Text Documents(*.txt)|*.txt|All Files(*.*)|*.*"
        If My.Computer.Info.OSVersion >= "6.1" Then
            Me.TaskbarItemInfo.Overlay = CType(Resources("OpenOverlay"), ImageSource)
        End If
        If openfile.ShowDialog = True Then
            If My.Computer.Info.OSVersion >= "6.1" Then
                Me.TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.Normal
            End If
            Dim itemcount As Integer = 0, load As New LoadFileDialog, int As Integer = 0
            For Each s As String In openfile.FileNames
                itemcount += 1
            Next
            Me.IsEnabled = False
            load.Show()
            For Each i As String In openfile.FileNames
                Dim f As New IO.FileInfo(i)
                NewDocument(f.Name)
                SelectedDocument.Editor.LoadDocument(f.FullName)
                SelectedDocument.SetFileType(f.Extension)
                SelectedDocument.Editor.FileChanged = False
                SelectedDocument.Editor.Height = SelectedDocument.Editor.Document.PageHeight
                SelectedDocument.Editor.Width = SelectedDocument.Editor.Document.PageWidth
                SelectedDocument.Ruler.Width = SelectedDocument.Editor.Width
                Dim ch As Semagsoft.DocRuler.Ruler = SelectedDocument.Ruler.Children.Item(0)
                ch.Width = SelectedDocument.Editor.Width
                If Not My.Settings.Options_RecentFiles.Contains(SelectedDocument.Editor.DocumentName) AndAlso My.Settings.Options_RecentFiles.Count < 50 Then
                    My.Settings.Options_RecentFiles.Add(SelectedDocument.Editor.DocumentName)
                End If
                If My.Computer.Info.OSVersion >= "6.1" Then
                    TaskbarItemInfo.ProgressValue = CDbl(int) / itemcount
                    int += 1
                End If
            Next
            load.i = True
            load.Close()
            Me.IsEnabled = True
        End If
        If My.Computer.Info.OSVersion >= "6.1" Then
            Me.TaskbarItemInfo.Overlay = Nothing
            Me.TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.None
        End If
    End Sub

#End Region

#Region "Recent"

    Private Sub recentfile_click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim i As Fluent.Button = e.Source, f As New IO.FileInfo(i.Tag)
        NewDocument(f.Name)
        SelectedDocument.Editor.LoadDocument(Convert.ToString(f.FullName))
        SelectedDocument.SetFileType(f.Extension)
        UpdateButtons()
        SelectedDocument.Editor.Focus()
    End Sub

    Private Sub recentfileremove_click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim i As MenuItem = e.Source
        Dim itemtoremove As Fluent.Button = Nothing
        For Each recentdoc As Fluent.Button In RecentFilesList.Children
            If recentdoc.Tag = i.ToolTip Then
                itemtoremove = recentdoc
            End If
        Next
        Dim stringtoremove As String = Nothing
        For Each s As String In My.Settings.Options_RecentFiles
            If s = i.ToolTip Then
                stringtoremove = s
            End If
        Next
        My.Settings.Options_RecentFiles.Remove(stringtoremove)
        RecentFilesList.Children.Remove(itemtoremove)
    End Sub

#End Region

#Region "Revert"

    Private Sub RevertMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles RevertMenuItem.Click
        SelectedDocument.Editor.LoadDocument(SelectedDocument.Editor.DocumentName)
        SelectedDocument.Editor.Height = SelectedDocument.Editor.Document.PageHeight
        SelectedDocument.Editor.Width = SelectedDocument.Editor.Document.PageWidth
        SelectedDocument.Ruler.Width = SelectedDocument.Editor.Width
        Dim ch As Semagsoft.DocRuler.Ruler = SelectedDocument.Ruler.Children.Item(0)
        ch.Width = SelectedDocument.Editor.Width
        UpdateButtons()
    End Sub

#End Region

#Region "Close/Close All But This/Close All"

    Private Sub CloseMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles CloseMenuItem.Click
        If SelectedDocument.Editor.FileChanged Then
            Dim SaveDialog As New SaveFileDialog, tr As New TextRange(SelectedDocument.Editor.Document.ContentStart, SelectedDocument.Editor.Document.ContentEnd),
                fs As IO.FileStream = IO.File.Open(My.Computer.FileSystem.SpecialDirectories.Temp + "\TVPre.xaml", IO.FileMode.Create)
            SaveDialog.Owner = Me
            SaveDialog.SetFileInfo(SelectedDocument.DocName, SelectedDocument.Editor)
            Markup.XamlWriter.Save(SelectedDocument.Editor.Document, fs)
            fs.Close()
            SaveDialog.ShowDialog()
            If SaveDialog.Res = "Yes" Then
                SaveMenuItem_Click(Nothing, Nothing)
                CloseDocument(SelectedDocument)
            ElseIf SaveDialog.Res = "No" Then
                CloseDocument(SelectedDocument)
            End If
        Else
            CloseDocument(SelectedDocument)
        End If
    End Sub

    Private Sub CloseAllButThisMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles CloseAllButThisMenuItem.Click
        While TabCell.Children.Count > 1
            If TabCell.Children.Item(0) IsNot SelectedDocument Then
                Dim t As DocumentTab = TryCast(TabCell.Children.Item(0), DocumentTab)
                If t.Editor.FileChanged Then
                    Dim SaveDialog As New SaveFileDialog, tr As New TextRange(t.Editor.Document.ContentStart, t.Editor.Document.ContentEnd),
                        fs As IO.FileStream = IO.File.Open(My.Computer.FileSystem.SpecialDirectories.Temp + "\TVPre.xaml", IO.FileMode.Create)
                    SaveDialog.Owner = Me
                    SaveDialog.SetFileInfo(t.DocName, t.Editor)
                    Markup.XamlWriter.Save(t.Editor.Document, fs)
                    fs.Close()
                    SaveDialog.ShowDialog()
                    If SaveDialog.Res = "Yes" Then
                        If t.Editor.DocumentName IsNot Nothing Then
                            t.Editor.SaveDocument(t.Editor.DocumentName)
                            TabCell.Children.RemoveAt(0)
                        Else
                            Dim d As New Microsoft.Win32.SaveFileDialog
                            d.Filter = ""
                            If d.ShowDialog Then
                                Dim file As New IO.FileInfo(d.FileName)
                                t.Editor.SaveDocument(d.FileName)
                                TabCell.Children.RemoveAt(0)
                            End If
                        End If
                    ElseIf SaveDialog.Res = "No" Then
                        TabCell.Children.RemoveAt(0)
                    Else
                        Exit While
                    End If
                Else
                    TabCell.Children.RemoveAt(0)
                End If
            Else
                Dim t As DocumentTab = TryCast(TabCell.Children.Item(1), DocumentTab)
                If t.Editor.FileChanged Then
                    Dim SaveDialog As New SaveFileDialog, tr As New TextRange(t.Editor.Document.ContentStart, t.Editor.Document.ContentEnd),
                        fs As IO.FileStream = IO.File.Open(My.Computer.FileSystem.SpecialDirectories.Temp + "\TVPre.xaml", IO.FileMode.Create)
                    SaveDialog.Owner = Me
                    SaveDialog.SetFileInfo(t.DocName, t.Editor)
                    Markup.XamlWriter.Save(SelectedDocument.Editor.Document, fs)
                    fs.Close()
                    SaveDialog.ShowDialog()
                    If SaveDialog.Res = "Yes" Then
                        If t.Editor.DocumentName IsNot Nothing Then
                            t.Editor.SaveDocument(t.Editor.DocumentName)
                            TabCell.Children.RemoveAt(1)
                        Else
                            Dim d As New Microsoft.Win32.SaveFileDialog
                            d.AddExtension = True
                            d.Filter = ""
                            d.FileName = t.DocName
                            If d.ShowDialog Then
                                Dim file As New IO.FileInfo(d.FileName)
                                t.Editor.SaveDocument(d.FileName)
                                TabCell.Children.RemoveAt(1)
                            End If
                        End If
                    ElseIf SaveDialog.Res = "No" Then
                        TabCell.Children.RemoveAt(1)
                    Else
                        Exit While
                    End If
                Else
                    TabCell.Children.RemoveAt(1)
                End If
            End If
        End While
    End Sub

    Private Sub CloseAllMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles CloseAllMenuItem.Click
        While TabCell.Children.Count > 0
            If SelectedDocument.Editor.FileChanged = True Then
                Dim save As New SaveFileDialog, tr As New TextRange(SelectedDocument.Editor.Document.ContentStart, SelectedDocument.Editor.Document.ContentEnd),
                    fs As IO.FileStream = IO.File.Open(My.Computer.FileSystem.SpecialDirectories.Temp + "\TVPre.xaml", IO.FileMode.Create)
                save.SetFileInfo(SelectedDocument.DocName, SelectedDocument.Editor)
                Markup.XamlWriter.Save(SelectedDocument.Editor.Document, fs)
                fs.Close()
                save.ShowDialog()
                If save.Res = "Yes" Then
                    SaveMenuItem_Click(Me, Nothing)
                    CloseDocument(SelectedDocument)
                ElseIf save.Res = "No" Then
                    CloseDocument(SelectedDocument)
                Else
                    Exit While
                End If
            Else
                CloseDocument(SelectedDocument)
            End If
        End While
    End Sub

#End Region

#Region "Save/Save As/Save Copy/Save All"

    Private Sub SaveMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles SaveMenuItem.Click
        If SelectedDocument.Editor.DocumentName Is Nothing Then
            SaveAsMenuItem_Click(Me, Nothing)
        ElseIf My.Computer.FileSystem.FileExists(SelectedDocument.Editor.DocumentName) Then
            SelectedDocument.Editor.SaveDocument(SelectedDocument.Editor.DocumentName)
        End If
        UpdateButtons()
    End Sub

    Private Sub SaveAsMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles SaveAsMenuItem.Click
        Dim save As New Microsoft.Win32.SaveFileDialog
        save.Filter = "XAML Package(*.xamlpackage)|*.xamlpackage|FlowDocument(*.xaml)|*.xaml|OpenXML Document(*.docx)|*.docx|HTML Document(*.html;*.htm)|*.html;*.htm|Rich Text Document(*.rtf)|*.rtf|Text Document(*.txt)|*.txt|All Files(*.*)|*.*"
        If save.ShowDialog Then
            SelectedDocument.Editor.SaveDocument(save.FileName)
            SelectedDocument.DocName = New IO.FileInfo(save.FileName).Name
            SelectedDocument.Title = SelectedDocument.DocName
            SelectedDocument.SetFileType(New IO.FileInfo(save.FileName).Extension)
        End If
        UpdateButtons()
    End Sub

    Private Sub SaveCopyMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles SaveCopyMenuItem.Click
        Dim save As New Microsoft.Win32.SaveFileDialog
        save.AddExtension = True
        save.Title = "Save Copy"
        save.Filter = "XAML Package(*.xamlpackage)|*.xamlpackage|FlowDocument(*.xaml)|*.xaml|OpenXML Document(*.docx)|*.docx|HTML Document(*.html;*.htm)|*.html;*.htm|Rich Text Document(*.rtf)|*.rtf|Text Document(*.txt)|*.txt|All Files(.)|*.*"
        If save.ShowDialog Then
            Dim file As New IO.FileInfo(save.FileName), fs As IO.FileStream = IO.File.Open(save.FileName, IO.FileMode.OpenOrCreate),
                tr As New TextRange(SelectedDocument.Editor.Document.ContentStart, SelectedDocument.Editor.Document.ContentEnd), p As DocumentTab = Parent
            If file.Extension = ".xaml" Then
                Markup.XamlWriter.Save(SelectedDocument.Editor.Document, fs)
            ElseIf file.Extension = ".docx" Then
                fs.Close()
                fs = Nothing
                Dim converter As New FlowDocumenttoOpenXML
                converter.Convert(SelectedDocument.Editor.Document, save.FileName)
                converter.Close()
            ElseIf file.Extension = ".html" Then
                fs.Close()
                fs = Nothing
                Dim s As String = Markup.XamlWriter.Save(SelectedDocument.Editor.Document)
                Try
                    My.Computer.FileSystem.WriteAllText(file.FullName, HTMLConverter.HtmlFromXamlConverter.ConvertXamlToHtml(s), False)
                Catch ex As Exception
                    Dim m As New MessageBoxDialog("Error saving document!", "Error", Nothing, Nothing)
                    m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
                    m.Owner = Me
                    m.ShowDialog()
                End Try
            ElseIf file.Extension = ".rtf" Then
                tr.Save(fs, Windows.DataFormats.Rtf)
            ElseIf file.Extension = ".txt" Then
                tr.Save(fs, Windows.DataFormats.Text)
            Else
                tr.Save(fs, Windows.DataFormats.Text)
            End If
            fs.Close()
        End If
    End Sub

    Private Sub SaveAllMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles SaveAllMenuItem.Click
        Dim int As Integer = 0
        If My.Computer.Info.OSVersion >= "6.1" Then
            Me.TaskbarItemInfo.Overlay = CType(Resources("SaveOverlay"), ImageSource)
            TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.Normal
        End If
        For Each i As DocumentTab In TabCell.Children
            If i.Editor.DocumentName IsNot Nothing Then
                i.Editor.SaveDocument(i.Editor.DocumentName)
            Else
                SaveAsMenuItem_Click(i, Nothing)
            End If
            If My.Computer.Info.OSVersion >= "6.1" Then
                TaskbarItemInfo.ProgressValue = CDbl(int) / TabCell.Children.Count
                int += 1
            End If
        Next
        If My.Computer.Info.OSVersion >= "6.1" Then
            Me.TaskbarItemInfo.Overlay = Nothing
            TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.None
        End If
    End Sub

#End Region

#Region "Import"

#Region "FTP"

    Private ImportFTPIsConnected As Boolean = False, ImportFTPWorkingDir As String = "/", ImportFTPFileToLoad As String = Nothing
    Private Sub ImportFTPHostBox_TextChanged(sender As Object, e As TextChangedEventArgs) Handles ImportFTPHostBox.TextChanged, ImportFTPUsernameBox.TextChanged, ImportFTPPasswordBox.TextChanged
        Try
            If ImportFTPHostBox.Text.Length > 0 AndAlso ImportFTPUsernameBox.Text.Length > 0 AndAlso ImportFTPPasswordBox.Text.Length > 0 Then
                ImportFTPConnect.IsEnabled = True
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ImportFTPConnect_Click(sender As Object, e As RoutedEventArgs) Handles ImportFTPConnect.Click
        If ImportFTPIsConnected Then
            ImportFTPIsConnected = False
            ImportFTPHostBox.IsEnabled = True
            ImportFTPUsernameBox.IsEnabled = True
            ImportFTPPasswordBox.IsEnabled = True
            ImportFTPConnect.Header = "Connect"
            ImportFTPConnect.Icon = New BitmapImage(New Uri("pack://application:,,,/Images/Document/Export/connect16.png"))
            'Dim ttip As Fluent.ScreenTip = ImportFTPConnect.ToolTip
            'ttip.Title = "Connect"
            'ttip.Text = "Connect to FTP Server"
            ImportFTPListBox.Items.Clear()
        Else
            'Try
            Dim myFtp As New Utilities.FTP.FTPclient(ImportFTPHostBox.Text, ImportFTPUsernameBox.Text, ImportFTPPasswordBox.Text)
            ImportFTPIsConnected = True
            ImportFTPHostBox.IsEnabled = False
            ImportFTPUsernameBox.IsEnabled = False
            ImportFTPPasswordBox.IsEnabled = False
            ImportFTPConnect.Header = "Disconnect"
            ImportFTPConnect.Icon = New BitmapImage(New Uri("pack://application:,,,/Images/Document/Export/disconnect16.png"))
            'Dim ttip As Fluent.ScreenTip = ImportFTPConnect.ToolTip
            'ttip.Title = "Disconnect"
            'ttip.Text = "Disconnect from FTP Server"
            For Each i As Utilities.FTP.FTPfileInfo In myFtp.ListDirectoryDetail(ImportFTPWorkingDir)
                If Not i.NameOnly = "." Then
                    Dim item As New FTPItem
                    item.Content = i.Filename
                    item.Name = i.NameOnly
                    item.FileName = i.Filename
                    item.FullName = i.FullName
                    ImportFTPListBox.Items.Add(item)
                End If
            Next
            'Catch ex As Exception
            'Dim m As New MessageBoxDialog(ex.message, "Error", Nothing, Nothing)
            'm.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            'm.Owner = Me
            'm.ShowDialog()
            'End Try
        End If
        BackStageMenu.IsOpen = True
    End Sub

    Private Sub ImportFTPListBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ImportFTPListBox.SelectionChanged
        If ImportFTPListBox.SelectedItem IsNot Nothing Then
            Dim i As FTPItem = ImportFTPListBox.SelectedItem
            If i.IsFile Then
                ImportFTPImportButton.IsEnabled = True
            Else
                ImportFTPImportButton.IsEnabled = False
            End If
        End If
    End Sub

    Private Sub ImportFTPListBox_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles ImportFTPListBox.MouseDoubleClick
        If ImportFTPListBox.SelectedItem IsNot Nothing Then
            Dim i As FTPItem = ImportFTPListBox.SelectedItem
            If i.IsFile Then
                ImportFTPImportButton_Click(Nothing, Nothing)
            Else
                Dim myFtp As New Utilities.FTP.FTPclient(ImportFTPHostBox.Text, ImportFTPUsernameBox.Text, ImportFTPPasswordBox.Text)
                If i.IsFile Then
                    ImportFTPWorkingDir += i.FileName
                Else
                    ImportFTPWorkingDir += i.FileName + "/"
                End If
                ImportFTPListBox.Items.Clear()
                For Each i2 As Utilities.FTP.FTPfileInfo In myFtp.ListDirectoryDetail(ImportFTPWorkingDir)
                    Dim item As New FTPItem
                    item.Content = i2.Filename
                    item.Name = i2.NameOnly
                    item.FileName = i2.Filename
                    item.FullName = i2.FullName
                    If i2.FileType = Utilities.FTP.FTPfileInfo.DirectoryEntryTypes.File Then
                        item.IsFile = True
                    End If
                    ImportFTPListBox.Items.Add(item)
                Next
            End If
        End If
    End Sub

    Private Sub ImportFTPImportButton_Click(sender As Object, e As RoutedEventArgs) Handles ImportFTPImportButton.Click
        Try
            Dim i As FTPItem = ImportFTPListBox.SelectedItem
            Dim myFtp As New Utilities.FTP.FTPclient(ImportFTPHostBox.Text, ImportFTPUsernameBox.Text, ImportFTPPasswordBox.Text)
            Dim file As String = My.Computer.FileSystem.SpecialDirectories.Temp + "\" + i.FileName
            If My.Computer.FileSystem.FileExists(file) Then
                My.Computer.FileSystem.DeleteFile(file)
            End If
            myFtp.Download(ImportFTPWorkingDir + i.FileName, file)
            ImportFTPFileToLoad = file
            NewMenuItem_Click(Nothing, Nothing)
            SelectedDocument.Editor.LoadDocument(ImportFTPFileToLoad)
            SelectedDocument.DocName = "New Document"
            SelectedDocument.Editor.DocumentName = Nothing
            SelectedDocument.SetDocumentTitle("New Document")
            My.Computer.FileSystem.DeleteFile(ImportFTPFileToLoad)
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, "Error", Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
    End Sub

#End Region

    Private Sub ImportArchiveMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles ImportArchiveMenuItem.Click
        Dim import As New Microsoft.Win32.OpenFileDialog
        import.Title = "Import Archive"
        import.Filter = "Zip Archive(*.zip)|*.zip|All Files(*.*)|*.*"
        If import.ShowDialog Then
            Dim zip As New Ionic.Zip.ZipFile(import.FileName)
            zip.Password = ZipImportPasswordBox.Text
            Try
                For Each item As Ionic.Zip.ZipEntry In zip.Entries
                    item.Extract(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData + "\Semagsoft", Ionic.Zip.ExtractExistingFileAction.OverwriteSilently)
                    Dim file As New IO.FileInfo(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData + "\Semagsoft\" + item.FileName)
                    Dim tab As New DocumentTab(file.Name, Background)
                    TabCell.Children.Add(tab)
                    tab.IsSelected = True
                    SelectedDocument = tab
                    tab.Editor.LoadDocument(file.FullName)
                Next
            Catch ex As Ionic.Zip.BadPasswordException
                MessageBox.Show("Password is incorrect!")
            End Try
        End If
    End Sub

    Private Sub ImportImageMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles ImportImageMenuItem.Click
        Dim import As New Microsoft.Win32.OpenFileDialog
        import.Title = "Import Image"
        import.Filter = "Supported Images(*.bmp;*.gif;*.jpeg;*.jpg;*.png)|*.bmp;*.gif;*.jpeg;*.jpg;*.png|BMP Images(*.bmp)|*.bmp|GIF Images(*.gif)|*.gif|JPEG Images(*.jpeg;*.jpg)|*.jpeg;*.jpg|PNG Images(*.png)|*.png|All Files(*.*)|*.*"
        If import.ShowDialog Then
            Dim tab As New DocumentTab("New Document", Background), img As New Image, b As New BitmapImage
            TabCell.Children.Add(tab)
            tab.IsSelected = True
            SelectedDocument = tab
            Dim t As TextPointer = SelectedDocument.Editor.CaretPosition
            b.BeginInit()
            b.UriSource = New Uri(import.FileName)
            b.EndInit()
            img.Height = b.Height
            img.Width = b.Width
            img.Source = b
            Dim inline As New InlineUIContainer(img)
            t.Paragraph.Inlines.Add(inline)
            SelectedDocument.Editor.FileChanged = False
            UpdateButtons()
        End If
    End Sub

#End Region

#Region "Export"

#Region "Export Xps"

    Private Sub ExportXPSMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ExportXPSMenuItem.Click
        UpdateDocumentPreview()
        Dim save As New Microsoft.Win32.SaveFileDialog
        save.Title = "Export XPS"
        save.Filter = "XPS Document(*.xps)|*.xps|All Files(*.*)|*.*"
        save.AddExtension = True
        If save.ShowDialog Then
            Dim NewXpsDocument As New Xps.Packaging.XpsDocument(save.FileName, IO.FileAccess.ReadWrite)
            Dim xpsw As Xps.XpsDocumentWriter = Xps.Packaging.XpsDocument.CreateXpsDocumentWriter(NewXpsDocument)
            xpsw.Write(DocumentPreview)
            NewXpsDocument.Close()
            xpsw = Nothing
        End If
    End Sub

#End Region

#Region "Export Wordpress"

    Private Sub ExportWordpressBlogBox_TextChanged(sender As Object, e As TextChangedEventArgs) Handles ExportWordpressBlogBox.TextChanged, ExportWordpressTitleBox.TextChanged, ExportWordpressUsernameBox.TextChanged, ExportWordpressPasswordBox.TextChanged
        Try
            If ExportWordpressBlogBox.Text.Length > 0 AndAlso ExportWordpressTitleBox.Text.Length > 0 AndAlso ExportWordpressUsernameBox.Text.Length > 0 AndAlso ExportWordpressPasswordBox.Text.Length > 0 Then
                ExportWordpressExportButton.IsEnabled = True
            Else
                ExportWordpressExportButton.IsEnabled = False
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ExportWordpressExportButton_Click(sender As Object, e As RoutedEventArgs) Handles ExportWordpressExportButton.Click
        BackStageMenu.IsOpen = True
        Dim newBlogPost As AppHelper.blogInfo = Nothing
        newBlogPost.title = ExportWordpressTitleBox.Text
        Dim t As New TextRange(SelectedDocument.Editor.Document.ContentStart, SelectedDocument.Editor.Document.ContentEnd)
        newBlogPost.description = t.Text
        Dim categories As AppHelper.IgetCatList = CType(CookComputing.XmlRpc.XmlRpcProxyGen.Create(GetType(AppHelper.IgetCatList)), AppHelper.IgetCatList)
        Dim clientProtocol As CookComputing.XmlRpc.XmlRpcClientProtocol = CType(categories, CookComputing.XmlRpc.XmlRpcClientProtocol)
        clientProtocol.Url = ExportWordpressBlogBox.Text + "/xmlrpc.php"
        Dim result As String = Nothing
        result = ""
        Try
            result = categories.NewPage(1, ExportWordpressUsernameBox.Text, ExportWordpressPasswordBox.Text, newBlogPost, 1)
            Dim m As New MessageBoxDialog("Posted to Blog successfullly! Post ID : " & result, "Success", Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/info32.png"))
            m.Owner = Me
            m.ShowDialog()
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, "Error", Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
    End Sub

#End Region

#Region "Export Email"

    Private Sub ExportEmailFromBox_TextChanged(sender As Object, e As TextChangedEventArgs) Handles ExportEmailFromBox.TextChanged, ExportEmailToBox.TextChanged, ExportEmailSubjectBox.TextChanged, ExportEmailBodyBox.TextChanged, ExportEmailPasswordBox.TextChanged
        Try
            If ExportEmailFromBox.Text.Length > 0 AndAlso ExportEmailToBox.Text.Length > 0 AndAlso ExportEmailSubjectBox.Text.Length > 0 AndAlso ExportEmailBodyBox.Text.Length > 0 AndAlso ExportEmailPasswordBox.Text.Length > 0 Then
                ExportEmailExportButton.IsEnabled = True
            Else
                ExportEmailExportButton.IsEnabled = False
            End If

            If ExportEmailSMTPBox.SelectedItem Is Nothing Then
                ExportEmailExportButton.IsEnabled = False
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ExportEmailExportButton_Click(sender As Object, e As RoutedEventArgs) Handles ExportEmailExportButton.Click
        Dim client As New System.Net.Mail.SmtpClient
        client.UseDefaultCredentials = False
        client.Credentials = New Net.NetworkCredential(ExportEmailFromBox.Text, ExportEmailPasswordBox.Text)
        client.Port = Convert.ToInt32(ExportEmailPortBox.Value)
        client.Host = ExportEmailSMTPBox.Text
        client.EnableSsl = True
        Dim mail As New Net.Mail.MailMessage()
        Dim addr() As String = ExportEmailToBox.Text.Split(",")
        Try
            mail.From = New Net.Mail.MailAddress(ExportEmailFromBox.Text, ExportEmailFromBox.Text, System.Text.Encoding.UTF8)
            Dim i As Byte
            For i = 0 To addr.Length - 1
                mail.To.Add(addr(i))
            Next
            mail.Subject = ExportEmailSubjectBox.Text
            mail.Body = ExportEmailBodyBox.Text
            Dim attach As System.Net.Mail.Attachment
            If SelectedDocument.Editor.DocumentName IsNot Nothing Then
                attach = New System.Net.Mail.Attachment(SelectedDocument.Editor.DocumentName)
            Else
                If Not My.Computer.FileSystem.DirectoryExists(My.Application.Info.DirectoryPath + "\Temp") Then
                    My.Computer.FileSystem.CreateDirectory(My.Application.Info.DirectoryPath + "\Temp")
                End If
                Dim file As New IO.FileInfo(My.Application.Info.DirectoryPath + "\Temp\document.xaml"), fs As IO.FileStream = IO.File.Open(file.FullName, IO.FileMode.OpenOrCreate),
                    tr As New TextRange(SelectedDocument.Editor.Document.ContentStart, SelectedDocument.Editor.Document.ContentEnd)
                tr.Save(fs, Windows.DataFormats.Xaml)
                fs.Close()
                attach = New System.Net.Mail.Attachment(file.FullName)
            End If
            mail.Attachments.Add(attach)
            mail.DeliveryNotificationOptions = Net.Mail.DeliveryNotificationOptions.OnFailure
            client.Send(mail)
            Dim m As New MessageBoxDialog("Message Sent", "Export", Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/info32.png"))
            m.Owner = Me
            m.ShowDialog()
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
    End Sub

#End Region

#Region "Export FTP"

    Private ExportFTPFileName As String = Nothing
    Private WithEvents UploadWorker As New ComponentModel.BackgroundWorker
    Private Sub UploadWorker_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles UploadWorker.DoWork
        Try
            Dim info As Collection = TryCast(e.Argument, Collection)
            Dim ftp As New Utilities.FTP.FTPclient(TryCast(info.Item(1), String), TryCast(info.Item(2), String), TryCast(info.Item(3), String)), fileinfo As New IO.FileInfo(ExportFTPFileName)
            ftp.Upload(ExportFTPFileName, "/" + fileinfo.Name)
            Dim m As New MessageBoxDialog("Document Uploaded", "Uploaded", Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/info32.png"))
            m.Owner = Me
            m.ShowDialog()
            e.Result = True
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
            e.Result = False
        End Try
    End Sub

    Private Sub UploadWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs) Handles UploadWorker.RunWorkerCompleted
        ExportFTPUploadPanel.Visibility = Windows.Visibility.Collapsed
    End Sub

    Private Sub ExportFTPHostBox_TextChanged(sender As Object, e As TextChangedEventArgs) Handles ExportFTPHostBox.TextChanged, ExportFTPUsernameBox.TextChanged, ExportFTPPasswordBox.TextChanged
        Try
            If ExportFTPHostBox.Text.Length > 0 AndAlso ExportFTPUsernameBox.Text.Length > 0 AndAlso ExportFTPPasswordBox.Text.Length > 0 Then
                ExportFTPExportButton.IsEnabled = True
            Else
                ExportFTPExportButton.IsEnabled = False
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ExportFTPExportButton_Click(sender As Object, e As RoutedEventArgs) Handles ExportFTPExportButton.Click
        Try
            ExportFTPUploadPanel.Visibility = Windows.Visibility.Visible
            If SelectedDocument.Editor.DocumentName IsNot Nothing Then
                ExportFTPFileName = SelectedDocument.Editor.DocumentName
            Else
                If Not My.Computer.FileSystem.DirectoryExists(My.Application.Info.DirectoryPath + "\Temp") Then
                    My.Computer.FileSystem.CreateDirectory(My.Application.Info.DirectoryPath + "\Temp")
                End If
                Dim file As New IO.FileInfo(My.Application.Info.DirectoryPath + "\Temp\document.xaml"), fs As IO.FileStream = IO.File.Open(file.FullName, IO.FileMode.OpenOrCreate),
                    tr As New TextRange(SelectedDocument.Editor.Document.ContentStart, SelectedDocument.Editor.Document.ContentEnd)
                tr.Save(fs, Windows.DataFormats.Xaml)
                fs.Close()
                ExportFTPFileName = file.FullName
            End If
            If My.Computer.Network.IsAvailable Then
                Dim pram As New Collection
                pram.Add(ExportFTPHostBox.Text)
                pram.Add(ExportFTPUsernameBox.Text)
                pram.Add(ExportFTPPasswordBox.Text)
                UploadWorker.RunWorkerAsync(pram)
                BackStageMenu.IsOpen = True
            Else
                Dim m As New MessageBoxDialog("Internet not found!", "Error", Nothing, Nothing)
                m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
                m.Owner = Me
                m.ShowDialog()
            End If
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
    End Sub

#End Region

#Region "Export Archive"

    Private Sub ZipCompressionComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ZipCompressionComboBox.SelectionChanged
        If ZipCompressionComboBox.SelectedIndex = 3 Then
            ZipCompressionSlider.Visibility = Visibility.Visible
        End If
    End Sub

    Private Sub ZipEncryptionComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ZipEncryptionComboBox.SelectionChanged
        If Not ZipEncryptionComboBox.SelectedIndex = 0 Then
            ZipExportPasswordBox.Visibility = Visibility.Visible
        End If
    End Sub

    Private Sub ExportArchiveMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ExportArchiveMenuItem.Click
        Dim export As New Microsoft.Win32.SaveFileDialog
        export.Title = "Export Archive"
        export.Filter = "Zip Arcive(*.zip)|*.zip|All Files(*.*)|*.*"
        If export.ShowDialog = True Then
            Dim n As String
            If SelectedDocument.DocName = Nothing Then
                n = "document.xaml"
            Else
                n = "document.xaml"
            End If

            If Not My.Computer.FileSystem.DirectoryExists(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData + "\Semagsoft") Then
                My.Computer.FileSystem.CreateDirectory(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData + "\Semagsoft")
            End If
            If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData + "\Semagsoft\" + n) Then
                My.Computer.FileSystem.DeleteFile(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData + "\Semagsoft\" + n)
            End If
            Dim file As New IO.FileInfo(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData + "\Semagsoft\" + n), fs As IO.FileStream = IO.File.Open(file.FullName, IO.FileMode.OpenOrCreate)
            Markup.XamlWriter.Save(SelectedDocument.Editor.Document, fs)
            fs.Dispose()
            'fs = Nothing
            Dim xd As New Xml.XmlDocument()
            xd.LoadXml(My.Computer.FileSystem.ReadAllText(file.FullName))
            Dim sb As New Text.StringBuilder()
            Dim sw As New IO.StringWriter(sb)
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
            My.Computer.FileSystem.WriteAllText(file.FullName, tex, False)

            Dim zip As New Ionic.Zip.ZipFile(export.FileName)
            If ZipCompressionComboBox.SelectedIndex = 0 Then
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.Default
            ElseIf ZipCompressionComboBox.SelectedIndex = 1 Then
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression
            ElseIf ZipCompressionComboBox.SelectedIndex = 2 Then
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestSpeed
            ElseIf ZipCompressionComboBox.SelectedIndex = 3 Then
                zip.CompressionLevel = Convert.ToInt32(ZipCompressionSlider.Value)
            End If
            If ZipEncryptionComboBox.SelectedIndex = 0 Then
                zip.Encryption = Ionic.Zip.EncryptionAlgorithm.None
                zip.Password = ZipExportPasswordBox.Text
            ElseIf ZipEncryptionComboBox.SelectedIndex = 1 Then
                zip.Encryption = Ionic.Zip.EncryptionAlgorithm.PkzipWeak
                zip.Password = ZipExportPasswordBox.Text
            ElseIf ZipEncryptionComboBox.SelectedIndex = 2 Then
                zip.Encryption = Ionic.Zip.EncryptionAlgorithm.WinZipAes128
                zip.Password = ZipExportPasswordBox.Text
            ElseIf ZipEncryptionComboBox.SelectedIndex = 3 Then
                zip.Encryption = Ionic.Zip.EncryptionAlgorithm.WinZipAes256
                zip.Password = ZipExportPasswordBox.Text
            End If
            zip.AddFile(file.FullName, "")
            zip.Save()
            'My.Computer.FileSystem.DeleteFile(file.FullName)
        End If
    End Sub

#End Region

#Region "Export Image"

    Private Sub ExportImageMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ExportImageMenuItem.Click
        UpdateDocumentPreview()
        Dim save As New Microsoft.Win32.SaveFileDialog
        save.Title = "Export Image"
        save.Filter = "PNG Image(*.png)|*.png|All Files(*.*)|*.*"
        save.AddExtension = True
        If save.ShowDialog Then
            DocPreviewScrollViewer.Visibility = Windows.Visibility.Visible
            Dim image As New Image(), v As Visual = TryCast(DocumentPreview, Visual), bmp As New RenderTargetBitmap(Convert.ToInt32(DocumentPreview.Width), Convert.ToInt32(DocumentPreview.Height), 96, 96, PixelFormats.Pbgra32)
            bmp.Render(v)
            image.Source = bmp
            Using fileStream = New IO.FileStream(save.FileName, IO.FileMode.Create)
                Dim encoder As BitmapEncoder = New PngBitmapEncoder()
                encoder.Frames.Add(BitmapFrame.Create(image.Source))
                encoder.Save(fileStream)
            End Using
            DocPreviewScrollViewer.Visibility = Windows.Visibility.Collapsed
        End If
    End Sub

#End Region

#Region "Export Sound"

    Private Sub ExportSoundMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ExportSoundMenuItem.Click
        Speech.Rate = My.Settings.Options_TTSS
        Dim save As New Microsoft.Win32.SaveFileDialog
        save.Title = "Export Sound"
        save.Filter = "Wave Sound(*.wav;*.wave)|*.wav;*.wave|All Files(*.*)|*.*"
        If save.ShowDialog = True Then
            Speech.SetOutputToWaveFile(save.FileName)
            Speech.Speak(SelectedDocument.Editor.Selection.Text)
            Speech.SetOutputToDefaultAudioDevice()
            Dim m As New MessageBoxDialog("Done", "Success", Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/info32.png"))
            m.Owner = Me
            m.ShowDialog()
        End If
    End Sub

#End Region

#End Region

#Region "Properties"

    Private Sub ReadOnlyMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles ReadOnlyMenuItem.Click
        If SelectedDocument.Editor.IsReadOnly Then
            If SelectedDocument.Editor.DocumentName IsNot Nothing Then
                Dim f As New IO.FileInfo(SelectedDocument.Editor.DocumentName)
                f.IsReadOnly = False
            End If
            SelectedDocument.Editor.IsReadOnly = False
        Else
            If SelectedDocument.Editor.DocumentName IsNot Nothing Then
                Dim f As New IO.FileInfo(SelectedDocument.Editor.DocumentName)
                f.IsReadOnly = True
            End If
            SelectedDocument.Editor.IsReadOnly = True
        End If
    End Sub

#End Region

#Region "Print"

    Private Sub PrintMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles PrintMenuItem.Click
        BackStageMenu.IsOpen = True
        UpdateDocumentPreview()
        Try
            DocPreviewScrollViewer.Visibility = Visibility.Visible
            Dim pd As New PrintDialog
            pd.PrintVisual(DocumentPreview, "test")
            DocPreviewScrollViewer.Visibility = Visibility.Collapsed
        Catch ex As Exception
            Dim m As New MessageBoxDialog("No printers found!", "Warning!", Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/info32.png"))
            m.Owner = Me
            m.ShowDialog()
            DocPreviewScrollViewer.Visibility = Visibility.Collapsed
        End Try
    End Sub

    Private Sub PageSetupMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles PageSetupMenuItem.Click
        BackStageMenu.IsOpen = True
        UpdateDocumentPreview()
        Dim pd As New PrintDialog
        If pd.ShowDialog = True Then
            DocPreviewScrollViewer.Visibility = Visibility.Visible
            pd.PrintVisual(DocumentPreview, SelectedDocument.DocName)
            DocPreviewScrollViewer.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub PrintPreview_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles PrintPreview.SizeChanged
        PrintPreview.FitToMaxPagesAcross(1)
        PrintPreview.FitToMaxPagesAcrossCommand.Execute("1", PrintPreview)
    End Sub

#End Region

    Private Sub ExitMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ExitMenuItem.Click
        Close()
    End Sub

#End Region

#Region "-Edit"

#Region "Undo/Redo/Cut/Copy/Paste/Delete/Select All"

    Private Sub UndoMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles UndoButton.Click
        SelectedDocument.Editor.Focus()
        SelectedDocument.Editor.Undo()
    End Sub

    Private Sub RedoMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles RedoButton.Click
        SelectedDocument.Editor.Focus()
        SelectedDocument.Editor.Redo()
    End Sub

    Private Sub CutMenuItem_Click(ByVal sender As System.Object, ByVal e As RoutedEventArgs) Handles CutButton.Click
        SelectedDocument.Editor.Focus()
        SelectedDocument.Editor.Cut()
    End Sub

    Private Sub CutParagraphMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles CutParagraphMenuItem.Click
        Try
            If SelectedDocument.Editor.Selection.IsEmpty Then
                Dim TRange As New TextRange(SelectedDocument.Editor.CaretPosition.Paragraph.ElementStart, SelectedDocument.Editor.CaretPosition.Paragraph.ElementEnd)
                If Not TRange.IsEmpty Then
                    SelectedDocument.Editor.Selection.Select(TRange.Start, TRange.End)
                    SelectedDocument.Editor.Cut()
                End If
            Else
                EditingCommands.MoveToLineStart.Execute(Nothing, SelectedDocument.Editor)
                Dim TRange As New TextRange(SelectedDocument.Editor.CaretPosition.Paragraph.ElementStart, SelectedDocument.Editor.CaretPosition.Paragraph.ElementEnd)
                If Not TRange.IsEmpty Then
                    SelectedDocument.Editor.Selection.Select(TRange.Start, TRange.End)
                    SelectedDocument.Editor.Cut()
                End If
            End If
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
        e.Handled = True
    End Sub

    Private Sub CutLineMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles CutLineMenuItem.Click
        EditingCommands.MoveToLineStart.Execute(Nothing, SelectedDocument.Editor)
        EditingCommands.SelectToLineEnd.Execute(Nothing, SelectedDocument.Editor)
        SelectedDocument.Editor.Cut()
        e.Handled = True
    End Sub

    Private Sub CopyMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles CopyButton.Click
        SelectedDocument.Editor.Focus()
        SelectedDocument.Editor.Copy()
    End Sub

    Private Sub CopyParagraphMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles CopyParagraphMenuItem.Click
        Try
            If SelectedDocument.Editor.Selection.IsEmpty Then
                Dim TRange As New TextRange(SelectedDocument.Editor.CaretPosition.Paragraph.ElementStart, SelectedDocument.Editor.CaretPosition.Paragraph.ElementEnd)
                If Not TRange.IsEmpty Then
                    SelectedDocument.Editor.Selection.Select(TRange.Start, TRange.End)
                    SelectedDocument.Editor.Copy()
                End If
            Else
                EditingCommands.MoveToLineStart.Execute(Nothing, SelectedDocument.Editor)
                Dim TRange As New TextRange(SelectedDocument.Editor.CaretPosition.Paragraph.ElementStart, SelectedDocument.Editor.CaretPosition.Paragraph.ElementEnd)
                If Not TRange.IsEmpty Then
                    SelectedDocument.Editor.Selection.Select(TRange.Start, TRange.End)
                    SelectedDocument.Editor.Copy()
                End If
            End If
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
        e.Handled = True
    End Sub

    Private Sub CopyLineMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles CopyLineMenuItem.Click
        EditingCommands.MoveToLineStart.Execute(Nothing, SelectedDocument.Editor)
        EditingCommands.SelectToLineEnd.Execute(Nothing, SelectedDocument.Editor)
        SelectedDocument.Editor.Copy()
        e.Handled = True
    End Sub

#Region "Paste"

    Private Sub PasteMenuItem_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles PasteButton.Click
        SelectedDocument.Editor.Focus()
        SelectedDocument.Editor.Paste()
    End Sub

    Private Sub PasteTextButton_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles PasteTextButton.Click
        SelectedDocument.Editor.CaretPosition.InsertTextInRun(Clipboard.GetText)
        e.Handled = True
    End Sub

    Private Sub PasteImageMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles PasteImageMenuItem.Click
        Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, img As New Image, b As New BitmapImage
        Dim ms As New IO.MemoryStream()
        ' no using here! BitmapImage will dispose the stream after loading
        My.Computer.Clipboard.GetImage.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp)
        b.BeginInit()
        b.CacheOption = BitmapCacheOption.OnLoad
        b.StreamSource = ms
        b.EndInit()
        img.Tag = New Thickness(0, 1, 1, 0)
        Dim trans As New TransformGroup
        trans.Children.Add(New RotateTransform(0))
        trans.Children.Add(New ScaleTransform(1, 1))
        img.LayoutTransform = trans
        img.Stretch = Stretch.Fill
        img.Height = b.Height
        img.Width = b.Width
        img.Source = b
        Dim inline As New InlineUIContainer(img)
        If t.Parent.GetType Is GetType(TableCell) Then
            Dim cell As TableCell = t.Parent
            cell.Blocks.Add(New Paragraph(inline))
        Else
            t.Paragraph.Inlines.Add(inline)
        End If
        e.Handled = True
    End Sub

#End Region

    Private Sub DeleteMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles DeleteButton.Click
        EditingCommands.Delete.Execute(Nothing, SelectedDocument.Editor)
    End Sub

    Private Sub DeleteParagraphMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles DeleteParagraphMenuItem.Click
        Try
            If SelectedDocument.Editor.Selection.IsEmpty Then
                Dim TRange As New TextRange(SelectedDocument.Editor.CaretPosition.Paragraph.ElementStart, SelectedDocument.Editor.CaretPosition.Paragraph.ElementEnd)
                If Not TRange.IsEmpty Then
                    SelectedDocument.Editor.Selection.Select(TRange.Start, TRange.End)
                    EditingCommands.Delete.Execute(Nothing, SelectedDocument.Editor)
                End If
            Else
                EditingCommands.MoveToLineStart.Execute(Nothing, SelectedDocument.Editor)
                Dim TRange As New TextRange(SelectedDocument.Editor.CaretPosition.Paragraph.ElementStart, SelectedDocument.Editor.CaretPosition.Paragraph.ElementEnd)
                If Not TRange.IsEmpty Then
                    SelectedDocument.Editor.Selection.Select(TRange.Start, TRange.End)
                    EditingCommands.Delete.Execute(Nothing, SelectedDocument.Editor)
                End If
            End If
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
        e.Handled = True
    End Sub

    Private Sub DeleteLineMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles DeleteLineMenuItem.Click
        EditingCommands.MoveToLineStart.Execute(Nothing, SelectedDocument.Editor)
        EditingCommands.SelectToLineEnd.Execute(Nothing, SelectedDocument.Editor)
        EditingCommands.Delete.Execute(Nothing, SelectedDocument.Editor)
        e.Handled = True
    End Sub

    Private Sub SelectAllMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles SelectAllButton.Click
        SelectedDocument.Editor.Focus()
        SelectedDocument.Editor.SelectAll()
    End Sub

    Private Sub SelectParagraphMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles SelectParagraphMenuItem.Click
        Try
            If SelectedDocument.Editor.Selection.IsEmpty Then
                Dim TRange As New TextRange(SelectedDocument.Editor.CaretPosition.Paragraph.ElementStart, SelectedDocument.Editor.CaretPosition.Paragraph.ElementEnd)
                If Not TRange.IsEmpty Then
                    SelectedDocument.Editor.Selection.Select(TRange.Start, TRange.End)
                End If
            Else
                EditingCommands.MoveToLineStart.Execute(Nothing, SelectedDocument.Editor)
                Dim TRange As New TextRange(SelectedDocument.Editor.CaretPosition.Paragraph.ElementStart, SelectedDocument.Editor.CaretPosition.Paragraph.ElementEnd)
                If Not TRange.IsEmpty Then
                    SelectedDocument.Editor.Selection.Select(TRange.Start, TRange.End)
                End If
            End If
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
        e.Handled = True
    End Sub

    Private Sub SelectLineMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles SelectLineMenuItem.Click
        EditingCommands.MoveToLineStart.Execute(Nothing, SelectedDocument.Editor)
        EditingCommands.SelectToLineEnd.Execute(Nothing, SelectedDocument.Editor)
        e.Handled = True
    End Sub

#End Region

#Region "Find/Replace/Go To"

    Private Sub FindButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles FindButton.Click
        SearchPanel.Show()
        SearchPanel.IsActive = True
        SearchPanel_FindTextBox.Focus()
    End Sub

    Private Sub ReplaceButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ReplaceButton.Click
        SearchPanel.Show()
        SearchPanel.IsActive = True
        SearchPanel_ReplaceTextBox.Focus()
    End Sub

    Private Sub GoToButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles GoToButton.Click
        SearchPanel.Show()
        SearchPanel.IsActive = True
        SearchPanel_LineSpinner.Focus()
    End Sub

#End Region

    Private Sub UppercaseMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles UppercaseMenuItem.Click
        SelectedDocument.Editor.Selection.Text = SelectedDocument.Editor.Selection.Text.ToUpper
    End Sub

    Private Sub LowercaseMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles LowercaseMenuItem.Click
        SelectedDocument.Editor.Selection.Text = SelectedDocument.Editor.Selection.Text.ToLower
    End Sub

#End Region

#Region "--ViewMenu"

    Private Sub HRulerButton_Checked(sender As Object, e As RoutedEventArgs) Handles HRulerButton.Checked
        For Each t As DocumentTab In TabCell.Children
            t.Ruler.Visibility = Visibility.Visible
        Next
        My.Settings.MainWindow_ShowRuler = True
        UpdateUI()
    End Sub

    Private Sub HRulerButton_Unchecked(sender As Object, e As RoutedEventArgs) Handles HRulerButton.Unchecked
        For Each t As DocumentTab In TabCell.Children
            t.Ruler.Visibility = Visibility.Collapsed
        Next
        My.Settings.MainWindow_ShowRuler = False
        UpdateUI()
    End Sub

    Private Sub StatusbarButton_Checked(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles StatusbarButton.Checked
        MainGrid.RowDefinitions.Item(2).Height = New GridLength(22)
        StatusBar.Visibility = Visibility.Visible
    End Sub

    Private Sub StatusbarButton_Unchecked(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles StatusbarButton.Unchecked
        MainGrid.RowDefinitions.Item(2).Height = New GridLength(0)
    End Sub

    Private Sub ZoomInButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ZoomInButton.Click
        zoomSlider.Value += 10
    End Sub

    Private Sub ZoomOutButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ZoomOutButton.Click
        zoomSlider.Value -= 10
    End Sub

    Private Sub ResetZoomButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ResetZoomButton.Click
        zoomSlider.Value = 100
    End Sub

    Private fullscreensetting As WindowState = WindowState.Normal, IsFullscreen As Boolean = False
    Private Sub FullscreenMenuItem_Checked(sender As Object, e As RoutedEventArgs) Handles FullscreenMenuItem.Checked
        fullscreensetting = WindowState
        If WindowState = WindowState.Maximized Then
            WindowState = WindowState.Normal
        End If
        WindowStyle = WindowStyle.None
        Topmost = True
        WindowState = WindowState.Maximized
        IsFullscreen = True
    End Sub

    Private Sub FullscreenMenuItem_Unchecked(sender As Object, e As RoutedEventArgs) Handles FullscreenMenuItem.Unchecked
        Topmost = False
        WindowStyle = WindowStyle.SingleBorderWindow
        WindowState = WindowState.Normal
        WindowState = fullscreensetting
        IsFullscreen = False
    End Sub

#End Region

#Region "Insert"

#Region "Tables"

    Private Sub TableGrid_Click(y As Integer, x As Integer) Handles TableGrid.Click
        Dim t As New Table, int As Integer = Convert.ToInt32(y), int2 As Integer = Convert.ToInt32(x)
        While Not int = 0
            Dim trg As New TableRowGroup, tr As New TableRow
            While Not int2 = 0
                Dim tc As New TableCell
                tc.BorderBrush = Brushes.Black
                tc.BorderThickness = New Thickness(1, 1, 1, 1)
                tr.Cells.Add(tc)
                int2 -= 1
            End While
            int2 = Convert.ToInt32(x)
            trg.Rows.Add(tr)
            t.RowGroups.Add(trg)
            int -= 1
        End While
        t.BorderThickness = New Thickness(1, 1, 1, 1)
        If SelectedDocument.Editor.CaretPosition.Paragraph IsNot Nothing Then
            Dim tc As TableCell = TryCast(SelectedDocument.Editor.CaretPosition.Paragraph.Parent, TableCell)
            If tc IsNot Nothing Then
                tc.Blocks.InsertBefore(SelectedDocument.Editor.CaretPosition.Paragraph, t)
            Else
                Dim listitem As ListItem = TryCast(SelectedDocument.Editor.CaretPosition.Paragraph.Parent, ListItem)
                If listitem IsNot Nothing Then
                    Dim list As List = TryCast(listitem.Parent, List)
                    If list IsNot Nothing Then
                        SelectedDocument.Editor.Document.Blocks.InsertAfter(list, t)
                    End If
                Else
                    SelectedDocument.Editor.Document.Blocks.InsertBefore(SelectedDocument.Editor.CaretPosition.Paragraph, t)
                End If
            End If
        Else
            Dim tc As TableCell = TryCast(SelectedDocument.Editor.CaretPosition.Parent, TableCell)
            If tc IsNot Nothing Then
                tc.Blocks.Add(t)
            Else
                Dim tr As TableRow = TryCast(SelectedDocument.Editor.CaretPosition.Parent, TableRow)
                If tr IsNot Nothing Then
                    Dim trg As TableRowGroup = tr.Parent, table As Table = trg.Parent
                    SelectedDocument.Editor.Document.Blocks.InsertAfter(TryCast(table, Block), t)
                Else
                    Dim flow As FlowDocument = TryCast(SelectedDocument.Editor.CaretPosition.Parent, FlowDocument)
                    If flow IsNot Nothing Then
                        SelectedDocument.Editor.Document.Blocks.InsertBefore(SelectedDocument.Editor.Document.Blocks.FirstBlock, t)
                    End If
                End If
            End If
        End If
        UpdateButtons()
    End Sub

    Private Sub TableMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles InsertTableMenuItem.Click
        Dim tableDialog As New TableDialog
        tableDialog.Owner = Me
        tableDialog.ShowDialog()
        If tableDialog.Res = "OK" Then
            Dim t As New Table, int As Integer = Convert.ToInt32(tableDialog.RowsTextBox.Value), int2 As Integer = Convert.ToInt32(tableDialog.CellsTextBox.Value)
            While Not int = 0
                Dim trg As New TableRowGroup, tr As New TableRow
                While Not int2 = 0
                    Dim tc As New TableCell
                    tc.Background = tableDialog.CellBackgroundColor
                    tc.BorderBrush = tableDialog.CellBorderColor
                    tc.BorderThickness = New Thickness(1, 1, 1, 1)
                    tr.Cells.Add(tc)
                    int2 -= 1
                End While
                int2 = Convert.ToInt32(tableDialog.CellsTextBox.Value)
                trg.Rows.Add(tr)
                t.RowGroups.Add(trg)
                int -= 1
            End While
            t.Background = tableDialog.BackgroundColor
            t.BorderBrush = tableDialog.BorderColor
            t.BorderThickness = New Thickness(1, 1, 1, 1)
            If SelectedDocument.Editor.CaretPosition.Paragraph IsNot Nothing Then
                Dim tc As TableCell = TryCast(SelectedDocument.Editor.CaretPosition.Paragraph.Parent, TableCell)
                If tc IsNot Nothing Then
                    tc.Blocks.InsertBefore(SelectedDocument.Editor.CaretPosition.Paragraph, t)
                Else
                    Dim listitem As ListItem = TryCast(SelectedDocument.Editor.CaretPosition.Paragraph.Parent, ListItem)
                    If listitem IsNot Nothing Then
                        Dim list As List = TryCast(listitem.Parent, List)
                        If list IsNot Nothing Then
                            SelectedDocument.Editor.Document.Blocks.InsertAfter(list, t)
                        End If
                    Else
                        SelectedDocument.Editor.Document.Blocks.InsertBefore(SelectedDocument.Editor.CaretPosition.Paragraph, t)
                    End If
                End If
            Else
                Dim tc As TableCell = TryCast(SelectedDocument.Editor.CaretPosition.Parent, TableCell)
                If tc IsNot Nothing Then
                    tc.Blocks.Add(t)
                Else
                    Dim tr As TableRow = TryCast(SelectedDocument.Editor.CaretPosition.Parent, TableRow)
                    If tr IsNot Nothing Then
                        Dim trg As TableRowGroup = tr.Parent, table As Table = trg.Parent
                        SelectedDocument.Editor.Document.Blocks.InsertAfter(TryCast(table, Block), t)
                    Else
                        Dim flow As FlowDocument = TryCast(SelectedDocument.Editor.CaretPosition.Parent, FlowDocument)
                        If flow IsNot Nothing Then
                            SelectedDocument.Editor.Document.Blocks.InsertBefore(SelectedDocument.Editor.Document.Blocks.FirstBlock, t)
                        End If
                    End If
                End If
            End If
            UpdateButtons()
        End If
    End Sub

#End Region

#Region "Image"

    Private Sub ImageButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ImageButton.Click
        Dim open As New Microsoft.Win32.OpenFileDialog
        open.Multiselect = True
        open.Title = "Add Images"
        open.Filter = "Supported Images(*.bmp;*.gif;*.jpeg;*.jpg;*.png)|*.bmp;*.gif;*.jpeg;*.jpg;*.png|Bitmap Images(*.bmp)|*.bmp|GIF Images(*.gif)|*.gif|JPEG Images(*.jpeg;*.jpg)|*.jpeg;*.jpg|PNG Images(*.png)|*.png|All Files(*.*)|*.*"
        If open.ShowDialog = True Then
            SelectedDocument.Editor.Focus()
            For Each i As String In open.FileNames
                Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, img As New Image, b As New BitmapImage
                b.BeginInit()
                b.UriSource = New Uri(i)
                b.EndInit()
                img.Tag = New Thickness(0, 1, 1, 0)
                Dim trans As New TransformGroup
                trans.Children.Add(New RotateTransform(0))
                trans.Children.Add(New ScaleTransform(1, 1))
                img.LayoutTransform = trans
                img.Stretch = Stretch.Fill
                img.Height = b.Height
                img.Width = b.Width
                img.Source = b
                Dim inline As New InlineUIContainer(img)
                If SelectedDocument.Editor.CaretPosition.Paragraph IsNot Nothing Then
                    t.Paragraph.Inlines.Add(inline)
                Else
                    Dim tc As TableCell = TryCast(SelectedDocument.Editor.CaretPosition.Parent, TableCell)
                    If tc IsNot Nothing Then
                        tc.Blocks.Add(New Paragraph(inline))
                    Else
                        Dim tr As TableRow = TryCast(SelectedDocument.Editor.CaretPosition.Parent, TableRow)
                        If tr IsNot Nothing Then
                            Dim trg As TableRowGroup = tr.Parent, table As Table = trg.Parent
                            SelectedDocument.Editor.Document.Blocks.InsertAfter(TryCast(table, Block), New Paragraph(inline))
                        Else
                            Dim flow As FlowDocument = TryCast(SelectedDocument.Editor.CaretPosition.Parent, FlowDocument)
                            If flow IsNot Nothing Then
                                SelectedDocument.Editor.Document.Blocks.InsertBefore(SelectedDocument.Editor.Document.Blocks.FirstBlock, New Paragraph(inline))
                            End If
                        End If
                    End If
                End If
            Next
            UpdateSelected()
        End If
    End Sub

#End Region

#Region "Object"

    Public Sub InsertObject(type As String)
        If SelectedDocument.Editor.CaretPosition.Paragraph IsNot Nothing Then
            If type = "button" Then
                Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New Button, i As New InlineUIContainer(b)
                b.Width = 72
                b.Height = 24
                b.Content = "Text"
                t.Paragraph.Inlines.Add(i)
            ElseIf type = "radiobutton" Then
                Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New RadioButton, i As New InlineUIContainer(b)
                b.Width = 72
                b.Height = 24
                b.Content = "Text"
                t.Paragraph.Inlines.Add(i)
            ElseIf type = "checkbox" Then
                Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New CheckBox, i As New InlineUIContainer(b)
                b.Width = 72
                b.Height = 24
                b.Content = "Text"
                t.Paragraph.Inlines.Add(i)
            ElseIf type = "textblock" Then
                Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New TextBlock, i As New InlineUIContainer(b)
                b.Width = 72
                b.Height = 24
                b.Text = "Text"
                t.Paragraph.Inlines.Add(i)
            End If
        Else
            Dim tc As TableCell = TryCast(SelectedDocument.Editor.CaretPosition.Parent, TableCell)
            If tc IsNot Nothing Then
                If type = "button" Then
                    Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New Button, i As New InlineUIContainer(b)
                    'b.Width = od.OW
                    ' b.Height = od.OH
                    ' b.Content = od.OT
                    tc.Blocks.Add(New Paragraph(i))
                ElseIf type = "radiobutton" Then
                    Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New RadioButton, i As New InlineUIContainer(b)
                    ' b.Width = od.OW
                    ' b.Height = od.OH
                    ' b.Content = od.OT
                    tc.Blocks.Add(New Paragraph(i))
                ElseIf type = "checkbox" Then
                    Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New CheckBox, i As New InlineUIContainer(b)
                    ' b.Width = od.OW
                    ' b.Height = od.OH
                    ' b.Content = od.OT
                    tc.Blocks.Add(New Paragraph(i))
                ElseIf type = "textblock" Then
                    Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New TextBlock, i As New InlineUIContainer(b)
                    'b.Width = od.OW
                    'b.Height = od.OH
                    'b.Text = od.OT
                    tc.Blocks.Add(New Paragraph(i))
                End If
            Else
                Dim tr As TableRow = TryCast(SelectedDocument.Editor.CaretPosition.Parent, TableRow)
                If tr IsNot Nothing Then
                    Dim trg As TableRowGroup = tr.Parent, table As Table = trg.Parent
                    If type = "button" Then
                        Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New Button, i As New InlineUIContainer(b)
                        'b.Width = od.OW
                        'b.Height = od.OH
                        'b.Content = od.OT
                        SelectedDocument.Editor.Document.Blocks.InsertAfter(TryCast(table, Block), New Paragraph(i))
                    ElseIf type = "radiobutton" Then
                        Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New RadioButton, i As New InlineUIContainer(b)
                        'b.Width = od.OW
                        'b.Height = od.OH
                        'b.Content = od.OT
                        SelectedDocument.Editor.Document.Blocks.InsertAfter(TryCast(table, Block), New Paragraph(i))
                    ElseIf type = "checkbox" Then
                        Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New CheckBox, i As New InlineUIContainer(b)
                        'b.Width = od.OW
                        'b.Height = od.OH
                        'b.Content = od.OT
                        SelectedDocument.Editor.Document.Blocks.InsertAfter(TryCast(table, Block), New Paragraph(i))
                    ElseIf type = "textblock" Then
                        Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New TextBlock, i As New InlineUIContainer(b)
                        'b.Width = od.OW
                        'b.Height = od.OH
                        'b.Text = od.OT
                        SelectedDocument.Editor.Document.Blocks.InsertAfter(TryCast(table, Block), New Paragraph(i))
                    End If
                Else
                    Dim fd As FlowDocument = TryCast(SelectedDocument.Editor.CaretPosition.Parent, FlowDocument)
                    If fd IsNot Nothing Then
                        If type = "button" Then
                            Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New Button, i As New InlineUIContainer(b)
                            'b.Width = od.OW
                            ' b.Height = od.OH
                            'b.Content = od.OT
                            SelectedDocument.Editor.Document.Blocks.InsertBefore(SelectedDocument.Editor.Document.Blocks.FirstBlock, New Paragraph(i))
                        ElseIf type = "radiobutton" Then
                            Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New RadioButton, i As New InlineUIContainer(b)
                            'b.Width = od.OW
                            'b.Height = od.OH
                            'b.Content = od.OT
                            SelectedDocument.Editor.Document.Blocks.InsertBefore(SelectedDocument.Editor.Document.Blocks.FirstBlock, New Paragraph(i))
                        ElseIf type = "checkbox" Then
                            Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New CheckBox, i As New InlineUIContainer(b)
                            'b.Width = od.OW
                            'b.Height = od.OH
                            'b.Content = od.OT
                            SelectedDocument.Editor.Document.Blocks.InsertBefore(SelectedDocument.Editor.Document.Blocks.FirstBlock, New Paragraph(i))
                        ElseIf type = "textblock" Then
                            Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, b As New TextBlock, i As New InlineUIContainer(b)
                            'b.Width = od.OW
                            'b.Height = od.OH
                            'b.Text = od.OT
                            SelectedDocument.Editor.Document.Blocks.InsertBefore(SelectedDocument.Editor.Document.Blocks.FirstBlock, New Paragraph(i))
                        End If
                    End If
                End If
            End If
        End If
        UpdateButtons()
    End Sub

    Private Sub InsertObject_Button_Click(sender As Object, e As RoutedEventArgs) Handles InsertObject_Button.Click
        InsertObject("button")
    End Sub

    Private Sub InsertObject_CheckBox_Click(sender As Object, e As RoutedEventArgs) Handles InsertObject_CheckBox.Click
        InsertObject("checkbox")
    End Sub

    Private Sub InsertObject_RadioButton_Click(sender As Object, e As RoutedEventArgs) Handles InsertObject_RadioButton.Click
        InsertObject("radiobutton")
    End Sub

    Private Sub InsertObject_TextBlock_Click(sender As Object, e As RoutedEventArgs) Handles InsertObject_TextBlock.Click
        InsertObject("textblock")
    End Sub

#End Region

#Region "Shape"

    Private Sub ShapeButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ShapeButton.Click
        Dim sd As New ShapeDialog
        sd.Owner = Me
        sd.ShowDialog()
        If sd.Res = "OK" Then
            Dim s As Shape = Nothing
            If sd.TypeComboBox.SelectedIndex = 0 Then
                s = New Ellipse
            ElseIf sd.TypeComboBox.SelectedIndex = 1 Then
                s = New Rectangle
            End If
            s.Height = sd.Shape.RenderSize.Height
            s.Width = sd.Shape.RenderSize.Width
            s.Stroke = sd.Shape.Stroke
            s.StrokeThickness = sd.Shape.StrokeThickness
            Dim i As New InlineUIContainer
            i.Child = s
            SelectedDocument.Editor.CaretPosition.Paragraph.Inlines.Add(i)
            UpdateButtons()
        End If
    End Sub

#End Region

#Region "Chart"

    Private Sub ChartButton_Click(sender As Object, e As RoutedEventArgs) Handles ChartButton.Click
        Dim d As New ChartDialog
        d.Owner = Me
        d.ShowDialog()
        If d.Res = "OK" Then
            Dim render As New RenderTargetBitmap(d.ChartWidth.Value, d.ChartHight.Value, 96, 96, PixelFormats.Default)
            render.Render(d.PreviewChart)
            Dim bsource As BitmapSource = render
            Dim img As New Image
            img.Source = bsource
            img.Height = d.ChartHight.Value
            img.Width = d.ChartWidth.Value
            Dim i As New InlineUIContainer
            i.Child = img
            SelectedDocument.Editor.CaretPosition.Paragraph.Inlines.Add(i)
            UpdateButtons()
        End If
    End Sub

#End Region

#Region "Link"

    Private Sub LinkMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles LinkButton.Click
        Dim l As New LinkDialog
        l.Owner = Me
        l.ShowDialog()
        If Not l.Link = Nothing Then
            Try
                SelectedDocument.Editor.Focus()
                Dim li As New Hyperlink(SelectedDocument.Editor.CaretPosition.DocumentStart, SelectedDocument.Editor.CaretPosition.DocumentEnd)
                Dim u As New Uri(l.Link)
                li.NavigateUri = u
                UpdateButtons()
            Catch ex As Exception
                Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
                m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
                m.Owner = Me
                m.ShowDialog()
            End Try
        End If
    End Sub

#End Region

#Region "FlowDocument"

    Private Sub InsertFlowDocumentButton_Click(sender As Object, e As RoutedEventArgs) Handles InsertFlowDocumentButton.Click
        Dim open As New Microsoft.Win32.OpenFileDialog
        open.Title = "Insert FlowDocument"
        open.Filter = "FlowDocument(*.xaml)|*.xaml"
        If open.ShowDialog Then
            Dim fs As IO.FileStream = IO.File.Open(open.FileName, IO.FileMode.Open, IO.FileAccess.Read)
            Dim content As FlowDocument = TryCast(Markup.XamlReader.Load(fs), FlowDocument)
            fs.Close()
            fs = Nothing
            For Each b As Block In content.Blocks
                Dim xaml As String = Markup.XamlWriter.Save(b)
                Dim newblock As Block = TryCast(Markup.XamlReader.Load(New Xml.XmlTextReader(New IO.StringReader(xaml))), Block)
                SelectedDocument.Editor.Document.Blocks.InsertBefore(SelectedDocument.Editor.CaretPosition.Paragraph, newblock)
            Next
        End If
    End Sub

#End Region

#Region "Rich Text Document"

    Private Sub InsertRichTextDocumentButton_Click(sender As Object, e As RoutedEventArgs) Handles InsertRichTextDocumentButton.Click
        Dim open As New Microsoft.Win32.OpenFileDialog
        open.Title = "Insert Rich Text Document"
        open.Filter = "Rich Text Document(*.rtf)|*.rtf"
        If open.ShowDialog Then
            Dim st As New IO.FileStream(open.FileName, IO.FileMode.Open, IO.FileAccess.Read)
            SelectedDocument.Editor.Selection.Load(st, DataFormats.Rtf)
            st.Close()
        End If
    End Sub

#End Region

#Region "Text File"

    Private Sub TextFileButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles TextFileButton.Click
        Dim o As New Microsoft.Win32.OpenFileDialog
        o.Title = "Insert Text File"
        o.Filter = "Text Document(*.txt)|*.txt|All Files(*.*)|*.*"
        If o.ShowDialog Then
            SelectedDocument.Editor.CaretPosition.InsertTextInRun(My.Computer.FileSystem.ReadAllText(o.FileName))
            UpdateButtons()
        End If
    End Sub

#End Region

    Private Sub EmoticonGrid_Click(img As BitmapSource) Handles EmoticonGrid.Click
        SelectedDocument.Editor.Focus()
        Dim t As TextPointer = SelectedDocument.Editor.CaretPosition, image As New Image
        image.Tag = New Thickness(0, 1, 1, 0)
        Dim trans As New TransformGroup
        trans.Children.Add(New RotateTransform(0))
        trans.Children.Add(New ScaleTransform(1, 1))
        image.LayoutTransform = trans
        image.Stretch = Stretch.Fill
        image.Height = img.Height
        image.Width = img.Width
        image.Source = img
        Dim inline As New InlineUIContainer(image)
        If SelectedDocument.Editor.CaretPosition.Paragraph IsNot Nothing Then
            t.Paragraph.Inlines.Add(inline)
        Else
            Dim tc As TableCell = TryCast(SelectedDocument.Editor.CaretPosition.Parent, TableCell)
            If tc IsNot Nothing Then
                tc.Blocks.Add(New Paragraph(inline))
            Else
                Dim tr As TableRow = TryCast(SelectedDocument.Editor.CaretPosition.Parent, TableRow)
                If tr IsNot Nothing Then
                    Dim trg As TableRowGroup = tr.Parent, table As Table = trg.Parent
                    SelectedDocument.Editor.Document.Blocks.InsertAfter(TryCast(table, Block), New Paragraph(inline))
                Else
                    Dim flow As FlowDocument = TryCast(SelectedDocument.Editor.CaretPosition.Parent, FlowDocument)
                    If flow IsNot Nothing Then
                        SelectedDocument.Editor.Document.Blocks.InsertBefore(SelectedDocument.Editor.Document.Blocks.FirstBlock, New Paragraph(inline))
                    End If
                End If
            End If
        End If
        UpdateSelected()
    End Sub

#Region "Symbol"

    Private Sub InsertSymbolGallery_Click(symbol As String) Handles InsertSymbolGallery.Click
        SelectedDocument.Editor.CaretPosition.InsertTextInRun(symbol)
        UpdateButtons()
    End Sub

#End Region

#Region "Video"

    Private Sub VideoButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles VideoButton.Click
        Dim dialog As New Microsoft.Win32.OpenFileDialog
        dialog.Title = "Insert Video"
        dialog.Filter = "Supported Videos(*.avi;*.mpeg;*.mpg;*.wmv)|*.avi;*.mpeg;*.mpg;*.wmv|AVI Videos(*.avi)|*.avi|MPEG Videos(*.mpeg;*.mpg)|*.mpeg;*.mpg|WMV Videos(*.wmv)|*.wmv|All Files(*.*)|*.*"
        If dialog.ShowDialog Then
            Try
                Dim m As New MediaElement
                m.Width = 320
                m.Height = 240
                m.Source = New Uri(dialog.FileName)
                m.LoadedBehavior = MediaState.Manual
                Dim i As New InlineUIContainer
                i.Child = m
                SelectedDocument.Editor.CaretPosition.Paragraph.Inlines.Add(i)
                UpdateSelected()
            Catch ex As Exception
                Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
                m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
                m.Owner = Me
                m.ShowDialog()
            End Try
        End If
    End Sub

#End Region

#Region "Horizontal Line"

    Private Sub HorizontalLineButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles HorizontalLineButton.Click
        Dim d As New InsertLineDialog
        d.Owner = Me
        d.ShowDialog()
        If d.Res = "OK" Then
            Dim line As New Separator, inline As New InlineUIContainer
            line.Background = Brushes.Gray
            line.Width = d.h
            line.HorizontalAlignment = Windows.HorizontalAlignment.Stretch
            inline.Child = line
            SelectedDocument.Editor.CaretPosition.Paragraph.Inlines.Add(inline)
            UpdateButtons()
        End If
    End Sub

#End Region

#Region "Header/Footer"

    Private Sub HeaderButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles HeaderButton.Click
        SelectedDocument.Editor.Document.Blocks.FirstBlock.ContentStart.InsertParagraphBreak()
        SelectedDocument.Editor.Document.Blocks.FirstBlock.ContentStart.InsertTextInRun("Header")
        UpdateButtons()
    End Sub

    Private Sub FooterButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles FooterButton.Click
        SelectedDocument.Editor.Document.ContentEnd.InsertParagraphBreak()
        SelectedDocument.Editor.Document.ContentEnd.InsertTextInRun("Footer")
        UpdateButtons()
    End Sub

#End Region

#Region "Date/Time"

    Private Sub DateMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles DateButton.Click
        SelectedDocument.Editor.CaretPosition.InsertTextInRun(Date.Now.ToString("M/dd/yyyy"))
        UpdateButtons()
    End Sub

    Private Sub MoreDateMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles MoreDateMenuItem.Click
        Dim d As New DateDialog
        d.Owner = Me
        d.ShowDialog()
        If d.Res = "OK" Then
            SelectedDocument.Editor.CaretPosition.InsertTextInRun(d.ListBox1.SelectedItem)
            UpdateButtons()
        End If
        e.Handled = True
    End Sub

    Private Sub TimeMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles TimeButton.Click
        SelectedDocument.Editor.CaretPosition.InsertTextInRun(Date.Now.ToString("h:mm tt"))
        UpdateButtons()
    End Sub

    Private Sub MoreTimeButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles MoreTimeMenuItem.Click
        Dim d As New TimeDialog
        d.Owner = Me
        d.ShowDialog()
        If d.Res = "OK" Then
            If d.RadioButton12.IsChecked Then
                If d.AMPMCheckBox.IsChecked Then
                    If d.SecCheckBox.IsChecked Then
                        SelectedDocument.Editor.CaretPosition.InsertTextInRun(Date.Now.ToString("h:mm:ss tt"))
                    Else
                        SelectedDocument.Editor.CaretPosition.InsertTextInRun(Date.Now.ToString("h:mm tt"))
                    End If
                Else
                    If d.SecCheckBox.IsChecked Then
                        SelectedDocument.Editor.CaretPosition.InsertTextInRun(Date.Now.ToString("h:mm:ss"))
                    Else
                        SelectedDocument.Editor.CaretPosition.InsertTextInRun(Date.Now.ToString("h:mm"))
                    End If
                End If
            Else
                If d.SecCheckBox.IsChecked Then
                    SelectedDocument.Editor.CaretPosition.InsertTextInRun(Date.Now.ToString("HH:mm:ss"))
                Else
                    SelectedDocument.Editor.CaretPosition.InsertTextInRun(Date.Now.ToString("HH:mm"))
                End If
            End If
            UpdateButtons()
        End If
        e.Handled = True
    End Sub

#End Region

#End Region

#Region "-Format"

    Private Sub ClearFormattingButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ClearFormattingButton.Click
        SelectedDocument.Editor.Selection.ClearAllProperties()
        UpdateSelected()
    End Sub

#Region "Font/Font Size/Font Color/Hightlight Color"

    Private UpdatingFont As Boolean = False
    Private Sub FontComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles FontComboBox.SelectionChanged
        If FontComboBox.IsLoaded AndAlso Not UpdatingFont Then
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, FontComboBox.SelectedItem)
            SelectedDocument.Editor.Focus()
            UpdateSelected()
        End If
    End Sub

    Private Sub FontSizeComboBox_KeyDown(sender As Object, e As KeyEventArgs) Handles FontSizeComboBox.KeyDown
        If e.Key = Key.Enter Then
            SelectedDocument.Editor.Focus()
        End If
    End Sub

    Private Sub FontSizeComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles FontSizeComboBox.SelectionChanged
        If FontSizeComboBox.IsLoaded AndAlso Not UpdatingFont Then
            Try
                Dim val As Double = Convert.ToDouble(FontSizeComboBox.SelectedValue)
                SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, val)
                UpdateSelected()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub IncreaseFontButton_Click(sender As Object, e As RoutedEventArgs) Handles IncreaseFontButton.Click
        FontSizeComboBox.SelectedIndex = FontSizeComboBox.SelectedIndex + 1
    End Sub

    Private Sub DecreaseFontButton_Click(sender As Object, e As RoutedEventArgs) Handles DecreaseFontButton.Click
        If Not FontSizeComboBox.SelectedIndex = 0 Then
            FontSizeComboBox.SelectedIndex = FontSizeComboBox.SelectedIndex - 1
        End If
    End Sub

    Private Sub FontColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles FontColorGallery.SelectedColorChanged
        If FontColorGallery.SelectedColor IsNot Nothing Then
            Dim c As New SolidColorBrush
            c.Color = FontColorGallery.SelectedColor
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, c)
            FontColorGallery.SelectedColor = Nothing
        End If
    End Sub

    Private Sub HighlightColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles HighlightColorGallery.SelectedColorChanged
        If HighlightColorGallery.SelectedColor IsNot Nothing Then
            Dim c As New SolidColorBrush
            c.Color = HighlightColorGallery.SelectedColor
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.BackgroundProperty, c)
            HighlightColorGallery.SelectedColor = Nothing
        End If
    End Sub

#End Region

#Region "Blod/Italic/Underline/Strikethrough"

    Private Sub BoldMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles BoldButton.Click
        EditingCommands.ToggleBold.Execute(Nothing, SelectedDocument.Editor)
        UpdateSelected()
    End Sub

    Private Sub ItalicMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ItalicButton.Click
        EditingCommands.ToggleItalic.Execute(Nothing, SelectedDocument.Editor)
        UpdateSelected()
    End Sub

    Private Sub UnderlineMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles UnderlineButton.Click
        EditingCommands.ToggleUnderline.Execute(Nothing, SelectedDocument.Editor)
        UpdateSelected()
    End Sub

    Private Sub StrikethroughButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles StrikethroughButton.Click
        SelectedDocument.Editor.ToggleStrikethrough()
        UpdateSelected()
    End Sub

    Private Sub DoubleStrikethroughButton_PreviewMouseDown(sender As Object, e As RoutedEventArgs) Handles DoubleStrikethroughButton.PreviewMouseDown
        If SelectedDocument.Editor.Selection IsNot Nothing Then
            Dim selectionTextRange As New TextRange(SelectedDocument.Editor.Selection.Start, SelectedDocument.Editor.Selection.[End])
            Dim tdc As TextDecorationCollection = TryCast(selectionTextRange.GetPropertyValue(Inline.TextDecorationsProperty), TextDecorationCollection)
            Dim tdc_clone As New TextDecorationCollection(tdc)
            Dim te1 As TextElement = selectionTextRange.Start.Paragraph
            tdc_clone.Add((New TextDecoration() With {.Location = TextDecorationLocation.Strikethrough, .PenOffset = te1.FontSize * 0.015}))
            tdc_clone.Add((New TextDecoration() With {.Location = TextDecorationLocation.Strikethrough, .PenOffset = te1.FontSize * -0.015}))
            selectionTextRange.ApplyPropertyValue(Inline.TextDecorationsProperty, tdc_clone)
            UpdateSelected()
        End If
    End Sub

    Private Sub DoubleUnderlineButton_PreviewMouseDown(sender As Object, e As RoutedEventArgs) Handles DoubleUnderlineButton.PreviewMouseDown
        Dim selectionTextRange As New TextRange(SelectedDocument.Editor.Selection.Start, SelectedDocument.Editor.Selection.[End])
        Dim tdc As TextDecorationCollection = TryCast(selectionTextRange.GetPropertyValue(Inline.TextDecorationsProperty), TextDecorationCollection)
        Dim tdc_clone As New TextDecorationCollection(tdc)
        Dim te1 As TextElement = selectionTextRange.Start.Paragraph
        tdc_clone.Add((New TextDecoration() With {.Location = TextDecorationLocation.Underline, .PenOffset = te1.FontSize * 0.05}))
        tdc_clone.Add((New TextDecoration() With {.Location = TextDecorationLocation.Underline, .PenOffset = te1.FontSize * -0.05}))
        selectionTextRange.ApplyPropertyValue(Inline.TextDecorationsProperty, tdc_clone)
        UpdateSelected()
    End Sub

#End Region

#Region "Subscript/Superscript"

    Private Sub SubscriptButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles SubscriptButton.Click
        SelectedDocument.Editor.ToggleSubscript()
        UpdateSelected()
    End Sub

    Private Sub SuperscriptButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles SuperscriptButton.Click
        SelectedDocument.Editor.ToggleSuperscript()
        UpdateSelected()
    End Sub

#End Region

#Region "Indent More/Indent Less"

    Private Sub IndentMoreButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles IndentMoreButton.Click
        EditingCommands.IncreaseIndentation.Execute(Nothing, SelectedDocument.Editor)
    End Sub

    Private Sub IndentLessButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles IndentLessButton.Click
        EditingCommands.DecreaseIndentation.Execute(Nothing, SelectedDocument.Editor)
    End Sub

#End Region

#Region "Bullet List/Number List"

    Private Sub SetListStyle(s As TextMarkerStyle)
        Dim run As Run = TryCast(SelectedDocument.Editor.CaretPosition.Parent, Run)
        If run IsNot Nothing Then
            Dim li As ListItem = TryCast(run.Parent, ListItem)
            If li IsNot Nothing Then
                MessageBox.Show(li.Parent.ToString)
            Else
                Dim p As Paragraph = TryCast(run.Parent, Paragraph)
                If p IsNot Nothing Then
                    Dim listitem As ListItem = TryCast(p.Parent, ListItem)
                    If listitem IsNot Nothing Then
                        Dim list2 As List = TryCast(listitem.Parent, List)
                        If list2 IsNot Nothing Then
                            list2.MarkerStyle = s
                        End If
                    Else
                        EditingCommands.ToggleBullets.Execute(Nothing, SelectedDocument.Editor)
                        Dim p2 As Paragraph = TryCast(SelectedDocument.Editor.CaretPosition.Parent, Paragraph)
                        If p2 IsNot Nothing Then
                            Dim listitem2 As ListItem = TryCast(p2.Parent, ListItem)
                            If listitem2 IsNot Nothing Then
                                Dim list2 As List = TryCast(listitem2.Parent, List)
                                If list2 IsNot Nothing Then
                                    list2.MarkerStyle = s
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            Dim p As Paragraph = TryCast(SelectedDocument.Editor.CaretPosition.Parent, Paragraph)
            If p IsNot Nothing Then
                Dim listitem As ListItem = TryCast(p.Parent, ListItem)
                If listitem IsNot Nothing Then
                    Dim list1 As List = TryCast(listitem.Parent, List)
                    If list1 IsNot Nothing Then
                        list1.MarkerStyle = s
                    End If
                Else
                    EditingCommands.ToggleBullets.Execute(Nothing, SelectedDocument.Editor)
                    Dim p2 As Paragraph = TryCast(SelectedDocument.Editor.CaretPosition.Parent, Paragraph)
                    If p2 IsNot Nothing Then
                        Dim listitem2 As ListItem = TryCast(p2.Parent, ListItem)
                        If listitem2 IsNot Nothing Then
                            Dim list2 As List = TryCast(listitem2.Parent, List)
                            If list2 IsNot Nothing Then
                                list2.MarkerStyle = s
                            End If
                        End If
                    End If
                End If
            Else
                Dim tablecell As TableCell = TryCast(SelectedDocument.Editor.CaretPosition.Parent, TableCell)
                If tablecell IsNot Nothing Then
                    Dim par As New Paragraph
                    tablecell.Blocks.Add(par)
                    EditingCommands.ToggleBullets.Execute(Nothing, SelectedDocument.Editor)
                    Dim p2 As Paragraph = TryCast(SelectedDocument.Editor.CaretPosition.Parent, Paragraph)
                    If p2 IsNot Nothing Then
                        Dim listitem2 As ListItem = TryCast(p2.Parent, ListItem)
                        If listitem2 IsNot Nothing Then
                            Dim list2 As List = TryCast(listitem2.Parent, List)
                            If list2 IsNot Nothing Then
                                list2.MarkerStyle = s
                            End If
                        End If
                    End If
                End If
            End If
        End If
        UpdateSelected()
    End Sub

    Private Sub BulletListMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles BulletListButton.Click
        EditingCommands.ToggleBullets.Execute(Nothing, SelectedDocument.Editor)
        UpdateSelected()
    End Sub

    Private Sub DiskBulletButton_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles DiskBulletButton.PreviewMouseDown
        SetListStyle(TextMarkerStyle.Disc)
    End Sub

    Private Sub CircleBulletButton_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles CircleBulletButton.PreviewMouseDown
        SetListStyle(TextMarkerStyle.Circle)
    End Sub

    Private Sub BoxBulletButton_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles BoxBulletButton.PreviewMouseDown
        SetListStyle(TextMarkerStyle.Box)
    End Sub

    Private Sub SquareBulletButton_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles SquareBulletButton.PreviewMouseDown
        SetListStyle(TextMarkerStyle.Square)
    End Sub

    Private Sub NumberListMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles NumberListButton.Click
        EditingCommands.ToggleNumbering.Execute(Nothing, SelectedDocument.Editor)
        UpdateSelected()
    End Sub

    Private Sub DecimalListButton_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles DecimalListButton.PreviewMouseDown
        SetListStyle(TextMarkerStyle.Decimal)
    End Sub

    Private Sub UpperLatinListButton_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles UpperLatinListButton.PreviewMouseDown
        SetListStyle(TextMarkerStyle.UpperLatin)
    End Sub

    Private Sub LowerLatinListButton_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles LowerLatinListButton.PreviewMouseDown
        SetListStyle(TextMarkerStyle.LowerLatin)
    End Sub

    Private Sub UpperRomanListButton_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles UpperRomanListButton.PreviewMouseDown
        SetListStyle(TextMarkerStyle.UpperRoman)
    End Sub

    Private Sub LowerRomanListButton_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles LowerRomanListButton.PreviewMouseDown
        SetListStyle(TextMarkerStyle.LowerRoman)
    End Sub

#End Region

#Region "Align Left/Center/Right/Justify"

    Private Sub AlignLeftMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles AlignLeftButton.Click
        EditingCommands.AlignLeft.Execute(Nothing, SelectedDocument.Editor)
        UpdateSelected()
    End Sub

    Private Sub AlignCenterMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles AlignCenterButton.Click
        EditingCommands.AlignCenter.Execute(Nothing, SelectedDocument.Editor)
        UpdateSelected()
    End Sub

    Private Sub AlignRightMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles AlignRightButton.Click
        EditingCommands.AlignRight.Execute(Nothing, SelectedDocument.Editor)
        UpdateSelected()
    End Sub

    Private Sub AlignJustifyMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles AlignJustifyButton.Click
        EditingCommands.AlignJustify.Execute(Nothing, SelectedDocument.Editor)
        UpdateSelected()
    End Sub

#End Region

#Region "Line Spacing"

    Private Sub LineSpacing1Point0_Click(sender As Object, e As RoutedEventArgs) Handles LineSpacing1Point0.Click
        If SelectedDocument.Editor.CaretPosition.Paragraph IsNot Nothing Then
            SelectedDocument.Editor.CaretPosition.Paragraph.LineHeight = 1.0
            UpdateSelected()
        End If
    End Sub

    Private Sub LineSpacing1Point15_Click(sender As Object, e As RoutedEventArgs) Handles LineSpacing1Point15.Click
        If SelectedDocument.Editor.CaretPosition.Paragraph IsNot Nothing Then
            SelectedDocument.Editor.CaretPosition.Paragraph.LineHeight = 1.15
            UpdateSelected()
        End If
    End Sub

    Private Sub LineSpacing1Point5_Click(sender As Object, e As RoutedEventArgs) Handles LineSpacing1Point5.Click
        If SelectedDocument.Editor.CaretPosition.Paragraph IsNot Nothing Then
            SelectedDocument.Editor.CaretPosition.Paragraph.LineHeight = 1.5
            UpdateSelected()
        End If
    End Sub

    Private Sub LineSpacing2Point0_Click(sender As Object, e As RoutedEventArgs) Handles LineSpacing2Point0.Click
        If SelectedDocument.Editor.CaretPosition.Paragraph IsNot Nothing Then
            SelectedDocument.Editor.CaretPosition.Paragraph.LineHeight = 2.0
            UpdateSelected()
        End If
    End Sub

    Private Sub LineSpacing2Point5_Click(sender As Object, e As RoutedEventArgs) Handles LineSpacing2Point5.Click
        If SelectedDocument.Editor.CaretPosition.Paragraph IsNot Nothing Then
            SelectedDocument.Editor.CaretPosition.Paragraph.LineHeight = 2.5
            UpdateSelected()
        End If
    End Sub

    Private Sub LineSpacing3Point0_Click(sender As Object, e As RoutedEventArgs) Handles LineSpacing3Point0.Click
        If SelectedDocument.Editor.CaretPosition.Paragraph IsNot Nothing Then
            SelectedDocument.Editor.CaretPosition.Paragraph.LineHeight = 3.0
            UpdateSelected()
        End If
    End Sub

    Private Sub CustomLineSpacingMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles CustomLineSpacingMenuItem.Click
        Dim d As New LineSpacingDialog
        If SelectedDocument.Editor.CaretPosition.Paragraph IsNot Nothing Then
            d.number = SelectedDocument.Editor.CaretPosition.Paragraph.LineHeight
        End If
        d.Owner = Me
        d.ShowDialog()
        If d.Res = "OK" Then
            Try
                SelectedDocument.Editor.CaretPosition.Paragraph.LineHeight = d.number
                UpdateSelected()
            Catch ex As Exception
                Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
                m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
                m.Owner = Me
                m.ShowDialog()
            End Try
        End If
    End Sub

#End Region

#Region "ltr/rtl"

    Private Sub LefttoRightButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles LeftToRightButton.Click
        SelectedDocument.Editor.Document.FlowDirection = FlowDirection.LeftToRight
    End Sub

    Private Sub RighttoLeftButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles RightToLeftButton.Click
        SelectedDocument.Editor.Document.FlowDirection = FlowDirection.RightToLeft
    End Sub

#End Region

#End Region

#Region "--Page Layout"

#Region "Margins"

    Private Sub NoneMarginsMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles NoneMarginsMenuItem.Click
        SelectedDocument.Editor.SetPageMargins(New Thickness(0, 0, 0, 0))
        UpdateButtons()
    End Sub

    Private Sub NormalMarginsMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles NormalMarginsMenuItem.Click
        SelectedDocument.Editor.SetPageMargins(New Thickness(96, 96, 96, 96))
        UpdateButtons()
    End Sub

    Private Sub NarrowMarginsMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles NarrowMarginsMenuItem.Click
        SelectedDocument.Editor.SetPageMargins(New Thickness(48, 48, 48, 48))
        UpdateButtons()
    End Sub

    Private Sub LeftMarginBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles LeftMarginBox.ValueChanged
        If IsLoaded Then
            SelectedDocument.Editor.SetPageMargins(New Thickness(LeftMarginBox.Value, SelectedDocument.Editor.docpadding.Top, SelectedDocument.Editor.docpadding.Right, SelectedDocument.Editor.docpadding.Bottom))
            UpdateButtons()
        End If
    End Sub

    Private Sub TopMarginBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles TopMarginBox.ValueChanged
        If IsLoaded Then
            SelectedDocument.Editor.SetPageMargins(New Thickness(SelectedDocument.Editor.docpadding.Left, TopMarginBox.Value, SelectedDocument.Editor.docpadding.Right, SelectedDocument.Editor.docpadding.Bottom))
            UpdateButtons()
        End If
    End Sub

    Private Sub RightMarginBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles RightMarginBox.ValueChanged
        If IsLoaded Then
            SelectedDocument.Editor.SetPageMargins(New Thickness(SelectedDocument.Editor.docpadding.Left, SelectedDocument.Editor.docpadding.Top, RightMarginBox.Value, SelectedDocument.Editor.docpadding.Bottom))
            UpdateButtons()
        End If
    End Sub

    Private Sub BottomMarginBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles BottomMarginBox.ValueChanged
        If IsLoaded Then
            SelectedDocument.Editor.SetPageMargins(New Thickness(SelectedDocument.Editor.docpadding.Left, SelectedDocument.Editor.docpadding.Top, SelectedDocument.Editor.docpadding.Right, BottomMarginBox.Value))
            UpdateButtons()
        End If
    End Sub

#End Region

#Region "Size"

    Private Sub NormalPageSizeMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles NormalPageSizeMenuItem.Click
        SelectedDocument.SetPageSize(1056, 816)
        UpdateButtons()
    End Sub

    Private Sub WidePageSizeMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles WidePageSizeMenuItem.Click
        SelectedDocument.SetPageSize(816, 1056)
        UpdateButtons()
    End Sub

    Private Sub PageHeightBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles PageHeightBox.ValueChanged
        If IsLoaded Then
            SelectedDocument.SetPageSize(PageHeightBox.Value, SelectedDocument.Editor.Document.PageWidth)
            UpdateButtons()
        End If
    End Sub

    Private Sub PageWidthBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles PageWidthBox.ValueChanged
        If IsLoaded Then
            SelectedDocument.SetPageSize(SelectedDocument.Editor.Document.PageHeight, PageWidthBox.Value)
            UpdateButtons()
        End If
    End Sub

#End Region

    Private Sub BackgroundColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles BackgroundColorGallery.SelectedColorChanged
        Dim b As New SolidColorBrush(BackgroundColorGallery.SelectedColor)
        SelectedDocument.Editor.Document.Background = b
    End Sub

    Private Sub BackgroundImageMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles BackgroundImageMenuItem.Click
        Dim open As New Microsoft.Win32.OpenFileDialog()
        open.Filter = "Supported Images(*.bmp;*.gif;*.jpeg;*.jpg;*.png)|*.bmp;*.gif;*.jpeg;*.jpg;*.png|Bitmap Images(*.bmp)|*.bmp|GIF Images(*.gif)|*.gif|JPEG Images(*.jpeg;*.jpg)|*.jpeg;*.jpg|PNG Images(*.png)|*.png|All Files(*.*)|*.*"
        If open.ShowDialog = True Then
            Dim ib As New ImageBrush(New BitmapImage(New Uri(open.FileName)))
            SelectedDocument.Editor.Document.Background = ib
        End If
    End Sub

#End Region

#Region "--NavigationMenuItem"

    Private Sub LineDownButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles LineDownButton.Click
        SelectedDocument.Editor.LineDown()
    End Sub

    Private Sub LineUpButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles LineUpButton.Click
        SelectedDocument.Editor.LineUp()
    End Sub

    Private Sub LineLeftButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles LineLeftButton.Click
        SelectedDocument.Editor.LineLeft()
    End Sub

    Private Sub LineRightButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles LineRightButton.Click
        SelectedDocument.Editor.LineRight()
    End Sub

    Private Sub PageDownButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles PageDownButton.Click
        SelectedDocument.Editor.PageDown()
    End Sub

    Private Sub PageUpButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles PageUpButton.Click
        SelectedDocument.Editor.PageUp()
    End Sub

    Private Sub PageLeftButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles PageLeftButton.Click
        SelectedDocument.Editor.PageLeft()
    End Sub

    Private Sub PageRightButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles PageRightButton.Click
        SelectedDocument.Editor.PageRight()
    End Sub

    Private Sub StartButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles StartButton.Click
        EditingCommands.MoveToDocumentStart.Execute(Nothing, SelectedDocument.Editor)
    End Sub

    Private Sub EndButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles EndButton.Click
        EditingCommands.MoveToDocumentEnd.Execute(Nothing, SelectedDocument.Editor)
    End Sub

#End Region

#Region "--ToolsMenuItem"

    Private Sub SpellCheckButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles SpellCheckButton.Click
        Dim d As New SpellCheckDialog, cp As TextPointer = SelectedDocument.Editor.Selection.Start.GetPositionAtOffset(1)
        d.Owner = Me
        Dim sp As SpellingError = SelectedDocument.Editor.GetSpellingError(cp)
        If sp IsNot Nothing Then
            For Each i As String In sp.Suggestions
                d.WordListBox.Items.Add(i)
            Next
            d.WordListBox.SelectedIndex = 0
            d.WordListBox.Focus()
        End If
        d.ShowDialog()
        If d.Res = "OK" Then
            sp.Correct(d.WordListBox.SelectedItem)
        End If
    End Sub

    Private Sub PreviousSpellingErrorMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles PreviousSpellingErrorMenuItem.Click
        Try
            Dim tp As TextPointer = SelectedDocument.Editor.GetNextSpellingErrorPosition(SelectedDocument.Editor.Selection.Start, LogicalDirection.Backward)
            Dim tr As TextRange = SelectedDocument.Editor.GetSpellingErrorRange(tp)
            SelectedDocument.Editor.Selection.Select(tr.Start, tr.End)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub NextSpellingErrorMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles NextSpellingErrorMenuItem.Click
        Try
            Dim tp As TextPointer = SelectedDocument.Editor.GetNextSpellingErrorPosition(SelectedDocument.Editor.Selection.End, LogicalDirection.Forward)
            Dim tr As TextRange = SelectedDocument.Editor.GetSpellingErrorRange(tp)
            SelectedDocument.Editor.Selection.Select(tr.Start, tr.End)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub IgnoreSpellingErrorMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles IgnoreSpellingErrorMenuItem.Click
        Dim cp As TextPointer = SelectedDocument.Editor.Selection.Start.GetPositionAtOffset(1)
        Dim sp As SpellingError = SelectedDocument.Editor.GetSpellingError(cp)
        If sp IsNot Nothing Then
            Dim r As TextRange = SelectedDocument.Editor.GetSpellingErrorRange(cp)
            If Not My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath + "\spellcheck_ignorelist.lex") Then
                Dim sr As IO.StreamWriter = IO.File.CreateText(My.Application.Info.DirectoryPath + "\spellcheck_ignorelist.lex")
                sr.Close()
            End If
            Dim fileIn As New IO.StreamReader(My.Application.Info.DirectoryPath + "\spellcheck_ignorelist.lex")
            Dim strData As String = ""
            Dim lngCount As Long = 1
            Dim canadd As Boolean = True
            While (Not (fileIn.EndOfStream))
                strData = fileIn.ReadLine()
                If strData Is r.Text Then
                    canadd = False
                End If
                lngCount = lngCount + 1
            End While
            fileIn.Close()
            If canadd Then
                My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath + "\spellcheck_ignorelist.lex", vbNewLine, True)
                My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath + "\spellcheck_ignorelist.lex", r.Text, True)
            End If
            sp.IgnoreAll()
        End If
    End Sub

    Private Sub CorrectAllMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles CorrectAllMenuItem.Click
        While True
            Dim sp As TextPointer = SelectedDocument.Editor.GetNextSpellingErrorPosition(SelectedDocument.Editor.Document.ContentStart.DocumentStart, LogicalDirection.Forward)
            If sp Is Nothing Then
                Exit While
            End If
            Dim se As SpellingError = SelectedDocument.Editor.GetSpellingError(SelectedDocument.Editor.GetNextSpellingErrorPosition(SelectedDocument.Editor.Document.ContentStart.DocumentStart, LogicalDirection.Forward))
            If se IsNot Nothing Then
                For Each i As String In se.Suggestions
                    se.Correct(i)
                    Exit For
                Next
            Else
                Exit While
            End If
        End While
    End Sub

    Private Sub TextToSpeechButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles TextToSpeechButton.Click
        Speech.Rate = My.Settings.Options_TTSS
        If Speech.State = System.Speech.Synthesis.SynthesizerState.Speaking Then
            Speech.SpeakAsyncCancelAll()
        Else
            Try
                Speech.SelectVoice(Speech.GetInstalledVoices.Item(My.Settings.Options_TTSV).VoiceInfo.Name)
                Speech.SpeakAsync(SelectedDocument.Editor.Selection.Text.ToString)
            Catch ex As Exception
                Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
                m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
                m.Owner = Me
                m.ShowDialog()
            End Try
        End If
    End Sub

    Private Sub TranslateButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles TranslateButton.Click
        Dim trans As New TranslateDialog(SelectedDocument.Editor.Selection.Text.ToString)
        trans.Owner = Me
        trans.ShowDialog()
        If trans.res = True Then
            SelectedDocument.Editor.Selection.Text = TryCast(trans.TranslatedText.Content, String)
        End If
    End Sub

    Private Sub DefinitionsButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles DefinitionsButton.Click
        If SelectedDocument.Editor.Selection.Text.Length > 0 Then
            Process.Start("http://www.bing.com/Dictionary/search?q=define+" + SelectedDocument.Editor.Selection.Text)
        Else
            Dim m As New MessageBoxDialog("Select a word first.", "Warning!", Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/info32.png"))
            m.Owner = Me
            m.ShowDialog()
        End If
    End Sub

    Private Sub OptionsMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles OptionsButton.Click
        Dim o As New OptionsDialog
        o.Owner = Me
        If o.ShowDialog = True Then
            For Each i As DocumentTab In TabCell.Children
                If My.Settings.Options_SpellCheck = True Then
                    SpellCheck.SetIsEnabled(i.Editor, True)
                Else
                    SpellCheck.SetIsEnabled(i.Editor, False)
                End If
            Next
            If o.PluginsCheckBox.IsChecked Then
                Dim plugins As New Plugins
                PluginsGroup.Items.Clear()
                For Each file As String In My.Computer.FileSystem.GetFiles(My.Application.Info.DirectoryPath + "\Plugins\")
                    Try
                        Dim plugin As New IO.FileInfo(file), pluginname As String = plugin.Name.Remove(plugin.Name.Length - 3)
                        Dim IsEventPlugin As Boolean = plugins.IsEventPlugin(pluginname, My.Computer.FileSystem.ReadAllText(file))
                        If IsEventPlugin Then
                            Dim ribbonbutton As New Fluent.Button
                            ribbonbutton.Header = pluginname
                            If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath + "\Plugins\" + pluginname + "32.png") Then
                                Dim bitmap As New BitmapImage(New Uri(My.Application.Info.DirectoryPath + "\Plugins\" + pluginname + "32.png"))
                                ribbonbutton.LargeIcon = bitmap
                            Else
                                Dim bitmap As New BitmapImage(New Uri(My.Application.Info.DirectoryPath + "\Images\Tools\plugins32.png"))
                                ribbonbutton.LargeIcon = bitmap
                            End If
                            Dim tip As New Fluent.ScreenTip
                            tip.Title = pluginname
                            tip.Image = New BitmapImage(New Uri("pack://application:,,,/Images/Tools/plugins48.png"))
                            ribbonbutton.Tag = file
                            AddHandler ribbonbutton.Click, New RoutedEventHandler(AddressOf RunPlugin)
                            PluginsGroup.Items.Add(ribbonbutton)
                        End If
                    Catch ex As Exception
                        Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
                        m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
                        m.Owner = Me
                        m.ShowDialog()
                    End Try
                Next
            Else
                PluginsGroup.Items.Clear()
                PluginsGroup.Visibility = Windows.Visibility.Collapsed
            End If
        End If
    End Sub

#End Region

#Region "Contextual Groups"

#Region "Update"

#Region "Table Cell"

    Private Sub SetSelectedTableCell(tablecell As TableCell)
        SelectedDocument.Editor.SelectedTableCell = tablecell
        Dim tr As TableRow = SelectedDocument.Editor.SelectedTableCell.Parent, trg As TableRowGroup = tr.Parent, table As Table = trg.Parent
        TableBorderSizeBox.Value = table.BorderThickness.Left
        TableCellSpacingBox.Value = table.CellSpacing
        TableCellBorderSizeBox.Value = tablecell.BorderThickness.Left
        EditTableCellGroup.Visibility = Windows.Visibility.Visible
    End Sub

#End Region

#Region "Image"

    Private Sub SetSelectedImage(Img As Image)
        SelectedDocument.Editor.SelectedImage = New Image
        ImageHeightBox.Value = Img.Height
        ImageWidthBox.Value = Img.Width
        If Img.Stretch = Stretch.Fill Then
            ImageResizeModeComboBox.SelectedIndex = 0
        ElseIf Img.Stretch = Stretch.Uniform Then
            ImageResizeModeComboBox.SelectedIndex = 1
        ElseIf Img.Stretch = Stretch.UniformToFill Then
            ImageResizeModeComboBox.SelectedIndex = 2
        ElseIf Img.Stretch = Stretch.None Then
            ImageResizeModeComboBox.SelectedIndex = 3
        End If
        SelectedDocument.Editor.SelectedImage = Img
        EditTableCellGroup.Visibility = Windows.Visibility.Collapsed
        EditImageGroup.Visibility = Windows.Visibility.Visible
        EditVideoGroup.Visibility = Windows.Visibility.Collapsed
        EditShapeGroup.Visibility = Windows.Visibility.Collapsed
        EditObjectGroup.Visibility = Windows.Visibility.Collapsed
    End Sub

#End Region

#Region "Video"

    Private Sub SetSelectedVideo(vid As MediaElement)
        SelectedDocument.Editor.SelectedVideo = vid
        VideoHeightBox.Value = vid.Height
        VideoWidthBox.Value = vid.Width
        If vid.Stretch = Stretch.Fill Then
            VideoResizeModeComboBox.SelectedIndex = 0
        ElseIf vid.Stretch = Stretch.Uniform Then
            VideoResizeModeComboBox.SelectedIndex = 1
        ElseIf vid.Stretch = Stretch.UniformToFill Then
            VideoResizeModeComboBox.SelectedIndex = 2
        ElseIf vid.Stretch = Stretch.None Then
            VideoResizeModeComboBox.SelectedIndex = 3
        End If
        'TODO:
        EditTableCellGroup.Visibility = Windows.Visibility.Collapsed
        EditImageGroup.Visibility = Windows.Visibility.Collapsed
        EditVideoGroup.Visibility = Windows.Visibility.Visible
        EditShapeGroup.Visibility = Windows.Visibility.Collapsed
        EditObjectGroup.Visibility = Windows.Visibility.Collapsed
    End Sub

#End Region

#Region "Shape"

    Private Sub SetSelectedShape(s As Shape)
        SelectedDocument.Editor.SelectedShape = s
        ShapeHeightBox.Value = s.Height
        ShapeWidthBox.Value = s.Width
        'TODO: Add support for Video Resize Mode
        'If vid.Stretch = Stretch.Fill Then
        '    VideoResizeModeComboBox.SelectedIndex = 0
        'ElseIf vid.Stretch = Stretch.Uniform Then
        '    VideoResizeModeComboBox.SelectedIndex = 1
        'ElseIf vid.Stretch = Stretch.UniformToFill Then
        '    VideoResizeModeComboBox.SelectedIndex = 2
        'ElseIf vid.Stretch = Stretch.None Then
        '    VideoResizeModeComboBox.SelectedIndex = 3
        'End If
        EditTableCellGroup.Visibility = Windows.Visibility.Collapsed
        EditImageGroup.Visibility = Windows.Visibility.Collapsed
        EditVideoGroup.Visibility = Windows.Visibility.Collapsed
        EditShapeGroup.Visibility = Windows.Visibility.Visible
        EditObjectGroup.Visibility = Windows.Visibility.Collapsed
    End Sub

#End Region

#Region "Object"

    Private Sub SetSelectedObject(uielement As UIElement)
        If uielement.GetType Is GetType(Button) OrElse uielement.GetType Is GetType(RadioButton) OrElse uielement.GetType Is GetType(CheckBox) OrElse uielement.GetType Is GetType(TextBlock) Then
            SelectedDocument.Editor.SelectedObject = uielement
            ObjectBorderGroup.Visibility = Windows.Visibility.Visible
            Dim button As Button = TryCast(uielement, Button)
            If button IsNot Nothing Then
                ObjectHeightBox.Value = button.Height
                ObjectWidthBox.Value = button.Width
                ObjectBorderSizeBox.Value = button.BorderThickness.Left
                ObjectTextBox.Text = button.Content.ToString
                If button.IsEnabled Then
                    ObjectEnabledCheckBox.IsChecked = True
                Else
                    ObjectEnabledCheckBox.IsChecked = False
                End If
            Else
                Dim radiobutton As RadioButton = TryCast(uielement, RadioButton)
                If radiobutton IsNot Nothing Then
                    ObjectHeightBox.Value = radiobutton.Height
                    ObjectWidthBox.Value = radiobutton.Width
                    ObjectBorderSizeBox.Value = radiobutton.BorderThickness.Left
                    ObjectTextBox.Text = radiobutton.Content.ToString
                    If radiobutton.IsEnabled Then
                        ObjectEnabledCheckBox.IsChecked = True
                    Else
                        ObjectEnabledCheckBox.IsChecked = False
                    End If
                Else
                    Dim checkbox As CheckBox = TryCast(uielement, CheckBox)
                    If checkbox IsNot Nothing Then
                        ObjectHeightBox.Value = checkbox.Height
                        ObjectWidthBox.Value = checkbox.Width
                        ObjectBorderSizeBox.Value = checkbox.BorderThickness.Left
                        ObjectTextBox.Text = checkbox.Content.ToString
                        If checkbox.IsEnabled Then
                            ObjectEnabledCheckBox.IsChecked = True
                        Else
                            ObjectEnabledCheckBox.IsChecked = False
                        End If
                    Else
                        Dim textblock As TextBlock = TryCast(uielement, TextBlock)
                        If textblock IsNot Nothing Then
                            ObjectHeightBox.Value = textblock.Height
                            ObjectWidthBox.Value = textblock.Width
                            ObjectBorderGroup.Visibility = Windows.Visibility.Collapsed
                            ObjectTextBox.Text = textblock.Text
                            If textblock.IsEnabled Then
                                ObjectEnabledCheckBox.IsChecked = True
                            Else
                                ObjectEnabledCheckBox.IsChecked = False
                            End If
                        End If
                    End If
                End If
            End If
            EditTableCellGroup.Visibility = Windows.Visibility.Collapsed
            EditImageGroup.Visibility = Windows.Visibility.Collapsed
            EditVideoGroup.Visibility = Windows.Visibility.Collapsed
            EditObjectGroup.Visibility = Windows.Visibility.Visible
        End If
    End Sub

#End Region

    Private Sub UpdateContextualTabs()
        If Not SelectedDocument.Editor.Selection.IsEmpty Then
            For Each block As Block In SelectedDocument.Editor.Document.Blocks
                Dim blockui As BlockUIContainer = TryCast(block, BlockUIContainer)
                If blockui IsNot Nothing Then
                    Dim uielement As UIElement = TryCast(blockui.Child, UIElement)
                    If uielement IsNot Nothing Then
                        If uielement.GetType Is GetType(Image) Then
                            SetSelectedImage(uielement)
                            Exit Sub
                        Else
                            SetSelectedObject(uielement)
                            Exit Sub
                        End If
                    End If
                Else
                    Dim list As List = TryCast(block, List)
                    If list IsNot Nothing Then
                        For Each litem As ListItem In list.ListItems
                            For Each b As Block In litem.Blocks
                                Dim bui As BlockUIContainer = TryCast(block, BlockUIContainer)
                                If bui IsNot Nothing Then
                                    'TODO: (20xx.xx) add support for blockui inside of listitems
                                Else
                                    Dim p As Paragraph = TryCast(b, Paragraph)
                                    If p IsNot Nothing Then
                                        For Each i As Inline In p.Inlines
                                            Dim iui As InlineUIContainer = TryCast(i, InlineUIContainer)
                                            If iui IsNot Nothing Then
                                                If SelectedDocument.Editor.Selection.Start.CompareTo(iui.ElementStart) = 0 AndAlso SelectedDocument.Editor.Selection.End.CompareTo(iui.ElementEnd) = 0 Then
                                                    If iui.Child.GetType Is GetType(Image) Then
                                                        SetSelectedImage(iui.Child)
                                                        Exit Sub
                                                    ElseIf iui.Child.GetType Is GetType(MediaElement) Then
                                                        SetSelectedVideo(iui.Child)
                                                        Exit Sub
                                                    ElseIf iui.Child.GetType Is GetType(Shape) Then
                                                        SetSelectedShape(iui.Child)
                                                        Exit Sub
                                                    Else
                                                        SetSelectedObject(iui.Child)
                                                        Exit Sub
                                                    End If
                                                End If
                                            End If
                                        Next
                                    Else
                                        'TODO: (20xx.xx) add suport for the rest of the "Container" types
                                    End If
                                End If
                            Next
                        Next
                    Else
                        Dim par As Paragraph = TryCast(block, Paragraph)
                        If par IsNot Nothing Then
                            For Each inline As Inline In par.Inlines
                                Dim inlineui As InlineUIContainer = TryCast(inline, InlineUIContainer)
                                If inlineui IsNot Nothing Then
                                    If SelectedDocument.Editor.Selection.Start.CompareTo(inlineui.ElementStart) = 0 AndAlso SelectedDocument.Editor.Selection.End.CompareTo(inlineui.ElementEnd) = 0 Then
                                        If inlineui.Child.GetType Is GetType(Image) Then
                                            SetSelectedImage(inlineui.Child)
                                            Exit Sub
                                        ElseIf inlineui.Child.GetType Is GetType(MediaElement) Then
                                            SetSelectedVideo(inlineui.Child)
                                            Exit Sub
                                        ElseIf inlineui.Child.GetType Is GetType(Shape) Then
                                            SetSelectedShape(inlineui.Child)
                                            Exit Sub
                                        Else
                                            SetSelectedObject(inlineui.Child)
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            Dim sec As Section = TryCast(block, Section)
                            If sec IsNot Nothing Then
                                'TODO: (20xx.xx) add support for sections
                            Else
                                Dim table As Table = TryCast(block, Table)
                                If table IsNot Nothing Then
                                    For Each rowgroup As TableRowGroup In table.RowGroups
                                        For Each row As TableRow In rowgroup.Rows
                                            For Each cell As TableCell In row.Cells
                                                For Each b As Block In cell.Blocks
                                                    Dim bui As BlockUIContainer = TryCast(b, BlockUIContainer)
                                                    If bui IsNot Nothing Then
                                                        Dim uie As UIElement = TryCast(bui.Child, UIElement)
                                                        If uie IsNot Nothing Then
                                                            If uie.GetType Is GetType(Image) Then
                                                                SetSelectedImage(uie)
                                                                Exit Sub
                                                            ElseIf uie.GetType Is GetType(MediaElement) Then
                                                                SetSelectedVideo(uie)
                                                                Exit Sub
                                                            ElseIf uie.GetType Is GetType(Shape) Then
                                                                SetSelectedShape(uie)
                                                                Exit Sub
                                                            Else
                                                                SetSelectedObject(uie)
                                                                Exit Sub
                                                            End If
                                                        End If
                                                    Else
                                                        Dim l As List = TryCast(b, List)
                                                        If l IsNot Nothing Then
                                                            'For Each litem As ListItem In l.ListItems
                                                            '    For Each b2 As Block In litem.Blocks
                                                            '        Dim bui2 As BlockUIContainer = TryCast(b2, BlockUIContainer)
                                                            '        If bui2 IsNot Nothing Then
                                                            '            'TODO: add support for blockui inside of listitems inside of tables
                                                            '        Else
                                                            '            Dim p As Paragraph = TryCast(b2, Paragraph)
                                                            '            If p IsNot Nothing Then
                                                            '                For Each i As Inline In p.Inlines
                                                            '                    Dim iui As InlineUIContainer = TryCast(i, InlineUIContainer)
                                                            '                    If iui IsNot Nothing Then
                                                            '                        If SelectedDocument.Editor.Selection.Start.CompareTo(iui.ElementStart) = 0 AndAlso SelectedDocument.Editor.Selection.End.CompareTo(iui.ElementEnd) = 0 Then
                                                            '                            If iui.Child.GetType Is GetType(Image) Then
                                                            '                                SetSelectedImage(iui.Child)
                                                            '                                Exit Sub
                                                            '                            Else
                                                            '                                SetSelectedObject(iui.Child)
                                                            '                                Exit Sub
                                                            '                            End If
                                                            '                        End If
                                                            '                    End If
                                                            '                Next
                                                            '            Else
                                                            '                'TODO: add suport for the rest of the "Container" types
                                                            '            End If
                                                            '        End If
                                                            '    Next
                                                            'Next
                                                        Else
                                                            Dim p As Paragraph = TryCast(b, Paragraph)
                                                            If p IsNot Nothing Then
                                                                For Each inl As Inline In p.Inlines
                                                                    Dim inlui As InlineUIContainer = TryCast(inl, InlineUIContainer)
                                                                    If inlui IsNot Nothing Then
                                                                        If SelectedDocument.Editor.Selection.Start.CompareTo(inlui.ElementStart) = 0 AndAlso SelectedDocument.Editor.Selection.End.CompareTo(inlui.ElementEnd) = 0 Then
                                                                            If inlui.Child.GetType Is GetType(Image) Then
                                                                                SetSelectedImage(inlui.Child)
                                                                                Exit Sub
                                                                            ElseIf inlui.Child.GetType Is GetType(MediaElement) Then
                                                                                SetSelectedVideo(inlui.Child)
                                                                                Exit Sub
                                                                            ElseIf inlui.Child.GetType Is GetType(Shape) Then
                                                                                SetSelectedShape(inlui.Child)
                                                                                Exit Sub
                                                                            Else
                                                                                SetSelectedObject(inlui.Child)
                                                                                Exit Sub
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next
                                                            Else
                                                                Dim s As Section = TryCast(b, Section)
                                                                If s IsNot Nothing Then
                                                                    'TODO:(20xx.xx) add support for sections inside of tables
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Next
                                            Next
                                        Next
                                    Next
                                Else
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        Else
            EditTableCellGroup.Visibility = Visibility.Collapsed
            EditImageGroup.Visibility = Visibility.Collapsed
            EditVideoGroup.Visibility = Visibility.Collapsed
            EditShapeGroup.Visibility = Visibility.Collapsed
            EditObjectGroup.Visibility = Visibility.Collapsed
            If SelectedDocument.Editor.CaretPosition.Paragraph IsNot Nothing Then
                Dim p As Paragraph = SelectedDocument.Editor.CaretPosition.Paragraph
                If p.Parent.GetType Is GetType(TableCell) Then
                    SetSelectedTableCell(p.Parent)
                End If
            End If
            If SelectedDocument.Editor.CaretPosition.Parent.GetType Is GetType(TableCell) Then
                SetSelectedTableCell(SelectedDocument.Editor.CaretPosition.Parent)
            End If
            '(Debug info)
            'Dim m As New MessageBoxDialog(SelectedDocument.Editor.CaretPosition.GetType.ToString, "Info!", Nothing, Nothing)
            'm.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/info32.png"))
            'm.Owner = Me
            'm.ShowDialog()
            '(End of Debug info)
        End If
    End Sub

#End Region

#Region "EditImageTab"

#Region "Resize"

    Private Sub ImageHeightBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles ImageHeightBox.ValueChanged
        If IsLoaded Then
            SelectedDocument.Editor.SelectedImage.Height = ImageHeightBox.Value
        End If
    End Sub

    Private Sub ImageWidthBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles ImageWidthBox.ValueChanged
        If IsLoaded Then
            SelectedDocument.Editor.SelectedImage.Width = ImageWidthBox.Value
        End If
    End Sub

    Private Sub ImageResizeModeComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ImageResizeModeComboBox.SelectionChanged
        If ImageResizeModeComboBox.SelectedIndex = 0 Then
            SelectedDocument.Editor.SelectedImage.Stretch = Stretch.Fill
        ElseIf ImageResizeModeComboBox.SelectedIndex = 1 Then
            SelectedDocument.Editor.SelectedImage.Stretch = Stretch.Uniform
        ElseIf ImageResizeModeComboBox.SelectedIndex = 2 Then
            SelectedDocument.Editor.SelectedImage.Stretch = Stretch.UniformToFill
        ElseIf ImageResizeModeComboBox.SelectedIndex = 3 Then
            SelectedDocument.Editor.SelectedImage.Stretch = Stretch.None
        End If
    End Sub

#End Region

#Region "Rotate"

    Private Sub RotateImageLeftMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles RotateImageLeftMenuItem.Click
        Dim imgprops As Thickness = SelectedDocument.Editor.SelectedImage.Tag
        imgprops.Left -= 90
        SelectedDocument.Editor.SelectedImage.Tag = imgprops
        Dim transform As TransformGroup = SelectedDocument.Editor.SelectedImage.LayoutTransform
        Dim rotate As New RotateTransform(imgprops.Left)
        transform.Children.Item(0) = rotate
        SelectedDocument.Editor.SelectedImage.LayoutTransform = transform
    End Sub

    Private Sub RotateImageRightMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles RotateImageRightMenuItem.Click
        Dim imgprops As Thickness = SelectedDocument.Editor.SelectedImage.Tag
        imgprops.Left += 90
        SelectedDocument.Editor.SelectedImage.Tag = imgprops
        Dim transform As TransformGroup = SelectedDocument.Editor.SelectedImage.LayoutTransform
        Dim rotate As New RotateTransform(imgprops.Left)
        transform.Children.Item(0) = rotate
        SelectedDocument.Editor.SelectedImage.LayoutTransform = transform
    End Sub

#End Region

#Region "Flip"

    Private Sub FlipImageHorizontalMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles FlipImageHorizontalMenuItem.Click
        'TODO: (20xx.xx) add flip image horizontal support
        Dim imgprops As Thickness = SelectedDocument.Editor.SelectedImage.Tag
        If imgprops.Top = 1 Then
            imgprops.Top = -1
        Else
            imgprops.Top = 1
        End If
        SelectedDocument.Editor.SelectedImage.Tag = imgprops
        Dim transform As TransformGroup = SelectedDocument.Editor.SelectedImage.LayoutTransform
        Dim hflip As New ScaleTransform(imgprops.Top, imgprops.Right)
        transform.Children.Item(1) = hflip
        SelectedDocument.Editor.SelectedImage.LayoutTransform = transform
    End Sub

    Private Sub FlipImageVerticalMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles FlipImageVerticalMenuItem.Click
        'TODO: (20xx.xx) add flip image vertical support
        Dim imgprops As Thickness = SelectedDocument.Editor.SelectedImage.Tag
        If imgprops.Right = 1 Then
            imgprops.Right = -1
        Else
            imgprops.Right = 1
        End If
        SelectedDocument.Editor.SelectedImage.Tag = imgprops
        Dim transform As TransformGroup = SelectedDocument.Editor.SelectedImage.LayoutTransform
        Dim vflip As New ScaleTransform(imgprops.Top, imgprops.Right)
        transform.Children.Item(2) = vflip
        SelectedDocument.Editor.SelectedImage.LayoutTransform = vflip
    End Sub

#End Region

#End Region

#Region "EditVideoTab"

#Region "Resize"

    Private Sub VideoHeightBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles VideoHeightBox.ValueChanged
        If IsLoaded Then
            SelectedDocument.Editor.SelectedVideo.Height = VideoHeightBox.Value
        End If
    End Sub

    Private Sub VideoWidthBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles VideoWidthBox.ValueChanged
        If IsLoaded Then
            SelectedDocument.Editor.SelectedVideo.Width = VideoWidthBox.Value
        End If
    End Sub

    Private Sub VideoResizeModeComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles VideoResizeModeComboBox.SelectionChanged
        If VideoResizeModeComboBox.SelectedIndex = 0 Then
            SelectedDocument.Editor.SelectedVideo.Stretch = Stretch.Fill
        ElseIf VideoResizeModeComboBox.SelectedIndex = 1 Then
            SelectedDocument.Editor.SelectedVideo.Stretch = Stretch.Uniform
        ElseIf VideoResizeModeComboBox.SelectedIndex = 2 Then
            SelectedDocument.Editor.SelectedVideo.Stretch = Stretch.UniformToFill
        ElseIf VideoResizeModeComboBox.SelectedIndex = 3 Then
            SelectedDocument.Editor.SelectedVideo.Stretch = Stretch.None
        End If
    End Sub

#End Region

    Private Sub VideoPlayMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles VideoPlayMenuItem.Click
        SelectedDocument.Editor.SelectedVideo.Play()
    End Sub

#End Region

#Region "EditTableCellTab"

    Private LastSelectedTableEditingTab As Integer = 0
    Private Sub MainBar_SelectedTabChanged(sender As Object, e As SelectionChangedEventArgs) Handles MainBar.SelectedTabChanged
        If MainBar.SelectedTabItem Is EditTableTab Then
            LastSelectedTableEditingTab = 0
        ElseIf MainBar.SelectedTabItem Is EditTableCellTab Then
            LastSelectedTableEditingTab = 1
        End If
    End Sub

#Region "EditTableTab"

    Private Sub TableBackgroundColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles TableBackgroundColorGallery.SelectedColorChanged
        Dim tr As TableRow = SelectedDocument.Editor.SelectedTableCell.Parent, trg As TableRowGroup = tr.Parent, table As Table = trg.Parent
        table.Background = New SolidColorBrush(TableBackgroundColorGallery.SelectedColor)
    End Sub

    Private Sub TableBackgroundImageMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles TableBackgroundImageMenuItem.Click
        Dim open As New Microsoft.Win32.OpenFileDialog()
        open.Filter = "Supported Images(*.bmp;*.gif;*.jpeg;*.jpg;*.png)|*.bmp;*.gif;*.jpeg;*.jpg;*.png|Bitmap Images(*.bmp)|*.bmp|GIF Images(*.gif)|*.gif|JPEG Images(*.jpeg;*.jpg)|*.jpeg;*.jpg|PNG Images(*.png)|*.png|All Files(*.*)|*.*"
        If open.ShowDialog = True Then
            Dim ib As New ImageBrush(New BitmapImage(New Uri(open.FileName)))
            Dim tr As TableRow = SelectedDocument.Editor.SelectedTableCell.Parent, trg As TableRowGroup = tr.Parent, table As Table = trg.Parent
            table.Background = ib
        End If
    End Sub

    Private Sub TableBorderColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles TableBorderColorGallery.SelectedColorChanged
        Dim tr As TableRow = SelectedDocument.Editor.SelectedTableCell.Parent, trg As TableRowGroup = tr.Parent, table As Table = trg.Parent
        table.BorderBrush = New SolidColorBrush(TableBorderColorGallery.SelectedColor)
    End Sub

    Private Sub TableBorderBackgroundImageMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles TableBorderBackgroundImageMenuItem.Click
        Dim open As New Microsoft.Win32.OpenFileDialog()
        open.Filter = "Supported Images(*.bmp;*.gif;*.jpeg;*.jpg;*.png)|*.bmp;*.gif;*.jpeg;*.jpg;*.png|Bitmap Images(*.bmp)|*.bmp|GIF Images(*.gif)|*.gif|JPEG Images(*.jpeg;*.jpg)|*.jpeg;*.jpg|PNG Images(*.png)|*.png|All Files(*.*)|*.*"
        If open.ShowDialog = True Then
            Dim ib As New ImageBrush(New BitmapImage(New Uri(open.FileName)))
            Dim tr As TableRow = SelectedDocument.Editor.SelectedTableCell.Parent, trg As TableRowGroup = tr.Parent, table As Table = trg.Parent
            table.BorderBrush = ib
        End If
    End Sub

    Private Sub TableBorderSizeBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles TableBorderSizeBox.ValueChanged
        If IsLoaded Then
            Dim tr As TableRow = SelectedDocument.Editor.SelectedTableCell.Parent, trg As TableRowGroup = tr.Parent, table As Table = trg.Parent
            Dim size As Integer = Convert.ToInt32(TableBorderSizeBox.Value)
            table.BorderThickness = New Thickness(size, size, size, size)
        End If
    End Sub

    Private Sub TableCellSpacingBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles TableCellSpacingBox.ValueChanged
        If IsLoaded Then
            Dim tr As TableRow = SelectedDocument.Editor.SelectedTableCell.Parent, trg As TableRowGroup = tr.Parent, table As Table = trg.Parent
            table.CellSpacing = TableCellSpacingBox.Value
        End If
    End Sub

#End Region

    Private Sub TableCellBackgroundColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles TableCellBackgroundColorGallery.SelectedColorChanged
        SelectedDocument.Editor.SelectedTableCell.Background = New SolidColorBrush(TableCellBackgroundColorGallery.SelectedColor)
    End Sub

    Private Sub TableCellBackgroundImageMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles TableCellBackgroundImageMenuItem.Click
        Dim open As New Microsoft.Win32.OpenFileDialog()
        open.Filter = "Supported Images(*.bmp;*.gif;*.jpeg;*.jpg;*.png)|*.bmp;*.gif;*.jpeg;*.jpg;*.png|Bitmap Images(*.bmp)|*.bmp|GIF Images(*.gif)|*.gif|JPEG Images(*.jpeg;*.jpg)|*.jpeg;*.jpg|PNG Images(*.png)|*.png|All Files(*.*)|*.*"
        If open.ShowDialog = True Then
            Dim ib As New ImageBrush(New BitmapImage(New Uri(open.FileName)))
            SelectedDocument.Editor.SelectedTableCell.Background = ib
        End If
    End Sub

    Private Sub TableCellBorderColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles TableCellBorderColorGallery.SelectedColorChanged
        SelectedDocument.Editor.SelectedTableCell.BorderBrush = New SolidColorBrush(TableCellBorderColorGallery.SelectedColor)
    End Sub

    Private Sub TableCellBorderBackgroundImageMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles TableCellBorderBackgroundImageMenuItem.Click
        Dim open As New Microsoft.Win32.OpenFileDialog()
        open.Filter = "Supported Images(*.bmp;*.gif;*.jpeg;*.jpg;*.png)|*.bmp;*.gif;*.jpeg;*.jpg;*.png|Bitmap Images(*.bmp)|*.bmp|GIF Images(*.gif)|*.gif|JPEG Images(*.jpeg;*.jpg)|*.jpeg;*.jpg|PNG Images(*.png)|*.png|All Files(*.*)|*.*"
        If open.ShowDialog = True Then
            Dim ib As New ImageBrush(New BitmapImage(New Uri(open.FileName)))
            SelectedDocument.Editor.SelectedTableCell.BorderBrush = ib
        End If
    End Sub

    Private Sub TableCellBorderSizeBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles TableCellBorderSizeBox.ValueChanged
        If IsLoaded Then
            Dim size As Double = TableCellBorderSizeBox.Value
            SelectedDocument.Editor.SelectedTableCell.BorderThickness = New Thickness(size, size, size, size)
        End If
    End Sub

#End Region

#Region "EditObjectTab"

    Private Sub ObjectHeightBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles ObjectHeightBox.ValueChanged
        If IsLoaded Then
            Dim button As Button = TryCast(SelectedDocument.Editor.SelectedObject, Button)
            If button IsNot Nothing Then
                button.Height = ObjectHeightBox.Value
            Else
                Dim radiobutton As RadioButton = TryCast(SelectedDocument.Editor.SelectedObject, RadioButton)
                If radiobutton IsNot Nothing Then
                    radiobutton.Height = ObjectHeightBox.Value
                Else
                    Dim checkbox As CheckBox = TryCast(SelectedDocument.Editor.SelectedObject, CheckBox)
                    If checkbox IsNot Nothing Then
                        checkbox.Height = ObjectHeightBox.Value
                    Else
                        Dim textblock As TextBlock = TryCast(SelectedDocument.Editor.SelectedObject, TextBlock)
                        If textblock IsNot Nothing Then
                            textblock.Height = ObjectHeightBox.Value
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ObjectWidthBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles ObjectWidthBox.ValueChanged
        If IsLoaded Then
            Dim button As Button = TryCast(SelectedDocument.Editor.SelectedObject, Button)
            If button IsNot Nothing Then
                button.Width = ObjectWidthBox.Value
            Else
                Dim radiobutton As RadioButton = TryCast(SelectedDocument.Editor.SelectedObject, RadioButton)
                If radiobutton IsNot Nothing Then
                    radiobutton.Width = ObjectWidthBox.Value
                Else
                    Dim checkbox As CheckBox = TryCast(SelectedDocument.Editor.SelectedObject, CheckBox)
                    If checkbox IsNot Nothing Then
                        checkbox.Width = ObjectWidthBox.Value
                    Else
                        Dim textblock As TextBlock = TryCast(SelectedDocument.Editor.SelectedObject, TextBlock)
                        If textblock IsNot Nothing Then
                            textblock.Width = ObjectWidthBox.Value
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ObjectForegroundColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles ObjectForegroundColorGallery.SelectedColorChanged
        Dim button As Button = TryCast(SelectedDocument.Editor.SelectedObject, Button)
        If button IsNot Nothing Then
            button.Foreground = New SolidColorBrush(ObjectForegroundColorGallery.SelectedColor)
        Else
            Dim radiobutton As RadioButton = TryCast(SelectedDocument.Editor.SelectedObject, RadioButton)
            If radiobutton IsNot Nothing Then
                radiobutton.Foreground = New SolidColorBrush(ObjectForegroundColorGallery.SelectedColor)
            Else
                Dim checkbox As CheckBox = TryCast(SelectedDocument.Editor.SelectedObject, CheckBox)
                If checkbox IsNot Nothing Then
                    checkbox.Foreground = New SolidColorBrush(ObjectForegroundColorGallery.SelectedColor)
                Else
                    Dim textblock As TextBlock = TryCast(SelectedDocument.Editor.SelectedObject, TextBlock)
                    If textblock IsNot Nothing Then
                        textblock.Foreground = New SolidColorBrush(ObjectForegroundColorGallery.SelectedColor)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ObjectBackgroundColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles ObjectBackgroundColorGallery.SelectedColorChanged
        Dim button As Button = TryCast(SelectedDocument.Editor.SelectedObject, Button)
        If button IsNot Nothing Then
            button.Background = New SolidColorBrush(ObjectBackgroundColorGallery.SelectedColor)
        Else
            Dim radiobutton As RadioButton = TryCast(SelectedDocument.Editor.SelectedObject, RadioButton)
            If radiobutton IsNot Nothing Then
                radiobutton.Background = New SolidColorBrush(ObjectBackgroundColorGallery.SelectedColor)
            Else
                Dim checkbox As CheckBox = TryCast(SelectedDocument.Editor.SelectedObject, CheckBox)
                If checkbox IsNot Nothing Then
                    checkbox.Background = New SolidColorBrush(ObjectBackgroundColorGallery.SelectedColor)
                Else
                    Dim textblock As TextBlock = TryCast(SelectedDocument.Editor.SelectedObject, TextBlock)
                    If textblock IsNot Nothing Then
                        textblock.Background = New SolidColorBrush(ObjectBackgroundColorGallery.SelectedColor)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ObjectBorderColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles ObjectBorderColorGallery.SelectedColorChanged
        Dim button As Button = TryCast(SelectedDocument.Editor.SelectedObject, Button)
        If button IsNot Nothing Then
            button.BorderBrush = New SolidColorBrush(ObjectBorderColorGallery.SelectedColor)
        Else
            Dim radiobutton As RadioButton = TryCast(SelectedDocument.Editor.SelectedObject, RadioButton)
            If radiobutton IsNot Nothing Then
                radiobutton.BorderBrush = New SolidColorBrush(ObjectBorderColorGallery.SelectedColor)
            Else
                Dim checkbox As CheckBox = TryCast(SelectedDocument.Editor.SelectedObject, CheckBox)
                If checkbox IsNot Nothing Then
                    checkbox.BorderBrush = New SolidColorBrush(ObjectBorderColorGallery.SelectedColor)
                End If
            End If
        End If
    End Sub

    Private Sub ObjectBorderSizeBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles ObjectBorderSizeBox.ValueChanged
        If IsLoaded Then
            Dim size As Integer = Convert.ToInt32(ObjectBorderSizeBox.Value)
            Dim button As Button = TryCast(SelectedDocument.Editor.SelectedObject, Button)
            If button IsNot Nothing Then
                button.BorderThickness = New Thickness(size, size, size, size)
            Else
                Dim radiobutton As RadioButton = TryCast(SelectedDocument.Editor.SelectedObject, RadioButton)
                If radiobutton IsNot Nothing Then
                    radiobutton.BorderThickness = New Thickness(size, size, size, size)
                Else
                    Dim checkbox As CheckBox = TryCast(SelectedDocument.Editor.SelectedObject, CheckBox)
                    If checkbox IsNot Nothing Then
                        checkbox.BorderThickness = New Thickness(size, size, size, size)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ObjectTextBox_TextChanged(sender As Object, e As TextChangedEventArgs) Handles ObjectTextBox.TextChanged
        Dim button As Button = TryCast(SelectedDocument.Editor.SelectedObject, Button)
        If button IsNot Nothing Then
            button.Content = ObjectTextBox.Text
        Else
            Dim radiobutton As RadioButton = TryCast(SelectedDocument.Editor.SelectedObject, RadioButton)
            If radiobutton IsNot Nothing Then
                radiobutton.Content = ObjectTextBox.Text
            Else
                Dim checkbox As CheckBox = TryCast(SelectedDocument.Editor.SelectedObject, CheckBox)
                If checkbox IsNot Nothing Then
                    checkbox.Content = ObjectTextBox.Text
                Else
                    Dim textblock As TextBlock = TryCast(SelectedDocument.Editor.SelectedObject, TextBlock)
                    If textblock IsNot Nothing Then
                        textblock.Text = ObjectTextBox.Text
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ObjectEnabledCheckBox_Click(sender As Object, e As RoutedEventArgs) Handles ObjectEnabledCheckBox.Click
        If ObjectEnabledCheckBox.IsChecked Then
            Dim button As Button = TryCast(SelectedDocument.Editor.SelectedObject, Button)
            If button IsNot Nothing Then
                button.IsEnabled = True
            Else
                Dim radiobutton As RadioButton = TryCast(SelectedDocument.Editor.SelectedObject, RadioButton)
                If radiobutton IsNot Nothing Then
                    radiobutton.IsEnabled = True
                Else
                    Dim checkbox As CheckBox = TryCast(SelectedDocument.Editor.SelectedObject, CheckBox)
                    If checkbox IsNot Nothing Then
                        checkbox.IsEnabled = True
                    Else
                        Dim textblock As TextBlock = TryCast(SelectedDocument.Editor.SelectedObject, TextBlock)
                        If textblock IsNot Nothing Then
                            textblock.IsEnabled = True
                        End If
                    End If
                End If
            End If
        Else
            Dim button As Button = TryCast(SelectedDocument.Editor.SelectedObject, Button)
            If button IsNot Nothing Then
                button.IsEnabled = False
            Else
                Dim radiobutton As RadioButton = TryCast(SelectedDocument.Editor.SelectedObject, RadioButton)
                If radiobutton IsNot Nothing Then
                    radiobutton.IsEnabled = False
                Else
                    Dim checkbox As CheckBox = TryCast(SelectedDocument.Editor.SelectedObject, CheckBox)
                    If checkbox IsNot Nothing Then
                        checkbox.IsEnabled = False
                    Else
                        Dim textblock As TextBlock = TryCast(SelectedDocument.Editor.SelectedObject, TextBlock)
                        If textblock IsNot Nothing Then
                            textblock.IsEnabled = False
                        End If
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#End Region

#Region "Toolbar Items"

#Region "--HelpMenuItem"

    Private Sub GetHelpButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles OnlineHelpButton.Click
        Process.Start("http://documenteditor.codeplex.com/")
    End Sub

    Private Sub ReportMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ReportButton.Click
        Try
            Process.Start("mailto:semagsoft@gmail.com")
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
    End Sub

    Private Sub DonateMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles DonateButton.Click
        Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=K4QJBR4UJ3W5E&lc=US&item_name=Semagsoft&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donateCC_LG%2egif%3aNonHosted")
    End Sub

    Private Sub GetPluginsButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles GetPluginsButton.Click
        Process.Start("http://documenteditor.codeplex.com")
    End Sub

    Private Sub WebsiteMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles HomepageButton.Click
        Process.Start("http://semagsoft.com/windows/documenteditor/")
    End Sub

    Private Sub AboutMenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles AboutButton.Click
        Dim ad As New AboutDialog
        ad.Owner = Me
        ad.ShowDialog()
    End Sub

#End Region

#End Region

#End Region

#Region "Docking Manager"

    Private Sub TabCell_ChildrenCollectionChanged(sender As Object, e As EventArgs) Handles TabCell.ChildrenCollectionChanged
        'debugText.Text = dockingManager.ActiveContent.ToString + " " + TabCell.ChildrenCount.ToString

        If TabCell.SelectedContent IsNot Nothing Then
            'If TabCell.ChildrenCount = 0 Then
            '    SelectedDocument = Nothing
            'Else
            '    Dim tab As DocumentTab = TryCast(TabCell.SelectedContent, DocumentTab)
            '    If tab IsNot Nothing Then
            '        SelectedDocument = TabCell.SelectedContent
            '    Else
            '        SelectedDocument = Nothing
            '    End If
            'End If

            'SelectedDocument = TabCell.SelectedContent
        Else

            'SelectedDocument = Nothing
        End If

        If TabCell.ChildrenCount = 0 Then
            SelectedDocument = Nothing
        End If

        UpdateButtons()
    End Sub

    Private Sub dockingManager_ActiveContentChanged(sender As Object, e As EventArgs) Handles dockingManager.ActiveContentChanged
        'debugText.Text = dockingManager.ActiveContent.ToString
        If dockingManager.Layout.ActiveContent IsNot Nothing Then
            Dim g As DocumentTab = TryCast(dockingManager.Layout.ActiveContent, DocumentTab)
            If g IsNot Nothing Then
                'debugText.Text = dockingManager.ActiveContent.ToString + " " + TabCell.ChildrenCount.ToString
                SelectedDocument = dockingManager.Layout.ActiveContent
            Else
                If dockingManager.Layout.LastFocusedDocument IsNot Nothing Then
                    'SelectedDocument = dockingManager.Layout.LastFocusedDocument.Content
                    'debugText.Text = dockingManager.ActiveContent.ToString + " " + TabCell.ChildrenCount.ToString
                Else
                    SelectedDocument = Nothing
                    'debugText.Text = dockingManager.ActiveContent.ToString + " " + TabCell.ChildrenCount.ToString
                End If
                'debugText.Text = dockingManager.ActiveContent.ToString ' + " " + TabCell.ChildrenCount.ToString
            End If
        End If
        UpdateButtons()
    End Sub

    Private Sub dockingManager_DocumentClosing(sender As Object, e As Xceed.Wpf.AvalonDock.DocumentClosingEventArgs) Handles dockingManager.DocumentClosing
        Dim b As DocumentTab = e.Document
        b.IsActive = True
        If b.Editor.FileChanged Then
            Dim SaveDialog As New SaveFileDialog, fs As IO.FileStream = IO.File.Open(My.Computer.FileSystem.SpecialDirectories.Temp + "\TVPre.xaml", IO.FileMode.Create),
                tr As New TextRange(b.Editor.Document.ContentStart, b.Editor.Document.ContentEnd)
            SaveDialog.Owner = Me
            SaveDialog.SetFileInfo(b.DocName, b.Editor)
            Markup.XamlWriter.Save(b.Editor.Document, fs)
            fs.Close()
            SaveDialog.ShowDialog()
            If SaveDialog.Res = "Yes" Then
                SaveMenuItem_Click(Me, Nothing)
            ElseIf SaveDialog.Res = Nothing Then
                e.Cancel = True
            End If
        Else
            SelectedDocument = Nothing
        End If

        UpdateButtons()
    End Sub

#End Region

#Region "Statusbar"

    Private Sub StatusbarInfo_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles LinesTextBlock.PreviewMouseDown, ColumnsTextBlock.PreviewMouseDown, WordCountTextBlock.PreviewMouseDown
        If e.ChangedButton = MouseButton.Left AndAlso e.ClickCount = 2 Then
            BackStageMenu.IsOpen = True
            PropertiesMenuItem.IsSelected = True
            PropertiesTabControl.SelectedIndex = 1
        End If
    End Sub

    Private Sub zoomSlider_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles zoomSlider.ValueChanged
        If zoomSlider.IsLoaded Then
            Dim t As New ScaleTransform(zoomSlider.Value / 100, zoomSlider.Value / 100)
            Dim rulerzoom As New ScaleTransform(zoomSlider.Value / 100, 1)
            SelectedDocument.Ruler.LayoutTransform = rulerzoom
            SelectedDocument.Editor.LayoutTransform = t
            SelectedDocument.Editor.ZoomLevel = zoomSlider.Value / 100
            UpdateButtons()
        End If
    End Sub

#End Region

#Region "Misc"

    Private Sub DocPreviewScrollViewer_Loaded(sender As Object, e As RoutedEventArgs) Handles DocPreviewScrollViewer.Loaded
        DocPreviewScrollViewer.Visibility = Visibility.Collapsed
    End Sub

    Private Sub SearchPanel_FindTextBox_TextChanged(sender As Object, e As TextChangedEventArgs) Handles SearchPanel_FindTextBox.TextChanged
        If SearchPanel_FindTextBox.Text.Length > 0 Then
            SearchPanel_FindButton.IsEnabled = True
        Else
            SearchPanel_FindButton.IsEnabled = False
        End If
    End Sub

    Private Sub SearchPanel_ReplaceTextBox_TextChanged(sender As Object, e As TextChangedEventArgs) Handles SearchPanel_ReplaceTextBox.TextChanged
        If SearchPanel_ReplaceTextBox.Text.Length > 0 Then
            SearchPanel_ReplaceButton.IsEnabled = True
        Else
            SearchPanel_ReplaceButton.IsEnabled = False
        End If
    End Sub

    Private Sub SearchPanel_FindButton_Click(sender As Object, e As RoutedEventArgs) Handles SearchPanel_FindButton.Click
        Dim p As TextRange = SelectedDocument.Editor.FindWordFromPosition(SelectedDocument.Editor.CaretPosition, SearchPanel_FindTextBox.Text)
        Try
            SelectedDocument.Editor.Selection.Select(p.Start, p.End)
            SelectedDocument.Editor.Focus()
        Catch ex As Exception
            Dim m As New MessageBoxDialog("Word not found.", "Warning!", Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/info32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
    End Sub

    Private Sub SearchPanel_ReplaceButton_Click(sender As Object, e As RoutedEventArgs) Handles SearchPanel_ReplaceButton.Click
        Dim p As TextRange = SelectedDocument.Editor.FindWordFromPosition(SelectedDocument.Editor.Document.ContentStart.DocumentStart, SearchPanel_FindTextBox.Text)
        Try
            SelectedDocument.Editor.Selection.Select(p.Start, p.End)
            SelectedDocument.Editor.Selection.Text = SearchPanel_ReplaceTextBox.Text
            SelectedDocument.Editor.Focus()
        Catch ex As Exception
            Dim m As New MessageBoxDialog("Word not found.", "Warning!", Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/info32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
    End Sub

    Private Sub SearchPanel_HideMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles SearchPanel_HideMenuItem.Click
        SearchPanel.Hide()
    End Sub

    Private Sub SearchPanel_GoToButton_Click(sender As Object, e As RoutedEventArgs) Handles SearchPanel_GoToButton.Click
        SelectedDocument.Editor.GoToLine(SearchPanel_LineSpinner.Value)
        SelectedDocument.Editor.Focus()
    End Sub

    Private Sub StylesGallery_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles StylesGallery.SelectionChanged
        If StylesGallery.SelectedItem Is Styles_Normal Then
            Dim c As New Color
            c = Colors.Black
            Dim scb As New SolidColorBrush(c)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, scb)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, My.Settings.Options_DefaultFont)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, My.Settings.Options_DefaultFontSize)
        ElseIf StylesGallery.SelectedItem Is Styles_NoSpacing Then
            Dim c As New Color
            c = Colors.Black
            Dim scb As New SolidColorBrush(c)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, scb)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, My.Settings.Options_DefaultFont)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, My.Settings.Options_DefaultFontSize)
        ElseIf StylesGallery.SelectedItem Is Styles_Heading1 Then
            Dim c As New Color
            c = ColorConverter.ConvertFromString("#2e74b5")
            Dim scb As New SolidColorBrush(c)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, scb)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, My.Settings.Options_DefaultFont)
            Dim fontSize As Double = 19
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSize)
        ElseIf StylesGallery.SelectedItem Is Styles_Heading2 Then
            Dim c As New Color
            c = ColorConverter.ConvertFromString("#2e74b5")
            Dim scb As New SolidColorBrush(c)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, scb)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, My.Settings.Options_DefaultFont)
            Dim fontSize As Double = 16
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSize)
        ElseIf StylesGallery.SelectedItem Is Styles_Heading3 Then
            Dim c As New Color
            c = ColorConverter.ConvertFromString("#1f4d78")
            Dim scb As New SolidColorBrush(c)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, scb)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, My.Settings.Options_DefaultFont)
            Dim fontSize As Double = 15
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSize)
            'ElseIf StylesGallery.SelectedItem Is Styles_Heading4 Then
            'ElseIf StylesGallery.SelectedItem Is Styles_Heading5 Then
        ElseIf StylesGallery.SelectedItem Is Styles_Title Then
            Dim c As New Color
            c = Colors.Black
            Dim scb As New SolidColorBrush(c)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, scb)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, My.Settings.Options_DefaultFont)
            Dim fontSize As Double = 28
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSize)
        ElseIf StylesGallery.SelectedItem Is Styles_Title Then
            Dim c As New Color
            c = Colors.Black
            Dim scb As New SolidColorBrush(c)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, scb)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, My.Settings.Options_DefaultFont)
            Dim fontSize As Double = 28
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSize)
        ElseIf StylesGallery.SelectedItem Is Styles_Subtitle Then
            Dim c As New Color
            c = Colors.Black
            Dim scb As New SolidColorBrush(c)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, scb)
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, My.Settings.Options_DefaultFont)
            Dim fontSize As Double = 14
            SelectedDocument.Editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSize)
        End If
        UpdateSelected()
    End Sub

    Private Sub SearchPanel_FindTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles SearchPanel_FindTextBox.KeyDown
        If SearchPanel_FindButton.IsEnabled Then
            SearchPanel_FindButton_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub SearchPanel_ReplaceTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles SearchPanel_ReplaceTextBox.KeyDown
        If SearchPanel_ReplaceButton.IsEnabled Then
            SearchPanel_ReplaceButton_Click(Nothing, Nothing)
        End If
    End Sub

#End Region

End Class