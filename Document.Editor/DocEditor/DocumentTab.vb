Public Class DocumentTab
    Inherits Xceed.Wpf.AvalonDock.Layout.LayoutDocument

#Region "Publics"

    Public DocName As String = Nothing, FileFormat As String = "Rich Text Document"
    Public BackgroundBrush As Brush

#Region "Events"

    Public Event Close As RoutedEventHandler
    Public Shared UpdateSelected As RoutedEvent = EventManager.RegisterRoutedEvent("UpdateSelected", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab))

    Public Shared InsertObjectEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertObject", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        InsertShapeEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertShape", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        InsertImageEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertImage", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        InsertLinkEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertLink", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        InsertFlowDocumentEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertFlowDocument", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        InsertRichTextFileEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertRichTextFile", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        InsertTextFileEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertTextFile", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        InsertSymbolEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertSymbol", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        InsertTableEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertTable", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        InsertVideoEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertVideo", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        InsertHorizontalLineEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertHorizontalLine", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        InsertHeaderEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertHeader", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        InsertFooterEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertFooter", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        InsertDateEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertDate", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        InsertTimeEvent As RoutedEvent = EventManager.RegisterRoutedEvent("InsertTime", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        ClearFormattingEvent As RoutedEvent = EventManager.RegisterRoutedEvent("ClearFormatting", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        FontEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Font", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        FontSizeEvent As RoutedEvent = EventManager.RegisterRoutedEvent("FontSize", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        FontColorEvent As RoutedEvent = EventManager.RegisterRoutedEvent("FontColor", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        HighlightColorEvent As RoutedEvent = EventManager.RegisterRoutedEvent("HighlightColor", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        BoldEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Bold", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        ItalicEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Italic", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        UnderlineEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Underline", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        StrikethroughEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Strikethrough", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        SubscriptEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Subscript", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        SuperscriptEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Superscript", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        IndentMoreEvent As RoutedEvent = EventManager.RegisterRoutedEvent("IndentMore", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        IndentLessEvent As RoutedEvent = EventManager.RegisterRoutedEvent("IndentLess", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        BulletListEvent As RoutedEvent = EventManager.RegisterRoutedEvent("BulletList", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        NumberListEvent As RoutedEvent = EventManager.RegisterRoutedEvent("NumberList", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        AlignLeftEvent As RoutedEvent = EventManager.RegisterRoutedEvent("AlignLeft", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        AlignCenterEvent As RoutedEvent = EventManager.RegisterRoutedEvent("AlignCenter", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        AlignRightEvent As RoutedEvent = EventManager.RegisterRoutedEvent("AlignRight", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        AlignJustifyEvent As RoutedEvent = EventManager.RegisterRoutedEvent("AlignJustify", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        LineSpacingEvent As RoutedEvent = EventManager.RegisterRoutedEvent("LineSpacing", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        LeftToRightEvent As RoutedEvent = EventManager.RegisterRoutedEvent("LeftToRight", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        RightToLeftEvent As RoutedEvent = EventManager.RegisterRoutedEvent("RightToLeft", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        UndoEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Undo", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        RedoEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Redo", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        CutEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Cut", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        CopyEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Copy", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        PasteEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Paste", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        DeleteEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Delete", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        SelectAllEvent As RoutedEvent = EventManager.RegisterRoutedEvent("SelectAll", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        FindEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Find", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        ReplaceEvent As RoutedEvent = EventManager.RegisterRoutedEvent("Replace", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab)),
        GoToEvent As RoutedEvent = EventManager.RegisterRoutedEvent("GoTo", RoutingStrategy.Tunnel, GetType(RoutedEventHandler), GetType(DocumentTab))

#End Region

    Public WithEvents Editor As New DocumentEditor, VSV As New ScrollViewer, DocumentView As New Grid, Vis As New Grid
    Public MyIndex As Integer

    Public ContentView As New Grid
    Public Ruler As New DockPanel, hruler As New Semagsoft.DocRuler.Ruler

    Public WithEvents TBoxContextMenu As ContextMenu,
        ObjectMenuItem As New MenuItem, ShapeMenuItem As New MenuItem, ImageMenuItem As New MenuItem, LinkMenuItem As New MenuItem, FlowDocumentMenuItem As New MenuItem, RichTextFileMenuItem As New MenuItem, TextFileMenuItem As New MenuItem, SymbolMenuItem As New MenuItem, TableMenuItem As New MenuItem, VideoMenuItem As New MenuItem, HorizontalLineMenuItem As New MenuItem, HeaderMenuItem As New MenuItem, FooterMenuItem As New MenuItem, DateMenuItem As New MenuItem, TimeMenuItem As New MenuItem,
        ClearFormattingMenuItem As New MenuItem, FontMenuItem As New MenuItem, FontSizeMenuItem As New MenuItem, FontColorMenuItem As New MenuItem, HighlightMenuItem As New MenuItem, BoldMenuItem As New MenuItem, ItalicMenuItem As New MenuItem, UnderlineMenuItem As New MenuItem, StrikethroughMenuItem As New MenuItem, SubscriptMenuItem As New MenuItem, SuperscriptMenuItem As New MenuItem, IndentMoreMenuItem As New MenuItem, IndentLessMenuItem As New MenuItem, BulletListMenuItem As New MenuItem, NumberListMenuItem As New MenuItem, AlignLeftMenuItem As New MenuItem, AlignCenterMenuItem As New MenuItem, AlignRightMenuItem As New MenuItem, AlignJustifyMenuItem As New MenuItem, LineSpacingMenuItem As New MenuItem, LefttoRightMenuItem As New MenuItem, RighttoLeftMenuItem As New MenuItem,
        UndoMenuItem As New MenuItem, RedoMenuItem As New MenuItem, CutMenuItem As New MenuItem, CopyMenuItem As New MenuItem, PasteMenuItem As New MenuItem, DeleteMenuItem As New MenuItem, SelectAllMenuItem As New MenuItem, FindMenuItem As New MenuItem, ReplaceMenuItem As New MenuItem, GoToMenuItem As New MenuItem, SetCaseMenuItem As New MenuItem, UppercaseMenuItem As New MenuItem, LowercaseMenuItem As New MenuItem

#End Region

#Region "FileBox"

    Public Sub New(ByVal name As String, ByVal background As Brush)
        CanFloat = False
        SetFileType(".xaml")
        Title = name
        BackgroundBrush = background
        DocName = name
        If My.Settings.Options_SpellCheck Then
            SpellCheck.SetIsEnabled(Editor, True)
        End If
        VSV.HorizontalAlignment = Windows.HorizontalAlignment.Stretch
        VSV.VerticalAlignment = Windows.VerticalAlignment.Stretch
        Editor.HorizontalAlignment = Windows.HorizontalAlignment.Center
        Editor.VerticalAlignment = Windows.VerticalAlignment.Top
        Editor.BorderBrush = New SolidColorBrush(Color.FromRgb(198, 198, 198))
        Editor.BorderThickness = New Thickness(1, 1, 1, 1)
        Editor.Width = Editor.Document.PageWidth
        Editor.Height = Editor.Document.PageHeight
        'DocumentView.VerticalAlignment = Windows.VerticalAlignment.Top
        'DocumentView.HorizontalAlignment = Windows.HorizontalAlignment.Center
        'DocumentView.Width = Editor.Width
        'DocumentView.Height = Editor.Height
        'DocumentView.Margin = New Thickness(0, -16, 0, 32)
        'DocumentView.Background = Brushes.Black
        'Dim editoreffect As New Effects.DropShadowEffect
        'editoreffect.BlurRadius = 16
        'editoreffect.Direction = 90
        'Editor.Effect = editoreffect
        Editor.ClipToBounds = False
        Editor.Margin = New Thickness(0, 16, 0, 16)
        If My.Settings.MainWindow_ShowRuler Then
            VSV.Margin = New Thickness(0, 23, 0, 0)
        Else
            If My.Settings.Options_Theme = 3 Then
                VSV.Margin = New Thickness(0, 2, 0, 2)
            Else
                VSV.Margin = New Thickness(0, 0, 0, 0)
            End If
            Ruler.Visibility = Windows.Visibility.Collapsed
        End If
        If My.Settings.Options_RulerMeasurement = 0 Then
            hruler.Unit = Semagsoft.DocRuler.Unit.Inch
        Else
            hruler.Unit = Semagsoft.DocRuler.Unit.Cm
        End If
        hruler.AutoSize = True
        hruler.Height = 24
        hruler.Width = Editor.Width
        hruler.ClipToBounds = True
        hruler.Margin = New Thickness(0, -1, 0, 0)
        Ruler.Children.Add(hruler)
        Ruler.VerticalAlignment = Windows.VerticalAlignment.Top
        Ruler.Width = Editor.Width
        Vis.Background = Brushes.Black
        'DocumentView.Children.Add(Vis)
        'DocumentView.Children.Add(Editor)
        VSV.Content = Editor
        'Vis.Width = Editor.Width
        'Vis.Height = Editor.Height
        'Vis.Background = Brushes.Black
        ContentView.Children.Add(Ruler)
        ContentView.Children.Add(VSV)
        'ContentView.Children.Add(Vis)
        If My.Settings.Options_Theme = 3 Then
            Dim b As New Border
            b.BorderBrush = Brushes.LightGray
            b.BorderThickness = New Thickness(0, 1, 0, 1)
            b.Margin = New Thickness(0, -2, 0, -2)
            ContentView.Children.Add(b)
            ContentView.Background = New LinearGradientBrush(Colors.White, Color.FromRgb(240, 240, 240), 90)
            ContentView.Margin = New Thickness(2, -6, 0, 0)
        Else
            ContentView.Margin = New Thickness(0, 1, 0, 0)
        End If
        Content = ContentView
        VSV.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        Editor.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled
        Editor.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        TBoxContextMenu = New ContextMenu
        Dim insertimg As New BitmapImage(New Uri("pack://application:,,,/Images/Common/add16.png")), inserticon As New Image
        Dim tableimg As New BitmapImage(New Uri("pack://application:,,,/Images/Insert/table16.png")), tableicon As New Image
        Dim dateimg As New BitmapImage(New Uri("pack://application:,,,/Images/Insert/date16.png")), dateicon As New Image
        Dim timeimg As New BitmapImage(New Uri("pack://application:,,,/Images/Insert/time16.png")), timeicon As New Image
        Dim imageimg As New BitmapImage(New Uri("pack://application:,,,/Images/Insert/image16.png")), imageicon As New Image
        Dim videoimg As New BitmapImage(New Uri("pack://application:,,,/Images/Insert/video16.png")), videoicon As New Image
        Dim objectimg As New BitmapImage(New Uri("pack://application:,,,/Images/Insert/object16.png")), objecticon As New Image
        Dim shapeimg As New BitmapImage(New Uri("pack://application:,,,/Images/Insert/shape16.png")), shapeicon As New Image
        Dim linkimg As New BitmapImage(New Uri("pack://application:,,,/Images/Insert/link16.png")), linkicon As New Image
        Dim flowdocumentimg As New BitmapImage(New Uri("pack://application:,,,/Images/Tab/xaml16.png")), flowdocumenticon As New Image
        Dim richtextfileimg As New BitmapImage(New Uri("pack://application:,,,/Images/Tab/rtf16.png")), richtextfileicon As New Image
        Dim textfileimg As New BitmapImage(New Uri("pack://application:,,,/Images/Insert/textfile16.png")), textfileicon As New Image
        Dim symbolimg As New BitmapImage(New Uri("pack://application:,,,/Images/Insert/symbol16.png")), symbolicon As New Image
        Dim lineimg As New BitmapImage(New Uri("pack://application:,,,/Images/Insert/horizontalline16.png")), lineicon As New Image
        Dim headerimg As New BitmapImage(New Uri("pack://application:,,,/Images/Insert/header16.png")), headericon As New Image
        Dim footerimg As New BitmapImage(New Uri("pack://application:,,,/Images/Insert/footer16.png")), footericon As New Image
        Dim clearformattingimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/clearformatting16.png")), clearformattingicon As New Image
        Dim fontcolorimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/fontfacecolor16.png")), fontcoloricon As New Image
        Dim fonthighlightcolorimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/fontcolor16.png")), fonthighlightcoloricon As New Image
        Dim boldimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/bold16.png")), boldicon As New Image
        Dim italicimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/italic16.png")), italicicon As New Image
        Dim Underlineimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/underline16.png")), underlineicon As New Image
        Dim strikethroughimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/strikethrough16.png")), strikethroughicon As New Image
        Dim subscriptimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/subscript16.png")), subscripticon As New Image
        Dim superscriptimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/superscript16.png")), superscripticon As New Image
        Dim indentmoreimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/indentmore16.png")), indentmoreicon As New Image
        Dim indentlessimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/indentless16.png")), indentlessicon As New Image
        Dim bulletlistimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/bulletlist16.png")), bulletlisticon As New Image
        Dim numberlistimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/numberlist16.png")), numberlisticon As New Image
        Dim alignleftimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/left16.png")), alignlefticon As New Image
        Dim aligncenterimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/center16.png")), aligncentericon As New Image
        Dim alignrightimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/right16.png")), alignrighticon As New Image
        Dim alignjustifyimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/justify16.png")), alignjustifyicon As New Image
        Dim linespacingimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/linespacing16.png")), linespacingicon As New Image
        Dim ltrimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/ltr16.png")), ltricon As New Image
        Dim rtlimg As New BitmapImage(New Uri("pack://application:,,,/Images/Format/rtl16.png")), rtlicon As New Image
        Dim undoimg As New BitmapImage(New Uri("pack://application:,,,/Images/Edit/undo16.png")), undoicon As New Image
        Dim redoimg As New BitmapImage(New Uri("pack://application:,,,/Images/Edit/redo16.png")), redoicon As New Image
        Dim cutimg As New BitmapImage(New Uri("pack://application:,,,/Images/Edit/cut16.png")), cuticon As New Image
        Dim copyimg As New BitmapImage(New Uri("pack://application:,,,/Images/Edit/copy16.png")), copyicon As New Image
        Dim pasteimg As New BitmapImage(New Uri("pack://application:,,,/Images/Edit/paste16.png")), pasteicon As New Image
        Dim deleteimg As New BitmapImage(New Uri("pack://application:,,,/Images/Edit/delete16.png")), deleteicon As New Image
        Dim selectallimg As New BitmapImage(New Uri("pack://application:,,,/Images/Edit/selectall16.png")), selectallicon As New Image
        Dim findimg As New BitmapImage(New Uri("pack://application:,,,/Images/Edit/find16.png")), findicon As New Image
        Dim replaceimg As New BitmapImage(New Uri("pack://application:,,,/Images/Edit/findreplace16.png")), Replaceicon As New Image
        Dim gotoimg As New BitmapImage(New Uri("pack://application:,,,/Images/Edit/goto16.png")), gotoicon As New Image
        Dim uppercaseimg As New BitmapImage(New Uri("pack://application:,,,/Images/Edit/uppercase16.png")), uppercaseicon As New Image
        Dim lowercaseimg As New BitmapImage(New Uri("pack://application:,,,/Images/Edit/lowercase16.png")), lowercaseicon As New Image
        inserticon.Width = 16
        inserticon.Height = 16
        inserticon.Source = insertimg
        tableicon.Width = 16
        tableicon.Height = 16
        tableicon.Source = tableimg
        dateicon.Width = 16
        dateicon.Height = 16
        dateicon.Source = dateimg
        timeicon.Width = 16
        timeicon.Height = 16
        timeicon.Source = timeimg
        imageicon.Width = 16
        imageicon.Height = 16
        imageicon.Source = imageimg
        videoicon.Width = 16
        videoicon.Height = 16
        videoicon.Source = videoimg
        objecticon.Width = 16
        objecticon.Height = 16
        objecticon.Source = objectimg
        shapeicon.Width = 16
        shapeicon.Height = 16
        shapeicon.Source = shapeimg
        linkicon.Width = 16
        linkicon.Height = 16
        linkicon.Source = linkimg
        flowdocumenticon.Width = 16
        flowdocumenticon.Height = 16
        flowdocumenticon.Source = flowdocumentimg
        richtextfileicon.Width = 16
        richtextfileicon.Height = 16
        richtextfileicon.Source = richtextfileimg
        textfileicon.Width = 16
        textfileicon.Height = 16
        textfileicon.Source = textfileimg
        symbolicon.Width = 16
        symbolicon.Height = 16
        symbolicon.Source = symbolimg
        lineicon.Width = 16
        lineicon.Height = 16
        lineicon.Source = lineimg
        headericon.Width = 16
        headericon.Height = 16
        headericon.Source = headerimg
        footericon.Width = 16
        footericon.Height = 16
        footericon.Source = footerimg
        clearformattingicon.Width = 16
        clearformattingicon.Height = 16
        clearformattingicon.Source = clearformattingimg
        fontcoloricon.Width = 16
        fontcoloricon.Height = 16
        fontcoloricon.Source = fontcolorimg
        fonthighlightcoloricon.Width = 16
        fonthighlightcoloricon.Height = 16
        fonthighlightcoloricon.Source = fonthighlightcolorimg
        boldicon.Width = 16
        boldicon.Height = 16
        boldicon.Source = boldimg
        italicicon.Width = 16
        italicicon.Height = 16
        italicicon.Source = italicimg
        underlineicon.Width = 16
        underlineicon.Height = 16
        underlineicon.Source = Underlineimg
        strikethroughicon.Width = 16
        strikethroughicon.Height = 16
        strikethroughicon.Source = strikethroughimg
        subscripticon.Width = 16
        subscripticon.Height = 16
        subscripticon.Source = subscriptimg
        superscripticon.Width = 16
        superscripticon.Height = 16
        superscripticon.Source = superscriptimg
        indentmoreicon.Width = 16
        indentmoreicon.Height = 16
        indentmoreicon.Source = indentmoreimg
        indentlessicon.Width = 16
        indentlessicon.Height = 16
        indentlessicon.Source = indentlessimg
        bulletlisticon.Width = 16
        bulletlisticon.Height = 16
        bulletlisticon.Source = bulletlistimg
        numberlisticon.Width = 16
        numberlisticon.Height = 16
        numberlisticon.Source = numberlistimg
        alignlefticon.Width = 16
        alignlefticon.Height = 16
        alignlefticon.Source = alignleftimg
        aligncentericon.Width = 16
        aligncentericon.Height = 16
        aligncentericon.Source = aligncenterimg
        alignrighticon.Width = 16
        alignrighticon.Height = 16
        alignrighticon.Source = alignrightimg
        alignjustifyicon.Width = 16
        alignjustifyicon.Height = 16
        alignjustifyicon.Source = alignjustifyimg
        linespacingicon.Width = 16
        linespacingicon.Height = 16
        linespacingicon.Source = linespacingimg
        ltricon.Width = 16
        ltricon.Height = 16
        ltricon.Source = ltrimg
        rtlicon.Width = 16
        rtlicon.Height = 16
        rtlicon.Source = rtlimg
        undoicon.Width = 16
        undoicon.Height = 16
        undoicon.Source = undoimg
        redoicon.Width = 16
        redoicon.Height = 16
        redoicon.Source = redoimg
        cuticon.Width = 16
        cuticon.Height = 16
        cuticon.Source = cutimg
        copyicon.Width = 16
        copyicon.Height = 16
        copyicon.Source = copyimg
        pasteicon.Width = 16
        pasteicon.Height = 16
        pasteicon.Source = pasteimg
        deleteicon.Width = 16
        deleteicon.Height = 16
        deleteicon.Source = deleteimg
        selectallicon.Width = 16
        selectallicon.Height = 16
        selectallicon.Source = selectallimg
        findicon.Width = 16
        findicon.Height = 16
        findicon.Source = findimg
        Replaceicon.Width = 16
        Replaceicon.Height = 16
        Replaceicon.Source = replaceimg
        gotoicon.Width = 16
        gotoicon.Height = 16
        gotoicon.Source = gotoimg
        uppercaseicon.Width = 16
        uppercaseicon.Height = 16
        uppercaseicon.Source = uppercaseimg
        lowercaseicon.Width = 16
        lowercaseicon.Height = 16
        lowercaseicon.Source = lowercaseimg
        ObjectMenuItem.Header = "Object"
        ShapeMenuItem.Header = "Shape"
        ImageMenuItem.Header = "Image"
        LinkMenuItem.Header = "Link"
        FlowDocumentMenuItem.Header = "FlowDocument"
        RichTextFileMenuItem.Header = "Rich Text File"
        TextFileMenuItem.Header = "Text File"
        SymbolMenuItem.Header = "Symbol"
        TableMenuItem.Header = "Table"
        TableMenuItem.Icon = tableicon
        DateMenuItem.Icon = dateicon
        TimeMenuItem.Icon = timeicon
        ImageMenuItem.Icon = imageicon
        VideoMenuItem.Icon = videoicon
        ObjectMenuItem.Icon = objecticon
        ShapeMenuItem.Icon = shapeicon
        LinkMenuItem.Icon = linkicon
        FlowDocumentMenuItem.Icon = flowdocumenticon
        RichTextFileMenuItem.Icon = richtextfileicon
        TextFileMenuItem.Icon = textfileicon
        SymbolMenuItem.Icon = symbolicon
        HorizontalLineMenuItem.Icon = lineicon
        HeaderMenuItem.Icon = headericon
        FooterMenuItem.Icon = footericon
        VideoMenuItem.Header = "Video"
        HorizontalLineMenuItem.Header = "Horizontal Line"
        HeaderMenuItem.Header = "Header"
        FooterMenuItem.Header = "Footer"
        DateMenuItem.Header = "Date"
        TimeMenuItem.Header = "Time"
        ClearFormattingMenuItem.Icon = clearformattingicon
        ClearFormattingMenuItem.Header = "Clear Formatting"
        FontMenuItem.Header = "Font"
        FontSizeMenuItem.Header = "Font Size"
        FontColorMenuItem.Icon = fontcoloricon
        FontColorMenuItem.Header = "Font Color"
        HighlightMenuItem.Icon = fonthighlightcoloricon
        HighlightMenuItem.Header = "Highlight Color"
        BoldMenuItem.Icon = boldicon
        BoldMenuItem.Header = "Bold"
        ItalicMenuItem.Icon = italicicon
        ItalicMenuItem.Header = "Italic"
        UnderlineMenuItem.Icon = underlineicon
        UnderlineMenuItem.Header = "Underline"
        StrikethroughMenuItem.Icon = strikethroughicon
        StrikethroughMenuItem.Header = "Strikethrough"
        SubscriptMenuItem.Icon = subscripticon
        SubscriptMenuItem.Header = "Subscript"
        SuperscriptMenuItem.Icon = superscripticon
        SuperscriptMenuItem.Header = "Superscript"
        IndentMoreMenuItem.Icon = indentmoreicon
        IndentMoreMenuItem.Header = "Indent More"
        IndentLessMenuItem.Icon = indentlessicon
        IndentLessMenuItem.Header = "Indent Less"
        BulletListMenuItem.Icon = bulletlisticon
        BulletListMenuItem.Header = "Bullet List"
        NumberListMenuItem.Icon = numberlisticon
        NumberListMenuItem.Header = "Number List"
        AlignLeftMenuItem.Icon = alignlefticon
        AlignLeftMenuItem.Header = "Align Left"
        AlignCenterMenuItem.Icon = aligncentericon
        AlignCenterMenuItem.Header = "Align Center"
        AlignRightMenuItem.Icon = alignrighticon
        AlignRightMenuItem.Header = "Align Right"
        AlignJustifyMenuItem.Icon = alignjustifyicon
        AlignJustifyMenuItem.Header = "Align Justify"
        LineSpacingMenuItem.Icon = linespacingicon
        LineSpacingMenuItem.Header = "Line Spacing"
        LefttoRightMenuItem.Icon = ltricon
        LefttoRightMenuItem.Header = "Left To Right"
        RighttoLeftMenuItem.Icon = rtlicon
        RighttoLeftMenuItem.Header = "Right To Left"
        UndoMenuItem.Header = "Undo"
        UndoMenuItem.InputGestureText = "Ctrl+Z"
        UndoMenuItem.ToolTip = "Undo (Ctrl + Z)"
        UndoMenuItem.Icon = undoicon
        RedoMenuItem.Header = "Redo"
        RedoMenuItem.InputGestureText = "Ctrl+Y"
        RedoMenuItem.ToolTip = "Redo (Ctrl + Y)"
        RedoMenuItem.Icon = redoicon
        CutMenuItem.Header = "Cut"
        CutMenuItem.InputGestureText = "Ctrl+X"
        CutMenuItem.ToolTip = "Cut (Ctrl + X)"
        CutMenuItem.Icon = cuticon
        CopyMenuItem.Header = "Copy"
        CopyMenuItem.InputGestureText = "Ctrl+C"
        CopyMenuItem.ToolTip = "Copy (Ctrl + C)"
        CopyMenuItem.Icon = copyicon
        PasteMenuItem.Header = "Paste"
        PasteMenuItem.InputGestureText = "Ctrl+V"
        PasteMenuItem.ToolTip = "Paste (Ctrl + V)"
        PasteMenuItem.Icon = pasteicon
        DeleteMenuItem.Header = "Delete"
        DeleteMenuItem.InputGestureText = "Del"
        DeleteMenuItem.ToolTip = "Delete (Del)"
        DeleteMenuItem.Icon = deleteicon
        SelectAllMenuItem.Header = "Select All"
        SelectAllMenuItem.InputGestureText = "Ctrl+A"
        SelectAllMenuItem.ToolTip = "Select All (Ctrl + A)"
        SelectAllMenuItem.Icon = selectallicon
        FindMenuItem.Header = "Find"
        FindMenuItem.InputGestureText = "Ctrl+F"
        FindMenuItem.ToolTip = "Find (Ctrl + F)"
        FindMenuItem.Icon = findicon
        ReplaceMenuItem.Header = "Replace"
        ReplaceMenuItem.InputGestureText = "Ctrl+H"
        ReplaceMenuItem.ToolTip = "Replace (Ctrl + H)"
        ReplaceMenuItem.Icon = Replaceicon
        GoToMenuItem.Header = "Go To"
        GoToMenuItem.InputGestureText = "Ctrl+G"
        GoToMenuItem.ToolTip = "Go To (Ctrl + G)"
        GoToMenuItem.Icon = gotoicon
        SetCaseMenuItem.Header = "Set Case"
        UppercaseMenuItem.Icon = uppercaseicon
        UppercaseMenuItem.Header = "Uppercase"
        LowercaseMenuItem.Icon = lowercaseicon
        LowercaseMenuItem.Header = "Lowercase"
        Dim insertmenuitem As New MenuItem
        insertmenuitem.Header = "Insert"
        insertmenuitem.Icon = inserticon
        insertmenuitem.Items.Add(TableMenuItem)
        insertmenuitem.Items.Add(New Separator)
        insertmenuitem.Items.Add(DateMenuItem)
        insertmenuitem.Items.Add(TimeMenuItem)
        insertmenuitem.Items.Add(New Separator)
        insertmenuitem.Items.Add(ImageMenuItem)
        insertmenuitem.Items.Add(VideoMenuItem)
        insertmenuitem.Items.Add(New Separator)
        insertmenuitem.Items.Add(ObjectMenuItem)
        insertmenuitem.Items.Add(ShapeMenuItem)
        insertmenuitem.Items.Add(LinkMenuItem)
        insertmenuitem.Items.Add(FlowDocumentMenuItem)
        insertmenuitem.Items.Add(RichTextFileMenuItem)
        insertmenuitem.Items.Add(TextFileMenuItem)
        insertmenuitem.Items.Add(SymbolMenuItem)
        insertmenuitem.Items.Add(HorizontalLineMenuItem)
        insertmenuitem.Items.Add(HeaderMenuItem)
        insertmenuitem.Items.Add(FooterMenuItem)
        Dim formatmenuitem As New MenuItem
        formatmenuitem.Header = "Format"
        formatmenuitem.Items.Add(ClearFormattingMenuItem)
        formatmenuitem.Items.Add(New Separator)
        Dim fontmenuitem2 As New MenuItem
        fontmenuitem2.Header = "Font"
        fontmenuitem2.Items.Add(FontMenuItem)
        fontmenuitem2.Items.Add(FontSizeMenuItem)
        fontmenuitem2.Items.Add(FontColorMenuItem)
        fontmenuitem2.Items.Add(HighlightMenuItem)
        formatmenuitem.Items.Add(fontmenuitem2)
        Dim stylemenuitem As New MenuItem
        stylemenuitem.Header = "Style"
        stylemenuitem.Items.Add(BoldMenuItem)
        stylemenuitem.Items.Add(ItalicMenuItem)
        stylemenuitem.Items.Add(UnderlineMenuItem)
        stylemenuitem.Items.Add(StrikethroughMenuItem)
        formatmenuitem.Items.Add(stylemenuitem)
        formatmenuitem.Items.Add(SubscriptMenuItem)
        formatmenuitem.Items.Add(SuperscriptMenuItem)
        formatmenuitem.Items.Add(IndentMoreMenuItem)
        formatmenuitem.Items.Add(IndentLessMenuItem)
        formatmenuitem.Items.Add(BulletListMenuItem)
        formatmenuitem.Items.Add(NumberListMenuItem)
        Dim alignmenuitem As New MenuItem
        alignmenuitem.Header = "Align"
        alignmenuitem.Items.Add(AlignLeftMenuItem)
        alignmenuitem.Items.Add(AlignCenterMenuItem)
        alignmenuitem.Items.Add(AlignRightMenuItem)
        alignmenuitem.Items.Add(AlignJustifyMenuItem)
        formatmenuitem.Items.Add(alignmenuitem)
        formatmenuitem.Items.Add(LineSpacingMenuItem)
        formatmenuitem.Items.Add(LefttoRightMenuItem)
        formatmenuitem.Items.Add(RighttoLeftMenuItem)
        TBoxContextMenu.Items.Add(insertmenuitem)
        TBoxContextMenu.Items.Add(formatmenuitem)
        TBoxContextMenu.Items.Add(New Separator)
        TBoxContextMenu.Items.Add(UndoMenuItem)
        TBoxContextMenu.Items.Add(RedoMenuItem)
        TBoxContextMenu.Items.Add(New Separator)
        TBoxContextMenu.Items.Add(CutMenuItem)
        TBoxContextMenu.Items.Add(CopyMenuItem)
        TBoxContextMenu.Items.Add(PasteMenuItem)
        TBoxContextMenu.Items.Add(DeleteMenuItem)
        TBoxContextMenu.Items.Add(New Separator)
        TBoxContextMenu.Items.Add(SelectAllMenuItem)
        TBoxContextMenu.Items.Add(New Separator)
        TBoxContextMenu.Items.Add(FindMenuItem)
        TBoxContextMenu.Items.Add(ReplaceMenuItem)
        TBoxContextMenu.Items.Add(GoToMenuItem)
        TBoxContextMenu.Items.Add(New Separator)
        TBoxContextMenu.Items.Add(SetCaseMenuItem)
        SetCaseMenuItem.Items.Add(UppercaseMenuItem)
        SetCaseMenuItem.Items.Add(LowercaseMenuItem)
        Editor.ContextMenu = TBoxContextMenu
    End Sub

    Public Sub SetPageSize(height As Double, width As Double)
        Editor.Document.PageHeight = height
        Editor.Height = height
        Editor.Document.PageWidth = width
        Editor.Width = width
        hruler.Width = width
        Ruler.Width = width
    End Sub

#End Region

#Region "Header"

    Public Sub SetFileType(ByVal s As String)
        Dim img As New BitmapImage
        Dim tip As New Fluent.ScreenTip
        tip.Title = DocName
        img.BeginInit()
        If s.ToLower = ".xaml" Then
            img.UriSource = New Uri("pack://application:,,,/Images/Tab/xaml16.png")
        ElseIf s.ToLower = ".rtf" Then
            img.UriSource = New Uri("pack://application:,,,/Images/Tab/rtf16.png")
        ElseIf s.ToLower = ".html" OrElse s.ToLower.ToLower = ".htm" Then
            img.UriSource = New Uri("pack://application:,,,/Images/Tab/html16.png")
        Else
            img.UriSource = New Uri("pack://application:,,,/Images/Tab/txt16.png")
        End If
        img.EndInit()
        IconSource = img
        'HeaderContent.FileTypeImage.ToolTip = s
        FileFormat = s
    End Sub

    Public Sub SetDocumentTitle(ByVal text As String)
        If Editor.FileChanged Then
            If Editor.IsReadOnly Then
                Title = text + "* (Read-only)"
            Else
                Title = text + "*"
            End If
        Else
            If Editor.IsReadOnly Then
                Title = text + " (Read-only)"
            Else
                Title = text
            End If
        End If
    End Sub

#End Region

#Region "TBox"

#Region "Context Menu"

    Private Sub TBoxContextMenu_Opened(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles TBoxContextMenu.Opened
        If Editor.CanUndo Then
            UndoMenuItem.IsEnabled = True
        Else
            UndoMenuItem.IsEnabled = False
        End If
        If Editor.CanRedo Then
            RedoMenuItem.IsEnabled = True
        Else
            RedoMenuItem.IsEnabled = False
        End If
        If My.Computer.Clipboard.ContainsText Then
            PasteMenuItem.IsEnabled = True
        Else
            PasteMenuItem.IsEnabled = False
        End If
    End Sub

    Private Sub ObjectMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles ObjectMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertObjectEvent))
    End Sub

    Private Sub ShapeMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles ShapeMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertShapeEvent))
    End Sub

    Private Sub ImageMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles ImageMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertImageEvent))
    End Sub

    Private Sub LinkMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles LinkMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertLinkEvent))
    End Sub

    Private Sub FlowDocumentMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles FlowDocumentMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertFlowDocumentEvent))
    End Sub

    Private Sub RichTextFileMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles RichTextFileMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertRichTextFileEvent))
    End Sub

    Private Sub TextFileMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles TextFileMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertTextFileEvent))
    End Sub

    Private Sub SymbolMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles SymbolMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertSymbolEvent))
    End Sub

    Private Sub TableMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles TableMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertTableEvent))
    End Sub

    Private Sub VideoMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles VideoMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertVideoEvent))
    End Sub

    Private Sub HorizontalLineMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles HorizontalLineMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertHorizontalLineEvent))
    End Sub

    Private Sub HeaderMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles HeaderMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertHeaderEvent))
    End Sub

    Private Sub FooterMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles FooterMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertFooterEvent))
    End Sub

    Private Sub DateMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles DateMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertDateEvent))
    End Sub

    Private Sub TimeMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles TimeMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(InsertTimeEvent))
    End Sub

    Private Sub ClearFormattingMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles ClearFormattingMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(ClearFormattingEvent))
    End Sub

    Private Sub FontMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles FontMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(FontEvent))
    End Sub

    Private Sub FontSizeMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles FontSizeMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(FontSizeEvent))
    End Sub

    Private Sub FontColorMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles FontColorMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(FontColorEvent))
    End Sub

    Private Sub HighlightMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles HighlightMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(HighlightColorEvent))
    End Sub

    Private Sub BoldMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles BoldMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(BoldEvent))
    End Sub

    Private Sub ItalicMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles ItalicMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(ItalicEvent))
    End Sub

    Private Sub UnderlineMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles UnderlineMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(UnderlineEvent))
    End Sub

    Private Sub StrikethroughMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles StrikethroughMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(StrikethroughEvent))
    End Sub

    Private Sub SubscriptMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles SubscriptMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(SubscriptEvent))
    End Sub

    Private Sub SuperscriptMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles SuperscriptMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(SuperscriptEvent))
    End Sub

    Private Sub IndentMoreMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles IndentMoreMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(IndentMoreEvent))
    End Sub

    Private Sub IndentLessMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles IndentLessMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(IndentLessEvent))
    End Sub

    Private Sub BulletListMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles BulletListMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(BulletListEvent))
    End Sub

    Private Sub NumberListMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles NumberListMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(NumberListEvent))
    End Sub

    Private Sub AlignLeftMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles AlignLeftMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(AlignLeftEvent))
    End Sub

    Private Sub AlignCenterMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles AlignCenterMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(AlignCenterEvent))
    End Sub

    Private Sub AlignRightMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles AlignRightMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(AlignRightEvent))
    End Sub

    Private Sub AlignJustifyMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles AlignJustifyMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(AlignJustifyEvent))
    End Sub

    Private Sub LineSpacingMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles LineSpacingMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(LineSpacingEvent))
    End Sub

    Private Sub LefttoRightMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles LefttoRightMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(LeftToRightEvent))
    End Sub

    Private Sub RighttoLeftMenuItem_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles RighttoLeftMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(RightToLeftEvent))
    End Sub

    Private Sub UndoMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles UndoMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(UndoEvent))
    End Sub

    Private Sub RedoMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles RedoMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(RedoEvent))
    End Sub

    Private Sub CutMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles CutMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(CutEvent))
    End Sub

    Private Sub CopyMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles CopyMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(CopyEvent))
    End Sub

    Private Sub PasteMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles PasteMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(PasteEvent))
    End Sub

    Private Sub DeleteMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles DeleteMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(DeleteEvent))
    End Sub

    Private Sub SelectAllMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles SelectAllMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(SelectAllEvent))
    End Sub

    Private Sub FindMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles FindMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(FindEvent))
    End Sub

    Private Sub ReplaceMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles ReplaceMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(ReplaceEvent))
    End Sub

    Private Sub GoToMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles GoToMenuItem.Click
        Editor.RaiseEvent(New RoutedEventArgs(GoToEvent))
    End Sub

    Private Sub UppercaseMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles UppercaseMenuItem.Click
        Editor.Selection.Text = Editor.Selection.Text.ToUpper
    End Sub

    Private Sub LowercaseMenuItem_Click(sender As Object, e As RoutedEventArgs) Handles LowercaseMenuItem.Click
        Editor.Selection.Text = Editor.Selection.Text.ToLower
    End Sub

#End Region

    Private Sub TBox_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Editor.SelectionChanged
        Editor.RaiseEvent(New RoutedEventArgs(UpdateSelected))
    End Sub

#End Region

    Private Sub Editor_TextChanged(sender As Object, e As TextChangedEventArgs) Handles Editor.TextChanged
        SetDocumentTitle(DocName)
    End Sub
End Class