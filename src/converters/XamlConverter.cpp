#include "XamlConverter.h"

#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QTextList>
#include <QTextFrame>
#include <QTextFrameFormat>
#include <QTextImageFormat>
#include <QRegularExpression>
#include <QBuffer>
#include <QDir>
#include <QFileInfo>
#include <QImage>
#include <QUrl>

bool XamlConverter::loadFromXaml(const QString &xml, QTextDocument *doc,
                                  QMarginsF &outMargins, QColor &outPageBackground)
{
    QXmlStreamReader reader(xml);

    doc->clear();
    QTextCursor cursor(doc);
    bool firstParagraph = true;

    while (!reader.atEnd() && !reader.hasError()) {
        reader.readNext();
        if (reader.isStartElement()) {
            QString name = reader.name().toString();

            if (name == QStringLiteral("Document")) {
                QXmlStreamAttributes attrs = reader.attributes();
                if (attrs.hasAttribute(QStringLiteral("Width")) &&
                    attrs.hasAttribute(QStringLiteral("Height"))) {
                    qreal w = attrs.value(QStringLiteral("Width")).toDouble();
                    qreal h = attrs.value(QStringLiteral("Height")).toDouble();
                    doc->setPageSize(QSizeF(w, h));
                }
                if (attrs.hasAttribute(QStringLiteral("Padding"))) {
                    QStringList pads = attrs.value(QStringLiteral("Padding"))
                                          .toString().split(QStringLiteral(","));
                    if (pads.size() == 4) {
                        outMargins = QMarginsF(pads[0].toDouble(),
                                               pads[1].toDouble(),
                                               pads[2].toDouble(),
                                               pads[3].toDouble());
                    }
                }
                if (attrs.hasAttribute(QStringLiteral("Background"))) {
                    QColor c(attrs.value(QStringLiteral("Background")).toString());
                    if (c.isValid())
                        outPageBackground = c;
                }

            } else if (name == QStringLiteral("Paragraph")) {
                QTextBlockFormat blockFmt;
                QXmlStreamAttributes attrs = reader.attributes();
                if (attrs.hasAttribute(QStringLiteral("Alignment"))) {
                    QString align = attrs.value(QStringLiteral("Alignment")).toString();
                    if (align == QStringLiteral("Center"))
                        blockFmt.setAlignment(Qt::AlignCenter);
                    else if (align == QStringLiteral("Right"))
                        blockFmt.setAlignment(Qt::AlignRight);
                    else if (align == QStringLiteral("Justify"))
                        blockFmt.setAlignment(Qt::AlignJustify);
                    else
                        blockFmt.setAlignment(Qt::AlignLeft);
                }
                if (attrs.hasAttribute(QStringLiteral("LineHeight"))) {
                    qreal lh = attrs.value(QStringLiteral("LineHeight")).toDouble();
                    blockFmt.setLineHeight(static_cast<int>(lh * 100),
                                           QTextBlockFormat::LineDistanceHeight);
                }
                if (attrs.hasAttribute(QStringLiteral("LeftMargin")))
                    blockFmt.setLeftMargin(attrs.value(QStringLiteral("LeftMargin")).toDouble());
                if (attrs.hasAttribute(QStringLiteral("Indent")))
                    blockFmt.setIndent(static_cast<int>(attrs.value(QStringLiteral("Indent")).toDouble()));

                if (firstParagraph) {
                    cursor.setBlockFormat(blockFmt);
                    firstParagraph = false;
                } else {
                    cursor.insertBlock(blockFmt);
                }

            } else if (name == QStringLiteral("Run")) {
                QTextCharFormat charFmt;
                QXmlStreamAttributes attrs = reader.attributes();
                if (attrs.hasAttribute(QStringLiteral("Bold")) &&
                    attrs.value(QStringLiteral("Bold")) == QStringLiteral("True"))
                    charFmt.setFontWeight(QFont::Bold);
                if (attrs.hasAttribute(QStringLiteral("Italic")) &&
                    attrs.value(QStringLiteral("Italic")) == QStringLiteral("True"))
                    charFmt.setFontItalic(true);
                if (attrs.hasAttribute(QStringLiteral("Underline")) &&
                    attrs.value(QStringLiteral("Underline")) == QStringLiteral("True"))
                    charFmt.setFontUnderline(true);
                if (attrs.hasAttribute(QStringLiteral("Strikethrough")) &&
                    attrs.value(QStringLiteral("Strikethrough")) == QStringLiteral("True"))
                    charFmt.setFontStrikeOut(true);
                if (attrs.hasAttribute(QStringLiteral("Subscript")) &&
                    attrs.value(QStringLiteral("Subscript")) == QStringLiteral("True"))
                    charFmt.setVerticalAlignment(QTextCharFormat::AlignSubScript);
                if (attrs.hasAttribute(QStringLiteral("Superscript")) &&
                    attrs.value(QStringLiteral("Superscript")) == QStringLiteral("True"))
                    charFmt.setVerticalAlignment(QTextCharFormat::AlignSuperScript);
                if (attrs.hasAttribute(QStringLiteral("FontFamily"))) {
                    QFont f(attrs.value(QStringLiteral("FontFamily")).toString());
                    charFmt.setFont(f);
                }
                if (attrs.hasAttribute(QStringLiteral("FontSize")))
                    charFmt.setFontPointSize(attrs.value(QStringLiteral("FontSize")).toDouble());
                if (attrs.hasAttribute(QStringLiteral("Foreground"))) {
                    QColor color(attrs.value(QStringLiteral("Foreground")).toString());
                    if (color.isValid())
                        charFmt.setForeground(color);
                }
                if (attrs.hasAttribute(QStringLiteral("Background"))) {
                    QColor color(attrs.value(QStringLiteral("Background")).toString());
                    if (color.isValid())
                        charFmt.setBackground(color);
                }

                QString text = reader.readElementText();
                cursor.insertText(text, charFmt);

            } else if (name == QStringLiteral("LineBreak")) {
                cursor.insertText(QStringLiteral("\n"));

            } else if (name == QStringLiteral("List")) {
                QXmlStreamAttributes attrs = reader.attributes();
                QString style = attrs.value(QStringLiteral("Style")).toString();
                QTextListFormat listFmt;
                if (style == QStringLiteral("Decimal"))
                    listFmt.setStyle(QTextListFormat::ListDecimal);
                else
                    listFmt.setStyle(QTextListFormat::ListDisc);
                listFmt.setIndent(1);

                // Read list items
                while (!reader.atEnd() && !reader.hasError()) {
                    reader.readNext();
                    if (reader.isStartElement()) {
                        if (reader.name().toString() == QStringLiteral("ListItem")) {
                            QTextBlockFormat blockFmt;
                            if (firstParagraph) {
                                cursor.setBlockFormat(blockFmt);
                                firstParagraph = false;
                            } else {
                                cursor.insertBlock(blockFmt);
                            }
                            QTextList *list = cursor.createList(listFmt);

                            // Read content inside list item
                            while (!reader.atEnd() && !reader.hasError()) {
                                reader.readNext();
                                if (reader.isStartElement()) {
                                    if (reader.name().toString() == QStringLiteral("Run")) {
                                        QTextCharFormat cf;
                                        QString t = reader.readElementText();
                                        cursor.insertText(t, cf);
                                    } else if (reader.name().toString() == QStringLiteral("LineBreak")) {
                                        cursor.insertText(QStringLiteral("\n"));
                                    }
                                } else if (reader.isEndElement()) {
                                    if (reader.name().toString() == QStringLiteral("ListItem"))
                                        break;
                                }
                            }
                        }
                    } else if (reader.isEndElement()) {
                        if (reader.name().toString() == QStringLiteral("List"))
                            break;
                    }
                }

            } else if (name == QStringLiteral("Image")) {
                QXmlStreamAttributes attrs = reader.attributes();
                QString source = attrs.value(QStringLiteral("Source")).toString();
                qreal width = attrs.value(QStringLiteral("Width")).toDouble();
                qreal height = attrs.value(QStringLiteral("Height")).toDouble();

                QImage img;
                if (source.startsWith(QStringLiteral("base64:"), Qt::CaseInsensitive)) {
                    QByteArray data = QByteArray::fromBase64(
                        source.mid(7).toUtf8());
                    img.loadFromData(data);
                } else {
                    QFileInfo fi(source);
                    QString absPath = fi.absoluteFilePath();
                    QString cleanPath = QDir::cleanPath(absPath);
                    // Reject if path normalization changes the path (detects traversal like /../)
                    if (!fi.isAbsolute() || absPath != cleanPath) {
                        continue;
                    }
                    // Verify canonical path matches when file exists (detects symlink escape)
                    QString canonical = fi.canonicalFilePath();
                    if (!canonical.isEmpty() && canonical != cleanPath) {
                        continue;
                    }
                    img.load(cleanPath);
                }

                if (!img.isNull()) {
                    static quint64 s_embeddedImageSeq = 0;
                    QString name = QStringLiteral("xaml_embed_%1").arg(++s_embeddedImageSeq);
                    doc->addResource(QTextDocument::ImageResource,
                        QUrl(name), img);
                    QTextImageFormat imgFmt;
                    imgFmt.setName(name);
                    if (width > 0) imgFmt.setWidth(width);
                    if (height > 0) imgFmt.setHeight(height);
                    cursor.insertImage(imgFmt);
                }
            }
        }
    }

    if (reader.hasError()) {
        doc->clear();
        return false;
    }

    return true;
}

static void writeTable(QXmlStreamWriter &writer, QTextTable *table)
{
    writer.writeStartElement(QStringLiteral("Table"));
    writer.writeAttribute(QStringLiteral("Rows"),
        QString::number(table->rows()));
    writer.writeAttribute(QStringLiteral("Cols"),
        QString::number(table->columns()));

    for (int row = 0; row < table->rows(); ++row) {
        writer.writeStartElement(QStringLiteral("TableRow"));
        for (int col = 0; col < table->columns(); ++col) {
            QTextTableCell cell = table->cellAt(row, col);
            if (cell.row() != row || cell.column() != col)
                continue; // skip merged cells already handled

            writer.writeStartElement(QStringLiteral("TableCell"));
            QTextCursor cellCursor = cell.firstCursorPosition();
            QTextBlock block = cellCursor.block();
            while (block.isValid() && block.blockFormat() != cell.lastCursorPosition().block().blockFormat()) {
                QTextBlockFormat bf = block.blockFormat();
                writer.writeStartElement(QStringLiteral("Paragraph"));
                if (bf.alignment() == Qt::AlignCenter)
                    writer.writeAttribute(QStringLiteral("Alignment"), QStringLiteral("Center"));
                else if (bf.alignment() == Qt::AlignRight)
                    writer.writeAttribute(QStringLiteral("Alignment"), QStringLiteral("Right"));

                for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                    QTextFragment fragment = it.fragment();
                    if (!fragment.isValid()) continue;
                    QTextCharFormat cf = fragment.charFormat();
                    QString text = fragment.text();

                    if (text == QStringLiteral("\n")) {
                        writer.writeEmptyElement(QStringLiteral("LineBreak"));
                        continue;
                    }

                    writer.writeStartElement(QStringLiteral("Run"));
                    if (cf.fontWeight() == QFont::Bold)
                        writer.writeAttribute(QStringLiteral("Bold"), QStringLiteral("True"));
                    if (cf.fontItalic())
                        writer.writeAttribute(QStringLiteral("Italic"), QStringLiteral("True"));
                    if (cf.fontUnderline())
                        writer.writeAttribute(QStringLiteral("Underline"), QStringLiteral("True"));
                    writer.writeCharacters(text);
                    writer.writeEndElement();
                }
                writer.writeEndElement();
                block = block.next();
            }
            writer.writeEndElement();
        }
        writer.writeEndElement();
    }
    writer.writeEndElement();
}

QString XamlConverter::saveToXaml(const QTextDocument *doc,
                                   const QMarginsF &margins,
                                   const QColor &pageBackground)
{
    QString xml;
    QXmlStreamWriter writer(&xml);
    writer.setAutoFormatting(true);

    writer.writeStartDocument();
    writer.writeStartElement(QStringLiteral("Document"));
    writer.writeAttribute(QStringLiteral("Width"),
                          QString::number(doc->pageSize().width()));
    writer.writeAttribute(QStringLiteral("Height"),
                          QString::number(doc->pageSize().height()));
    writer.writeAttribute(QStringLiteral("Padding"),
                          QStringLiteral("%1,%2,%3,%4")
                              .arg(margins.left())
                              .arg(margins.top())
                              .arg(margins.right())
                              .arg(margins.bottom()));
    if (pageBackground.isValid())
        writer.writeAttribute(QStringLiteral("Background"), pageBackground.name());

    QTextBlock block = doc->begin();
    while (block.isValid()) {
        QTextBlockFormat blockFmt = block.blockFormat();

        // Check if this block is part of a list
        QTextList *list = block.textList();
        if (list) {
            // Only write the list once (at first block)
            if (list->itemNumber(block) == 0) {
                QTextListFormat listFmt = list->format();
                writer.writeStartElement(QStringLiteral("List"));
                if (listFmt.style() == QTextListFormat::ListDecimal)
                    writer.writeAttribute(QStringLiteral("Style"), QStringLiteral("Decimal"));
                else
                    writer.writeAttribute(QStringLiteral("Style"), QStringLiteral("Bullet"));

                // Write all list items
                for (int i = 0; i < list->count(); ++i) {
                    QTextBlock itemBlock = list->item(i);
                    writer.writeStartElement(QStringLiteral("ListItem"));

                    writer.writeStartElement(QStringLiteral("Paragraph"));
                    QString align;
                    if (blockFmt.alignment() == Qt::AlignCenter)
                        align = QStringLiteral("Center");
                    else if (blockFmt.alignment() == Qt::AlignRight)
                        align = QStringLiteral("Right");
                    else if (blockFmt.alignment() == Qt::AlignJustify)
                        align = QStringLiteral("Justify");
                    if (!align.isEmpty())
                        writer.writeAttribute(QStringLiteral("Alignment"), align);

                    for (QTextBlock::iterator it = itemBlock.begin(); !it.atEnd(); ++it) {
                        QTextFragment fragment = it.fragment();
                        if (!fragment.isValid()) continue;
                        QTextCharFormat charFmt = fragment.charFormat();
                        QString text = fragment.text();

                        if (text == QStringLiteral("\n")) {
                            writer.writeEmptyElement(QStringLiteral("LineBreak"));
                            continue;
                        }

                        writer.writeStartElement(QStringLiteral("Run"));
                        if (charFmt.fontWeight() == QFont::Bold)
                            writer.writeAttribute(QStringLiteral("Bold"), QStringLiteral("True"));
                        if (charFmt.fontItalic())
                            writer.writeAttribute(QStringLiteral("Italic"), QStringLiteral("True"));
                        if (charFmt.fontUnderline())
                            writer.writeAttribute(QStringLiteral("Underline"), QStringLiteral("True"));
                        if (charFmt.fontStrikeOut())
                            writer.writeAttribute(QStringLiteral("Strikethrough"), QStringLiteral("True"));
                        writer.writeCharacters(text);
                        writer.writeEndElement();
                    }
                    writer.writeEndElement(); // Paragraph
                    writer.writeEndElement(); // ListItem
                }
                writer.writeEndElement(); // List
            }
            block = block.next();
            continue;
        }

        // Check if this block is inside a table
        QTextFrame *blockFrame = doc->frameAt(block.position());
        QTextTable *table = nullptr;
        while (blockFrame) {
            table = qobject_cast<QTextTable *>(blockFrame);
            if (table) break;
            blockFrame = blockFrame->parentFrame();
        }
        if (table) {
            writeTable(writer, table);
            // Skip all blocks in this table
            int tableEnd = table->lastPosition();
            while (block.isValid() && block.position() <= tableEnd)
                block = block.next();
            continue;
        }

        // Regular paragraph
        writer.writeStartElement(QStringLiteral("Paragraph"));

        QString align;
        if (blockFmt.alignment() == Qt::AlignCenter)
            align = QStringLiteral("Center");
        else if (blockFmt.alignment() == Qt::AlignRight)
            align = QStringLiteral("Right");
        else if (blockFmt.alignment() == Qt::AlignJustify)
            align = QStringLiteral("Justify");
        else
            align = QStringLiteral("Left");
        writer.writeAttribute(QStringLiteral("Alignment"), align);

        qreal lh = blockFmt.lineHeight();
        if (lh > 0)
            writer.writeAttribute(QStringLiteral("LineHeight"), QString::number(lh / 100.0));

        if (blockFmt.leftMargin() > 0)
            writer.writeAttribute(QStringLiteral("LeftMargin"),
                                  QString::number(blockFmt.leftMargin()));
        if (blockFmt.indent() > 0)
            writer.writeAttribute(QStringLiteral("Indent"),
                                  QString::number(blockFmt.indent()));

        for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
            QTextFragment fragment = it.fragment();
            if (!fragment.isValid())
                continue;

            QTextCharFormat charFmt = fragment.charFormat();
            QString text = fragment.text();

            // Check for image
            if (charFmt.isImageFormat()) {
                QTextImageFormat imgFmt = charFmt.toImageFormat();
                QVariant resource = doc->resource(QTextDocument::ImageResource,
                    QUrl(imgFmt.name()));
                if (resource.isValid()) {
                    QImage img = resource.value<QImage>();
                    if (!img.isNull()) {
                        QByteArray ba;
                        QBuffer buf(&ba);
                        buf.open(QIODevice::WriteOnly);
                        img.save(&buf, "PNG");
                        buf.close();

                        writer.writeEmptyElement(QStringLiteral("Image"));
                        writer.writeAttribute(QStringLiteral("Source"),
                            QStringLiteral("base64:") + QString::fromLatin1(ba.toBase64()));
                        writer.writeAttribute(QStringLiteral("Width"),
                            QString::number(imgFmt.width() > 0 ? imgFmt.width() : img.width()));
                        writer.writeAttribute(QStringLiteral("Height"),
                            QString::number(imgFmt.height() > 0 ? imgFmt.height() : img.height()));
                        continue;
                    }
                }
                // Fallback: write as runs with path reference
            }

            if (text == QStringLiteral("\n")) {
                writer.writeEmptyElement(QStringLiteral("LineBreak"));
                continue;
            }

            writer.writeStartElement(QStringLiteral("Run"));

            if (charFmt.fontWeight() == QFont::Bold)
                writer.writeAttribute(QStringLiteral("Bold"), QStringLiteral("True"));
            if (charFmt.fontItalic())
                writer.writeAttribute(QStringLiteral("Italic"), QStringLiteral("True"));
            if (charFmt.fontUnderline())
                writer.writeAttribute(QStringLiteral("Underline"), QStringLiteral("True"));
            if (charFmt.fontStrikeOut())
                writer.writeAttribute(QStringLiteral("Strikethrough"), QStringLiteral("True"));
            if (charFmt.verticalAlignment() == QTextCharFormat::AlignSubScript)
                writer.writeAttribute(QStringLiteral("Subscript"), QStringLiteral("True"));
            if (charFmt.verticalAlignment() == QTextCharFormat::AlignSuperScript)
                writer.writeAttribute(QStringLiteral("Superscript"), QStringLiteral("True"));

            writer.writeAttribute(QStringLiteral("FontFamily"),
                                  charFmt.font().family());

            if (charFmt.fontPointSize() > 0)
                writer.writeAttribute(QStringLiteral("FontSize"),
                                      QString::number(charFmt.fontPointSize()));

            QColor fgColor = charFmt.foreground().color();
            if (fgColor.isValid() && fgColor != QColor(Qt::black))
                writer.writeAttribute(QStringLiteral("Foreground"), fgColor.name());

            QColor bgColor = charFmt.background().color();
            if (bgColor.isValid() && bgColor != QColor(Qt::transparent))
                writer.writeAttribute(QStringLiteral("Background"), bgColor.name());

            writer.writeCharacters(text);
            writer.writeEndElement();
        }

        writer.writeEndElement();
        block = block.next();
    }

    writer.writeEndElement();
    writer.writeEndDocument();

    return xml;
}
