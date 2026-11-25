#include "DocxConverter.h"
#include "DocxCommon.h"

#include "miniz.h"

#include <QXmlStreamWriter>
#include <QBuffer>
#include <QImage>
#include <QUrl>
#include <QColor>
#include <QFileInfo>
#include <QTextDocument>
#include <QTextCharFormat>
#include <QTextList>
#include <QTextTable>
#include <QTextImageFormat>
#include <QMap>
#include <QVector>
#include <QPair>
#include <QSet>
#include <QDebug>

// Streams zip output into a QByteArray (the write target for the DOCX writer).
static size_t zipAppendCallback(void *pOpaque, mz_uint64 fileOfs, const void *pBuf, size_t n)
{
    QByteArray *buffer = static_cast<QByteArray *>(pOpaque);
    if (fileOfs > static_cast<mz_uint64>(buffer->size()))
        return 0;
    if (fileOfs < static_cast<mz_uint64>(buffer->size()))
        buffer->truncate(static_cast<int>(fileOfs));
    buffer->append(static_cast<const char *>(pBuf), static_cast<int>(n));
    return n;
}

static bool addZipFile(mz_zip_archive *zip, const QString &name, const QByteArray &data)
{
    return mz_zip_writer_add_mem(zip, name.toUtf8().constData(),
                                 data.constData(), data.size(), MZ_BEST_COMPRESSION);
}

// Writes a run's text, splitting on soft line separators (U+2028) so that each
// line break becomes a w:br element instead of a literal control character.
static void writeTextWithBreaks(QXmlStreamWriter &w, const QString &text)
{
    const QStringList parts = text.split(QChar::LineSeparator);
    for (int i = 0; i < parts.size(); ++i) {
        if (!parts.at(i).isEmpty())
            w.writeTextElement(QLatin1String("w:t"), parts.at(i));
        if (i + 1 < parts.size()) {
            w.writeStartElement(QLatin1String("w:r"));
            w.writeEmptyElement(QLatin1String("w:br"));
            w.writeEndElement();
        }
    }
}

static QString imageFormatFromName(const QString &name)
{
    QString ext = QFileInfo(name).suffix().toLower();
    if (ext == QLatin1String("jpg") || ext == QLatin1String("jpeg"))
        return QStringLiteral("JPEG");
    if (ext == QLatin1String("gif"))
        return QStringLiteral("GIF");
    if (ext == QLatin1String("bmp"))
        return QStringLiteral("BMP");
    return QStringLiteral("PNG");
}


static void writeTextRunXml(QXmlStreamWriter &w, const QTextCharFormat &cf,
                            const QString &text,
                            QMap<QString, QString> &hyperlinkUrlToRelId,
                            QMap<QString, QString> &hyperlinkRelIdToUrl,
                            int &relationCounter)
{
    bool isLink = cf.isAnchor() && !cf.anchorHref().isEmpty();
    if (isLink) {
        QString href = cf.anchorHref();
        if (!hyperlinkUrlToRelId.contains(href)) {
            ++relationCounter;
            QString rid = QStringLiteral("rId%1").arg(relationCounter);
            hyperlinkUrlToRelId[href] = rid;
            hyperlinkRelIdToUrl[rid] = href;
        }
        w.writeStartElement(QLatin1String("w:hyperlink"));
        w.writeAttribute(QLatin1String("r:id"), hyperlinkUrlToRelId[href]);
    }
    w.writeStartElement(QLatin1String("w:r"));
    w.writeStartElement(QLatin1String("w:rPr"));

    if (cf.fontWeight() == QFont::Bold)
        w.writeEmptyElement(QLatin1String("w:b"));
    if (cf.fontItalic())
        w.writeEmptyElement(QLatin1String("w:i"));
    if (cf.fontUnderline()) {
        w.writeStartElement(QLatin1String("w:u"));
        w.writeAttribute(QLatin1String("w:val"), QLatin1String("single"));
        w.writeEndElement();
    }
    if (cf.fontStrikeOut())
        w.writeEmptyElement(QLatin1String("w:strike"));
    if (cf.verticalAlignment() == QTextCharFormat::AlignSubScript) {
        w.writeStartElement(QLatin1String("w:vertAlign"));
        w.writeAttribute(QLatin1String("w:val"), QLatin1String("subscript"));
        w.writeEndElement();
    } else if (cf.verticalAlignment() == QTextCharFormat::AlignSuperScript) {
        w.writeStartElement(QLatin1String("w:vertAlign"));
        w.writeAttribute(QLatin1String("w:val"), QLatin1String("superscript"));
        w.writeEndElement();
    }

    QStringList families = cf.fontFamilies().toStringList();
    if (!families.isEmpty()) {
        QString fam = families.first();
        if (!fam.isEmpty() && fam != QLatin1String("Segoe UI")) {
            w.writeStartElement(QLatin1String("w:rFonts"));
            w.writeAttribute(QLatin1String("w:ascii"), fam);
            w.writeAttribute(QLatin1String("w:hAnsi"), fam);
            if (families.size() > 1)
                w.writeAttribute(QLatin1String("w:eastAsia"), families[1]);
            if (families.size() > 2)
                w.writeAttribute(QLatin1String("w:cs"), families[2]);
            w.writeEndElement();
        }
    }

    if (cf.fontPointSize() > 0) {
        int half = ptToHalfPt(cf.fontPointSize());
        w.writeStartElement(QLatin1String("w:sz"));
        w.writeAttribute(QLatin1String("w:val"), QString::number(half));
        w.writeEndElement();
        w.writeStartElement(QLatin1String("w:szCs"));
        w.writeAttribute(QLatin1String("w:val"), QString::number(half));
        w.writeEndElement();
    }

    QColor fg = cf.foreground().color();
    if (fg.isValid() && fg != QColor(Qt::black)) {
        w.writeStartElement(QLatin1String("w:color"));
        w.writeAttribute(QLatin1String("w:val"), fg.name().mid(1));
        w.writeEndElement();
    }

    w.writeEndElement();
    writeTextWithBreaks(w, text);
    w.writeEndElement();
    if (isLink)
        w.writeEndElement();
}

static void writeTableXml(QXmlStreamWriter &w, QTextTable *table,
                          QTextDocument *doc,
                          int &imageCounter, int &relationCounter,
                          QStringList &imageRels, QStringList &imageTargets,
                          mz_zip_archive &zip,
                          QSet<QString> &usedStyles,
                          QMap<QString, QString> &hyperlinkUrlToRelId,
                          QMap<QString, QString> &hyperlinkRelIdToUrl)
{
    w.writeStartElement(QLatin1String("w:tbl"));
    w.writeStartElement(QLatin1String("w:tblPr"));
    w.writeStartElement(QLatin1String("w:tblW"));
    w.writeAttribute(QLatin1String("w:w"), QLatin1String("5000"));
    w.writeAttribute(QLatin1String("w:type"), QLatin1String("pct"));
    w.writeEndElement();
    w.writeEndElement();

    for (int row = 0; row < table->rows(); ++row) {
        w.writeStartElement(QLatin1String("w:tr"));
        for (int col = 0; col < table->columns();) {
            QTextTableCell cell = table->cellAt(row, col);
            const int startRow = cell.row();
            const int startCol = cell.column();
            const int cSpan = cell.columnSpan();
            const int rSpan = cell.rowSpan();

            if (startRow == row && startCol == col) {
                w.writeStartElement(QLatin1String("w:tc"));

                if (cSpan > 1 || rSpan > 1) {
                    w.writeStartElement(QLatin1String("w:tcPr"));
                    if (cSpan > 1) {
                        w.writeStartElement(QLatin1String("w:gridSpan"));
                        w.writeAttribute(QLatin1String("w:val"), QString::number(cSpan));
                        w.writeEndElement();
                    }
                    if (rSpan > 1) {
                        w.writeStartElement(QLatin1String("w:vMerge"));
                        w.writeAttribute(QLatin1String("w:val"), QLatin1String("restart"));
                        w.writeEndElement();
                    }
                    w.writeEndElement();
                }

            QTextCursor cellCursor = cell.firstCursorPosition();
            QTextBlock block = cellCursor.block();
            int cellEnd = cell.lastPosition();

            while (block.isValid() && block.position() <= cellEnd) {
                QTextBlockFormat bf = block.blockFormat();

                w.writeStartElement(QLatin1String("w:p"));
                w.writeStartElement(QLatin1String("w:pPr"));

                QString pStyle = bf.property(QTextFormat::UserProperty).toString();
                if (!pStyle.isEmpty()) {
                    w.writeStartElement(QLatin1String("w:pStyle"));
                    w.writeAttribute(QLatin1String("w:val"), pStyle);
                    w.writeEndElement();
                    usedStyles.insert(pStyle);
                }

                Qt::Alignment align = bf.alignment();
                if (align == Qt::AlignCenter) {
                    w.writeStartElement(QLatin1String("w:jc"));
                    w.writeAttribute(QLatin1String("w:val"), QLatin1String("center"));
                    w.writeEndElement();
                } else if (align == Qt::AlignRight) {
                    w.writeStartElement(QLatin1String("w:jc"));
                    w.writeAttribute(QLatin1String("w:val"), QLatin1String("right"));
                    w.writeEndElement();
                } else if (align == Qt::AlignJustify) {
                    w.writeStartElement(QLatin1String("w:jc"));
                    w.writeAttribute(QLatin1String("w:val"), QLatin1String("both"));
                    w.writeEndElement();
                }

                if (bf.leftMargin() > 0 || bf.textIndent() != 0) {
                    w.writeStartElement(QLatin1String("w:ind"));
                    if (bf.leftMargin() > 0)
                        w.writeAttribute(QLatin1String("w:left"), QString::number(pxToTwip(bf.leftMargin())));
                    if (bf.textIndent() != 0) {
                        if (bf.textIndent() > 0)
                            w.writeAttribute(QLatin1String("w:firstLine"), QString::number(pxToTwip(bf.textIndent())));
                        else
                            w.writeAttribute(QLatin1String("w:hanging"), QString::number(pxToTwip(-bf.textIndent())));
                    }
                    w.writeEndElement();
                }

                w.writeEndElement();

                for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                    QTextFragment fragment = it.fragment();
                    if (!fragment.isValid()) continue;

                    QTextCharFormat cf = fragment.charFormat();
                    QString text = fragment.text();

                    if (cf.isImageFormat()) {
                        QTextImageFormat imgFmt = cf.toImageFormat();
                        QVariant res = doc->resource(QTextDocument::ImageResource, QUrl(imgFmt.name()));
                        if (res.isValid()) {
                            QImage img = qvariant_cast<QImage>(res);
                            if (!img.isNull()) {
                                ++imageCounter;
                                ++relationCounter;
                                QString relId = QStringLiteral("rId%1").arg(relationCounter);
                                QString fmt = imageFormatFromName(imgFmt.name());
                                QString ext = fmt.toLower();
                                if (ext == QLatin1String("jpeg")) ext = QStringLiteral("jpg");
                                QString target = QStringLiteral("media/image%1.%2").arg(imageCounter).arg(ext);
                                imageRels.append(relId);
                                imageTargets.append(target);

                                QByteArray imgData;
                                QBuffer imgBuf(&imgData);
                                if (!imgBuf.open(QIODevice::WriteOnly) || !img.save(&imgBuf, fmt.toLatin1().constData())) {
                                    continue;
                                }
                                imgBuf.close();
                                if (imgData.isEmpty()) {
                                    continue;
                                }
                                if (!addZipFile(&zip, QLatin1String("word/") + target, imgData)) {
                                    continue;
                                }

                                w.writeStartElement(QLatin1String("w:r"));
                                w.writeStartElement(QLatin1String("w:rPr"));
                                w.writeEndElement();
                                w.writeStartElement(QLatin1String("w:drawing"));
                                w.writeNamespace(NS_WP, QLatin1String("wp"));
                                w.writeNamespace(NS_A, QLatin1String("a"));
                                w.writeStartElement(QLatin1String("wp:inline"));
                                w.writeStartElement(QLatin1String("wp:extent"));
                                qreal iw = imgFmt.width() > 0 ? imgFmt.width() : img.width();
                                qreal ih = imgFmt.height() > 0 ? imgFmt.height() : img.height();
                                w.writeAttribute(QLatin1String("cx"), QString::number(pxToEmu(iw)));
                                w.writeAttribute(QLatin1String("cy"), QString::number(pxToEmu(ih)));
                                w.writeEndElement();
                                w.writeStartElement(QLatin1String("wp:docPr"));
                                w.writeAttribute(QLatin1String("id"), QString::number(imageCounter));
                                w.writeAttribute(QLatin1String("name"), QStringLiteral("Image%1").arg(imageCounter));
                                w.writeEndElement();
                                w.writeStartElement(QLatin1String("a:graphic"));
                                w.writeStartElement(QLatin1String("a:graphicData"));
                                w.writeAttribute(QLatin1String("uri"), QLatin1String("http://schemas.openxmlformats.org/drawingml/2006/picture"));
                                w.writeStartElement(QLatin1String("pic:pic"));
                                w.writeNamespace(QLatin1String("http://schemas.openxmlformats.org/drawingml/2006/picture"), QLatin1String("pic"));
                                w.writeStartElement(QLatin1String("pic:blipFill"));
                                w.writeStartElement(QLatin1String("a:blip"));
                                w.writeAttribute(QLatin1String("r:embed"), relId);
                                w.writeEndElement();
                                w.writeEndElement();
                                w.writeStartElement(QLatin1String("pic:spPr"));
                                w.writeStartElement(QLatin1String("a:xfrm"));
                                w.writeStartElement(QLatin1String("a:off"));
                                w.writeAttribute(QLatin1String("x"), QLatin1String("0"));
                                w.writeAttribute(QLatin1String("y"), QLatin1String("0"));
                                w.writeEndElement();
                                w.writeStartElement(QLatin1String("a:ext"));
                                w.writeAttribute(QLatin1String("cx"), QString::number(pxToEmu(iw)));
                                w.writeAttribute(QLatin1String("cy"), QString::number(pxToEmu(ih)));
                                w.writeEndElement();
                                w.writeEndElement();
                                w.writeStartElement(QLatin1String("a:prstGeom"));
                                w.writeAttribute(QLatin1String("prst"), QLatin1String("rect"));
                                w.writeEndElement();
                                w.writeEndElement();
                                w.writeEndElement();
                                w.writeEndElement();
                                w.writeEndElement();
                                w.writeEndElement();
                                w.writeEndElement();
                                w.writeEndElement();
                            }
                        }
                        continue;
                    }

                    writeTextRunXml(w, cf, text, hyperlinkUrlToRelId,
                                    hyperlinkRelIdToUrl, relationCounter);
                }

                w.writeEndElement();
                block = block.next();
            }

                w.writeEndElement();

                col += cSpan;
            } else {
                w.writeStartElement(QLatin1String("w:tc"));
                w.writeStartElement(QLatin1String("w:tcPr"));
                if (cSpan > 1) {
                    w.writeStartElement(QLatin1String("w:gridSpan"));
                    w.writeAttribute(QLatin1String("w:val"), QString::number(cSpan));
                    w.writeEndElement();
                }
                w.writeEmptyElement(QLatin1String("w:vMerge"));
                w.writeEndElement();
                w.writeEmptyElement(QLatin1String("w:p"));
                w.writeEndElement();
                col += cSpan;
            }
        }
        w.writeEndElement();
    }

    w.writeEndElement();
}

QByteArray DocxConverter::saveToDocx(const QTextDocument *doc,
                                      const QMarginsF &margins,
                                      const QColor &pageBackground)
{
    QByteArray zipBuffer;
    mz_zip_archive zip;
    mz_zip_zero_struct(&zip);
    zip.m_pWrite = zipAppendCallback;
    zip.m_pIO_opaque = &zipBuffer;
    if (!mz_zip_writer_init_v2(&zip, 0, 0))
        return {};

    {
        QString xml;
        QXmlStreamWriter w(&xml);
        w.setAutoFormatting(true);
        w.writeStartDocument();
        w.writeStartElement(QLatin1String("Types"));
        w.writeDefaultNamespace(NS_CT);
        w.writeStartElement(QLatin1String("Default"));
        w.writeAttribute(QLatin1String("Extension"), QLatin1String("rels"));
        w.writeAttribute(QLatin1String("ContentType"), QLatin1String("application/vnd.openxmlformats-package.relationships+xml"));
        w.writeEndElement();
        w.writeStartElement(QLatin1String("Default"));
        w.writeAttribute(QLatin1String("Extension"), QLatin1String("xml"));
        w.writeAttribute(QLatin1String("ContentType"), QLatin1String("application/xml"));
        w.writeEndElement();
        w.writeStartElement(QLatin1String("Default"));
        w.writeAttribute(QLatin1String("Extension"), QLatin1String("png"));
        w.writeAttribute(QLatin1String("ContentType"), QLatin1String("image/png"));
        w.writeEndElement();
        w.writeStartElement(QLatin1String("Default"));
        w.writeAttribute(QLatin1String("Extension"), QLatin1String("jpg"));
        w.writeAttribute(QLatin1String("ContentType"), QLatin1String("image/jpeg"));
        w.writeEndElement();
        w.writeStartElement(QLatin1String("Default"));
        w.writeAttribute(QLatin1String("Extension"), QLatin1String("gif"));
        w.writeAttribute(QLatin1String("ContentType"), QLatin1String("image/gif"));
        w.writeEndElement();
        w.writeStartElement(QLatin1String("Default"));
        w.writeAttribute(QLatin1String("Extension"), QLatin1String("bmp"));
        w.writeAttribute(QLatin1String("ContentType"), QLatin1String("image/bmp"));
        w.writeEndElement();
        w.writeStartElement(QLatin1String("Override"));
        w.writeAttribute(QLatin1String("PartName"), QLatin1String("/word/document.xml"));
        w.writeAttribute(QLatin1String("ContentType"), QLatin1String("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"));
        w.writeEndElement();
        w.writeStartElement(QLatin1String("Override"));
        w.writeAttribute(QLatin1String("PartName"), QLatin1String("/word/styles.xml"));
        w.writeAttribute(QLatin1String("ContentType"), QLatin1String("application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"));
        w.writeEndElement();
        w.writeStartElement(QLatin1String("Override"));
        w.writeAttribute(QLatin1String("PartName"), QLatin1String("/word/numbering.xml"));
        w.writeAttribute(QLatin1String("ContentType"), QLatin1String("application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"));
        w.writeEndElement();
        w.writeEndElement();
        w.writeEndDocument();
        if (!addZipFile(&zip, QLatin1String("[Content_Types].xml"), xml.toUtf8())) {
            mz_zip_writer_end(&zip);
            return {};
        }
    }

    {
        QString xml;
        QXmlStreamWriter w(&xml);
        w.setAutoFormatting(true);
        w.writeStartDocument();
        w.writeStartElement(QLatin1String("Relationships"));
        w.writeDefaultNamespace(NS_REL);
        w.writeStartElement(QLatin1String("Relationship"));
        w.writeAttribute(QLatin1String("Id"), QLatin1String("rId1"));
        w.writeAttribute(QLatin1String("Type"), QLatin1String("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"));
        w.writeAttribute(QLatin1String("Target"), QLatin1String("word/document.xml"));
        w.writeEndElement();
        w.writeEndElement();
        w.writeEndDocument();
        if (!addZipFile(&zip, QLatin1String("_rels/.rels"), xml.toUtf8())) {
            mz_zip_writer_end(&zip);
            return {};
        }
    }

    int imageCounter = 0;
    int relationCounter = 1;
    QStringList imageRels;
    QStringList imageTargets;
    QMap<QString, QString> hyperlinkUrlToRelId;
    QMap<QString, QString> hyperlinkRelIdToUrl;
    QSet<QString> usedStyles;

    {
        QString xml;
        QXmlStreamWriter w(&xml);
        w.setAutoFormatting(true);
        w.writeStartDocument();
        w.writeStartElement(QLatin1String("w:document"));
        w.writeNamespace(NS_W, QLatin1String("w"));
        w.writeNamespace(NS_R, QLatin1String("r"));
        w.writeStartElement(QLatin1String("w:body"));

        QTextBlock block = doc->begin();
        QSet<quintptr> writtenTables;
        // Deterministic numbering-id assignment: ids are handed out in order
        // of first appearance, so document.xml and numbering.xml agree.
        QMap<quintptr, int> numIdByList;

        while (block.isValid()) {
            QTextFrame *blockFrame = doc->frameAt(block.position());
            QTextTable *table = nullptr;
            while (blockFrame) {
                table = qobject_cast<QTextTable *>(blockFrame);
                if (table) break;
                blockFrame = blockFrame->parentFrame();
            }

            if (table) {
                quintptr tp = reinterpret_cast<quintptr>(table);
                if (!writtenTables.contains(tp)) {
                    writeTableXml(w, table, const_cast<QTextDocument*>(doc),
                                  imageCounter, relationCounter,
                                  imageRels, imageTargets, zip,
                                  usedStyles,
                                  hyperlinkUrlToRelId, hyperlinkRelIdToUrl);
                    writtenTables.insert(tp);
                }
                int tableEnd = table->lastPosition();
                while (block.isValid() && block.position() <= tableEnd)
                    block = block.next();
                continue;
            }

            QTextBlockFormat bf = block.blockFormat();
            QTextList *list = block.textList();

            w.writeStartElement(QLatin1String("w:p"));
            w.writeStartElement(QLatin1String("w:pPr"));

            QString pStyle = bf.property(QTextFormat::UserProperty).toString();
            if (!pStyle.isEmpty()) {
                w.writeStartElement(QLatin1String("w:pStyle"));
                w.writeAttribute(QLatin1String("w:val"), pStyle);
                w.writeEndElement();
                usedStyles.insert(pStyle);
            }

            Qt::Alignment align = bf.alignment();
            if (align == Qt::AlignCenter) {
                w.writeStartElement(QLatin1String("w:jc"));
                w.writeAttribute(QLatin1String("w:val"), QLatin1String("center"));
                w.writeEndElement();
            } else if (align == Qt::AlignRight) {
                w.writeStartElement(QLatin1String("w:jc"));
                w.writeAttribute(QLatin1String("w:val"), QLatin1String("right"));
                w.writeEndElement();
            } else if (align == Qt::AlignJustify) {
                w.writeStartElement(QLatin1String("w:jc"));
                w.writeAttribute(QLatin1String("w:val"), QLatin1String("both"));
                w.writeEndElement();
            }

            if (bf.lineHeight() > 0 && bf.lineHeight() != 100) {
                w.writeStartElement(QLatin1String("w:spacing"));
                w.writeAttribute(QLatin1String("w:line"), QString::number(propToDocxLine(bf.lineHeight())));
                w.writeAttribute(QLatin1String("w:lineRule"), QLatin1String("auto"));
                w.writeEndElement();
            }

            if (bf.leftMargin() > 0 || bf.textIndent() != 0) {
                w.writeStartElement(QLatin1String("w:ind"));
                if (bf.leftMargin() > 0)
                    w.writeAttribute(QLatin1String("w:left"), QString::number(pxToTwip(bf.leftMargin())));
                if (bf.textIndent() != 0) {
                    if (bf.textIndent() > 0)
                        w.writeAttribute(QLatin1String("w:firstLine"), QString::number(pxToTwip(bf.textIndent())));
                    else
                        w.writeAttribute(QLatin1String("w:hanging"), QString::number(pxToTwip(-bf.textIndent())));
                }
                w.writeEndElement();
            }

            if (list) {
                QTextListFormat lf = list->format();
                const quintptr key = reinterpret_cast<quintptr>(list);
                int numId = numIdByList.value(key);
                if (numId == 0) {
                    numId = numIdByList.size() + 1;
                    numIdByList.insert(key, numId);
                }
                int ilvl = 0;
                if (list->itemNumber(block) >= 0)
                    ilvl = qMax(0, lf.indent() - 1);
                w.writeStartElement(QLatin1String("w:numPr"));
                w.writeTextElement(QLatin1String("w:ilvl"), QString::number(ilvl));
                w.writeTextElement(QLatin1String("w:numId"), QString::number(numId));
                w.writeEndElement();
            }

            w.writeEndElement();

            for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                QTextFragment fragment = it.fragment();
                if (!fragment.isValid()) continue;

                QTextCharFormat cf = fragment.charFormat();
                QString text = fragment.text();

                if (cf.isImageFormat()) {
                    QTextImageFormat imgFmt = cf.toImageFormat();
                    QVariant res = doc->resource(QTextDocument::ImageResource, QUrl(imgFmt.name()));
                    if (res.isValid()) {
                        QImage img = qvariant_cast<QImage>(res);
                        if (!img.isNull()) {
                            ++imageCounter;
                            ++relationCounter;
                            QString relId = QStringLiteral("rId%1").arg(relationCounter);
                            QString fmt = imageFormatFromName(imgFmt.name());
                            QString ext = fmt.toLower();
                            if (ext == QLatin1String("jpeg")) ext = QStringLiteral("jpg");
                            QString target = QStringLiteral("media/image%1.%2").arg(imageCounter).arg(ext);
                            imageRels.append(relId);
                            imageTargets.append(target);

                            QByteArray imgData;
                            QBuffer imgBuf(&imgData);
                            if (!imgBuf.open(QIODevice::WriteOnly) || !img.save(&imgBuf, fmt.toLatin1().constData())) {
                                continue;
                            }
                            imgBuf.close();
                            if (imgData.isEmpty()) {
                                continue;
                            }
                            if (!addZipFile(&zip, QLatin1String("word/") + target, imgData)) {
                                continue;
                            }

                            w.writeStartElement(QLatin1String("w:r"));
                            w.writeStartElement(QLatin1String("w:rPr"));
                            w.writeEndElement();
                            w.writeStartElement(QLatin1String("w:drawing"));
                            w.writeNamespace(NS_WP, QLatin1String("wp"));
                            w.writeNamespace(NS_A, QLatin1String("a"));
                            w.writeStartElement(QLatin1String("wp:inline"));
                            w.writeStartElement(QLatin1String("wp:extent"));
                            qreal iw = imgFmt.width() > 0 ? imgFmt.width() : img.width();
                            qreal ih = imgFmt.height() > 0 ? imgFmt.height() : img.height();
                            w.writeAttribute(QLatin1String("cx"), QString::number(pxToEmu(iw)));
                            w.writeAttribute(QLatin1String("cy"), QString::number(pxToEmu(ih)));
                            w.writeEndElement();
                            w.writeStartElement(QLatin1String("wp:docPr"));
                            w.writeAttribute(QLatin1String("id"), QString::number(imageCounter));
                            w.writeAttribute(QLatin1String("name"), QStringLiteral("Image%1").arg(imageCounter));
                            w.writeEndElement();
                            w.writeStartElement(QLatin1String("a:graphic"));
                            w.writeStartElement(QLatin1String("a:graphicData"));
                            w.writeAttribute(QLatin1String("uri"), QLatin1String("http://schemas.openxmlformats.org/drawingml/2006/picture"));
                            w.writeStartElement(QLatin1String("pic:pic"));
                            w.writeNamespace(QLatin1String("http://schemas.openxmlformats.org/drawingml/2006/picture"), QLatin1String("pic"));
                            w.writeStartElement(QLatin1String("pic:blipFill"));
                            w.writeStartElement(QLatin1String("a:blip"));
                            w.writeAttribute(QLatin1String("r:embed"), relId);
                            w.writeEndElement();
                            w.writeEndElement();
                            w.writeStartElement(QLatin1String("pic:spPr"));
                            w.writeStartElement(QLatin1String("a:xfrm"));
                            w.writeStartElement(QLatin1String("a:off"));
                            w.writeAttribute(QLatin1String("x"), QLatin1String("0"));
                            w.writeAttribute(QLatin1String("y"), QLatin1String("0"));
                            w.writeEndElement();
                            w.writeStartElement(QLatin1String("a:ext"));
                            w.writeAttribute(QLatin1String("cx"), QString::number(pxToEmu(iw)));
                            w.writeAttribute(QLatin1String("cy"), QString::number(pxToEmu(ih)));
                            w.writeEndElement();
                            w.writeEndElement();
                            w.writeStartElement(QLatin1String("a:prstGeom"));
                            w.writeAttribute(QLatin1String("prst"), QLatin1String("rect"));
                            w.writeEndElement();
                            w.writeEndElement();
                            w.writeEndElement();
                            w.writeEndElement();
                            w.writeEndElement();
                            w.writeEndElement();
                            w.writeEndElement();
                            w.writeEndElement();
                        }
                    }
                    continue;
                }

                writeTextRunXml(w, cf, text, hyperlinkUrlToRelId,
                                hyperlinkRelIdToUrl, relationCounter);
            }

            w.writeEndElement();

            block = block.next();
        }

        w.writeStartElement(QLatin1String("w:sectPr"));
        w.writeStartElement(QLatin1String("w:pgSz"));
        w.writeAttribute(QLatin1String("w:w"), QString::number(pxToTwip(doc->pageSize().width())));
        w.writeAttribute(QLatin1String("w:h"), QString::number(pxToTwip(doc->pageSize().height())));
        w.writeEndElement();
        w.writeStartElement(QLatin1String("w:pgMar"));
        w.writeAttribute(QLatin1String("w:top"), QString::number(pxToTwip(margins.top())));
        w.writeAttribute(QLatin1String("w:right"), QString::number(pxToTwip(margins.right())));
        w.writeAttribute(QLatin1String("w:bottom"), QString::number(pxToTwip(margins.bottom())));
        w.writeAttribute(QLatin1String("w:left"), QString::number(pxToTwip(margins.left())));
        w.writeEndElement();
        if (pageBackground.isValid() && pageBackground != QColor(Qt::white)) {
            w.writeStartElement(QLatin1String("w:shd"));
            w.writeAttribute(QLatin1String("w:fill"), pageBackground.name().mid(1));
            w.writeAttribute(QLatin1String("w:val"), QLatin1String("clear"));
            w.writeEndElement();
        }
        w.writeEndElement();

        w.writeEndElement();
        w.writeEndElement();
        w.writeEndDocument();

        if (!addZipFile(&zip, QLatin1String("word/document.xml"), xml.toUtf8())) {
            mz_zip_writer_end(&zip);
            return {};
        }
    }

    {
        QString xml;
        QXmlStreamWriter w(&xml);
        w.setAutoFormatting(true);
        w.writeStartDocument();
        w.writeStartElement(QLatin1String("Relationships"));
        w.writeDefaultNamespace(NS_REL);
        w.writeStartElement(QLatin1String("Relationship"));
        w.writeAttribute(QLatin1String("Id"), QLatin1String("rId1"));
        w.writeAttribute(QLatin1String("Type"), QLatin1String("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"));
        w.writeAttribute(QLatin1String("Target"), QLatin1String("styles.xml"));
        w.writeEndElement();
        for (int i = 0; i < imageRels.size(); ++i) {
            w.writeStartElement(QLatin1String("Relationship"));
            w.writeAttribute(QLatin1String("Id"), imageRels[i]);
            w.writeAttribute(QLatin1String("Type"), QLatin1String("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"));
            w.writeAttribute(QLatin1String("Target"), imageTargets[i]);
            w.writeEndElement();
        }
        for (auto it = hyperlinkRelIdToUrl.constBegin(); it != hyperlinkRelIdToUrl.constEnd(); ++it) {
            w.writeStartElement(QLatin1String("Relationship"));
            w.writeAttribute(QLatin1String("Id"), it.key());
            w.writeAttribute(QLatin1String("Type"), QLatin1String("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"));
            w.writeAttribute(QLatin1String("Target"), it.value());
            w.writeAttribute(QLatin1String("TargetMode"), QLatin1String("External"));
            w.writeEndElement();
        }
        w.writeEndElement();
        w.writeEndDocument();
        if (!addZipFile(&zip, QLatin1String("word/_rels/document.xml.rels"), xml.toUtf8())) {
            mz_zip_writer_end(&zip);
            return {};
        }
    }

    {
        struct ListEntry {
            QTextListFormat::Style style;
            int indent;
        };
        QSet<quintptr> seen;
        QList<QPair<quintptr, ListEntry>> uniqueLists; // document order = numId order
        QTextBlock b = doc->begin();
        while (b.isValid()) {
            QTextList *list = b.textList();
            if (list) {
                quintptr ptr = reinterpret_cast<quintptr>(list);
                if (!seen.contains(ptr)) {
                    seen.insert(ptr);
                    QTextListFormat lf = list->format();
                    uniqueLists.append({ ptr, { lf.style(), lf.indent() } });
                }
            }
            b = b.next();
        }

        QString xml;
        QXmlStreamWriter w(&xml);
        w.setAutoFormatting(true);
        w.writeStartDocument();
        w.writeStartElement(QLatin1String("w:numbering"));
        w.writeNamespace(NS_W, QLatin1String("w"));

        int abstractNumId = 0;
        for (int listIdx = 0; listIdx < uniqueLists.size(); ++listIdx) {
            const auto &entry = uniqueLists.at(listIdx);
            int numId = listIdx + 1; // matches the document.xml numId assignment
            ++abstractNumId;

            w.writeStartElement(QLatin1String("w:abstractNum"));
            w.writeAttribute(QLatin1String("w:abstractNumId"), QString::number(abstractNumId));

            int levels = qMax(1, entry.second.indent);
            for (int lvl = 0; lvl < levels; ++lvl) {
                w.writeStartElement(QLatin1String("w:lvl"));
                w.writeAttribute(QLatin1String("w:ilvl"), QString::number(lvl));
                w.writeStartElement(QLatin1String("w:start"));
                w.writeAttribute(QLatin1String("w:val"), QLatin1String("1"));
                w.writeEndElement();
                w.writeStartElement(QLatin1String("w:numFmt"));
                QString fmtVal;
                switch (entry.second.style) {
                case QTextListFormat::ListDecimal:    fmtVal = QLatin1String("decimal"); break;
                case QTextListFormat::ListLowerAlpha: fmtVal = QLatin1String("lowerLetter"); break;
                case QTextListFormat::ListUpperAlpha: fmtVal = QLatin1String("upperLetter"); break;
                case QTextListFormat::ListLowerRoman: fmtVal = QLatin1String("lowerRoman"); break;
                case QTextListFormat::ListUpperRoman: fmtVal = QLatin1String("upperRoman"); break;
                default:                              fmtVal = QLatin1String("bullet"); break;
                }
                w.writeAttribute(QLatin1String("w:val"), fmtVal);
                w.writeEndElement();
                w.writeStartElement(QLatin1String("w:lvlJc"));
                w.writeAttribute(QLatin1String("w:val"), QLatin1String("left"));
                w.writeEndElement();
                w.writeStartElement(QLatin1String("w:pPr"));
                w.writeStartElement(QLatin1String("w:ind"));
                int leftIndent = (entry.second.indent + lvl) * 360;
                w.writeAttribute(QLatin1String("w:left"), QString::number(leftIndent));
                w.writeAttribute(QLatin1String("w:hanging"), QLatin1String("360"));
                w.writeEndElement();
                w.writeEndElement();
                w.writeEndElement();
            }

            w.writeEndElement();

            w.writeStartElement(QLatin1String("w:num"));
            w.writeAttribute(QLatin1String("w:numId"), QString::number(numId));
            w.writeTextElement(QLatin1String("w:abstractNumId"), QString::number(abstractNumId));
            w.writeEndElement();
        }

        w.writeEndElement();
        w.writeEndDocument();
        if (!addZipFile(&zip, QLatin1String("word/numbering.xml"), xml.toUtf8())) {
            mz_zip_writer_end(&zip);
            return {};
        }
    }

    {
        QString xml;
        QXmlStreamWriter w(&xml);
        w.setAutoFormatting(true);
        w.writeStartDocument();
        w.writeStartElement(QLatin1String("w:styles"));
        w.writeDefaultNamespace(NS_W);
        w.writeStartElement(QLatin1String("w:style"));
        w.writeAttribute(QLatin1String("w:type"), QLatin1String("paragraph"));
        w.writeAttribute(QLatin1String("w:default"), QLatin1String("1"));
        w.writeAttribute(QLatin1String("w:styleId"), QLatin1String("Normal"));
        w.writeStartElement(QLatin1String("w:name"));
        w.writeAttribute(QLatin1String("w:val"), QLatin1String("Normal"));
        w.writeEndElement();
        w.writeEndElement();
        for (const QString &sn : usedStyles) {
            if (sn == QLatin1String("Normal")) continue;
            w.writeStartElement(QLatin1String("w:style"));
            w.writeAttribute(QLatin1String("w:type"), QLatin1String("paragraph"));
            w.writeAttribute(QLatin1String("w:styleId"), sn);
            w.writeStartElement(QLatin1String("w:name"));
            w.writeAttribute(QLatin1String("w:val"), sn);
            w.writeEndElement();
            w.writeStartElement(QLatin1String("w:basedOn"));
            w.writeAttribute(QLatin1String("w:val"), QLatin1String("Normal"));
            w.writeEndElement();
            w.writeEndElement();
        }
        w.writeEndElement();
        w.writeEndDocument();
        if (!addZipFile(&zip, QLatin1String("word/styles.xml"), xml.toUtf8())) {
            mz_zip_writer_end(&zip);
            return {};
        }
    }

    if (!mz_zip_writer_finalize_archive(&zip)) {
        mz_zip_writer_end(&zip);
        return {};
    }
    mz_zip_writer_end(&zip);
    return zipBuffer;
}
