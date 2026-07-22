#include "DocxConverter.h"

#include <QtCore/private/qzipreader_p.h>
#include <QtCore/private/qzipwriter_p.h>

#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QTextCursor>
#include <QTextBlock>
#include <QTextBlockFormat>
#include <QTextCharFormat>
#include <QTextList>
#include <QTextListFormat>
#include <QTextTable>
#include <QTextFrame>
#include <QTextImageFormat>
#include <QTextFrameFormat>
#include <QBuffer>
#include <QImage>
#include <QUrl>
#include <QColor>
#include <QMap>
#include <QVector>
#include <QPair>
#include <QFileInfo>
#include <QtNumeric>
#include <QSet>
#include <QDebug>

#define NS_W  QStringLiteral("http://schemas.openxmlformats.org/wordprocessingml/2006/main")
#define NS_REL QStringLiteral("http://schemas.openxmlformats.org/package/2006/relationships")
#define NS_A  QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/main")
#define NS_WP QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
#define NS_CT QStringLiteral("http://schemas.openxmlformats.org/package/2006/content-types")
#define NS_R  QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships")

static constexpr qreal EMU_PER_PT    = 12700.0;
static constexpr qreal TWIP_PER_PT   = 20.0;
static constexpr qreal HALFPT_PER_PT = 2.0;

static qreal     emuToPt(qlonglong emu)   { return static_cast<qreal>(emu) / EMU_PER_PT; }
static qlonglong ptToEmu(qreal pt)        { return qRound64(pt * EMU_PER_PT); }
static qreal     twipToPt(int twip)       { return static_cast<qreal>(twip) / TWIP_PER_PT; }
static int       ptToTwip(qreal pt)       { return qRound(pt * TWIP_PER_PT); }
static qreal     halfPtToPt(int half)     { return static_cast<qreal>(half) / HALFPT_PER_PT; }
static int       ptToHalfPt(qreal pt)     { return qRound(pt * HALFPT_PER_PT); }
static int       propToDocxLine(int pct)  { return qRound(pct * 240.0 / 100.0); }
static int       docxLineToProportional(int line)
{
    if (line <= 0) return 100;
    return qRound(line * 100.0 / 240.0);
}

static void skipElement(QXmlStreamReader &r)
{
    int depth = 1;
    while (depth > 0 && r.readNext() != QXmlStreamReader::Invalid) {
        if (r.isStartElement())  ++depth;
        if (r.isEndElement())    --depth;
    }
}

static QColor parseColor(const QString &val, const QColor &fallback = QColor())
{
    if (val.isEmpty() || val == QLatin1String("auto"))
        return fallback;
    if (val.size() == 6) {
        bool ok;
        int rgb = val.toInt(&ok, 16);
        if (ok)
            return QColor((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
    }
    return fallback;
}

static void readRPr(QXmlStreamReader &r, QTextCharFormat &fmt)
{
    while (r.readNext() != QXmlStreamReader::Invalid) {
        if (r.isEndElement() && r.namespaceUri() == NS_W && r.name().toString() == QLatin1String("rPr"))
            return;
        if (!r.isStartElement() || r.namespaceUri() != NS_W)
            continue;
        QString n = r.name().toString();

        if (n == QLatin1String("b") || n == QLatin1String("bCs")) {
            QString v = r.attributes().value(QLatin1String("w:val")).toString();
            if (v.isEmpty() || v == QLatin1String("true") || v == QLatin1String("1"))
                fmt.setFontWeight(QFont::Bold);
        } else if (n == QLatin1String("i") || n == QLatin1String("iCs")) {
            QString v = r.attributes().value(QLatin1String("w:val")).toString();
            if (v.isEmpty() || v == QLatin1String("true") || v == QLatin1String("1"))
                fmt.setFontItalic(true);
        } else if (n == QLatin1String("u")) {
            QString v = r.attributes().value(QLatin1String("w:val")).toString();
            if (!v.isEmpty() && v != QLatin1String("none"))
                fmt.setFontUnderline(true);
        } else if (n == QLatin1String("strike")) {
            QString v = r.attributes().value(QLatin1String("w:val")).toString();
            if (v.isEmpty() || v == QLatin1String("true") || v == QLatin1String("1"))
                fmt.setFontStrikeOut(true);
        } else if (n == QLatin1String("vertAlign")) {
            QString v = r.attributes().value(QLatin1String("w:val")).toString();
            if (v == QLatin1String("subscript"))
                fmt.setVerticalAlignment(QTextCharFormat::AlignSubScript);
            else if (v == QLatin1String("superscript"))
                fmt.setVerticalAlignment(QTextCharFormat::AlignSuperScript);
        } else if (n == QLatin1String("sz") || n == QLatin1String("szCs")) {
            QString v = r.attributes().value(QLatin1String("w:val")).toString();
            int half = v.toInt();
            if (half > 0 && (fmt.fontPointSize() <= 0 || n == QLatin1String("sz")))
                fmt.setFontPointSize(halfPtToPt(half));
        } else if (n == QLatin1String("rFonts")) {
            QStringList families;
            QString ascii = r.attributes().value(QLatin1String("w:ascii")).toString();
            QString hAnsi = r.attributes().value(QLatin1String("w:hAnsi")).toString();
            QString eastAsia = r.attributes().value(QLatin1String("w:eastAsia")).toString();
            QString cs = r.attributes().value(QLatin1String("w:cs")).toString();
            if (!ascii.isEmpty())     families.append(ascii);
            if (!hAnsi.isEmpty() && hAnsi != ascii) families.append(hAnsi);
            if (!eastAsia.isEmpty())  families.append(eastAsia);
            if (!cs.isEmpty() && cs != ascii && cs != hAnsi) families.append(cs);
            if (!families.isEmpty())
                fmt.setFontFamilies(families);
        } else if (n == QLatin1String("color")) {
            QColor c = parseColor(r.attributes().value(QLatin1String("w:val")).toString());
            if (c.isValid())
                fmt.setForeground(c);
        } else if (n == QLatin1String("shd")) {
            QColor c = parseColor(r.attributes().value(QLatin1String("w:fill")).toString());
            if (c.isValid())
                fmt.setBackground(c);
        } else {
            skipElement(r);
        }
    }
}

struct PPrResult {
    int numId = -1;
    int ilvl = 0;
    QTextBlockFormat blockFmt;
    Qt::Alignment align = Qt::AlignLeft;
};

static PPrResult readPPr(QXmlStreamReader &r)
{
    PPrResult res;
    while (r.readNext() != QXmlStreamReader::Invalid) {
        if (r.isEndElement() && r.namespaceUri() == NS_W && r.name().toString() == QLatin1String("pPr"))
            return res;
        if (!r.isStartElement() || r.namespaceUri() != NS_W)
            continue;
        QString n = r.name().toString();

        if (n == QLatin1String("jc")) {
            QString v = r.attributes().value(QLatin1String("w:val")).toString();
            if (v == QLatin1String("center"))         res.align = Qt::AlignCenter;
            else if (v == QLatin1String("right"))     res.align = Qt::AlignRight;
            else if (v == QLatin1String("both"))      res.align = Qt::AlignJustify;
            else                                        res.align = Qt::AlignLeft;
        } else if (n == QLatin1String("spacing")) {
            QString line = r.attributes().value(QLatin1String("w:line")).toString();
            if (!line.isEmpty())
                res.blockFmt.setLineHeight(docxLineToProportional(line.toInt()), QTextBlockFormat::ProportionalHeight);
        } else if (n == QLatin1String("ind")) {
            auto a = r.attributes();
            QString left = a.value(QLatin1String("w:left")).toString();
            QString firstLine = a.value(QLatin1String("w:firstLine")).toString();
            QString hanging = a.value(QLatin1String("w:hanging")).toString();
            if (!left.isEmpty())
                res.blockFmt.setLeftMargin(twipToPt(left.toInt()));
            if (!firstLine.isEmpty())
                res.blockFmt.setTextIndent(twipToPt(firstLine.toInt()));
            else if (!hanging.isEmpty())
                res.blockFmt.setTextIndent(-twipToPt(hanging.toInt()));
        } else if (n == QLatin1String("numPr")) {
            while (r.readNext() != QXmlStreamReader::Invalid) {
                if (r.isEndElement() && r.namespaceUri() == NS_W && r.name().toString() == QLatin1String("numPr"))
                    break;
                if (!r.isStartElement()) continue;
                QString sn = r.name().toString();
                if (sn == QLatin1String("numId"))
                    res.numId = r.readElementText().toInt();
                else if (sn == QLatin1String("ilvl"))
                    res.ilvl = r.readElementText().toInt();
            }
        } else {
            skipElement(r);
        }
    }
    return res;
}

struct NumLevel {
    QTextListFormat::Style style = QTextListFormat::ListDisc;
    int indent = 1;
};

struct NumDef {
    QVector<NumLevel> levels;
};

static QMap<int, NumDef> parseNumbering(const QByteArray &xml)
{
    QMap<int, NumDef> result;
    QMap<int, NumDef> abstractDefs;

    QXmlStreamReader r(xml);
    while (!r.atEnd()) {
        r.readNext();
        if (!r.isStartElement() || r.namespaceUri() != NS_W)
            continue;
        QString n = r.name().toString();

        if (n == QLatin1String("abstractNum")) {
            int absId = r.attributes().value(QLatin1String("w:abstractNumId")).toInt();
            NumDef def;
            def.levels.resize(9);
            while (r.readNext() != QXmlStreamReader::Invalid) {
                if (r.isEndElement() && r.name().toString() == QLatin1String("abstractNum"))
                    break;
                if (!r.isStartElement() || r.name().toString() != QLatin1String("lvl"))
                    continue;
                QString ilvlStr = r.attributes().value(QLatin1String("w:ilvl")).toString();
                int ilvl = ilvlStr.isEmpty() ? 0 : ilvlStr.toInt();
                if (ilvl < 0 || ilvl >= 9) { skipElement(r); continue; }

                NumLevel &lv = def.levels[ilvl];
                while (r.readNext() != QXmlStreamReader::Invalid) {
                    if (r.isEndElement() && r.name().toString() == QLatin1String("lvl"))
                        break;
                    if (!r.isStartElement()) continue;
                    QString ln = r.name().toString();
                    if (ln == QLatin1String("numFmt")) {
                        QString fmt = r.attributes().value(QLatin1String("w:val")).toString();
                        if (fmt == QLatin1String("decimal"))          lv.style = QTextListFormat::ListDecimal;
                        else if (fmt == QLatin1String("bullet"))      lv.style = QTextListFormat::ListDisc;
                        else if (fmt == QLatin1String("lowerLetter")) lv.style = QTextListFormat::ListLowerAlpha;
                        else if (fmt == QLatin1String("upperLetter")) lv.style = QTextListFormat::ListUpperAlpha;
                        else if (fmt == QLatin1String("lowerRoman"))  lv.style = QTextListFormat::ListLowerRoman;
                        else if (fmt == QLatin1String("upperRoman"))  lv.style = QTextListFormat::ListUpperRoman;
                        else                                           lv.style = QTextListFormat::ListDisc;
                    } else if (ln == QLatin1String("ind")) {
                        QString left = r.attributes().value(QLatin1String("w:left")).toString();
                        if (!left.isEmpty())
                            lv.indent = qMax(1, left.toInt() / 360);
                    } else {
                        skipElement(r);
                    }
                }
            }
            abstractDefs[absId] = def;
        } else if (n == QLatin1String("num")) {
            int numId = r.attributes().value(QLatin1String("w:numId")).toInt();
            int absId = 0;
            while (r.readNext() != QXmlStreamReader::Invalid) {
                if (r.isEndElement() && r.name().toString() == QLatin1String("num"))
                    break;
                if (r.isStartElement() && r.name().toString() == QLatin1String("abstractNumId"))
                    absId = r.readElementText().toInt();
            }
            if (abstractDefs.contains(absId))
                result[numId] = abstractDefs[absId];
        }
    }
    return result;
}

struct RunFragment {
    QString text;
    bool isLineBreak = false;
    bool isTab = false;
    bool isImage = false;
    QByteArray imageData;
    QString imageFileName;
    qreal imageWidth = 0;
    qreal imageHeight = 0;
    QTextCharFormat charFmt;
};

struct ImageRel {
    QByteArray data;
    QString target;
};

static QVector<RunFragment> collectRuns(QXmlStreamReader &r,
                                        const QMap<QString, ImageRel> &imagesByRelId)
{
    QVector<RunFragment> fragments;
    QTextCharFormat currentFmt;

    while (r.readNext() != QXmlStreamReader::Invalid) {
        if (r.isEndElement() && r.namespaceUri() == NS_W && r.name().toString() == QLatin1String("r")) {
            currentFmt = QTextCharFormat();
            continue;
        }
        if (r.isEndElement() && r.namespaceUri() == NS_W && r.name().toString() == QLatin1String("p"))
            break;
        if (!r.isStartElement())
            continue;
        QString ns = r.namespaceUri().toString();
        QString n = r.name().toString();

        if (ns == NS_W && n == QLatin1String("rPr")) {
            readRPr(r, currentFmt);
        } else if (ns == NS_W && n == QLatin1String("t")) {
            RunFragment f;
            f.text = r.readElementText();
            f.charFmt = currentFmt;
            fragments.append(f);
        } else if (ns == NS_W && n == QLatin1String("br")) {
            RunFragment f;
            f.isLineBreak = true;
            fragments.append(f);
        } else if (ns == NS_W && n == QLatin1String("tab")) {
            RunFragment f;
            f.isTab = true;
            fragments.append(f);
        } else if (ns == NS_W && n == QLatin1String("drawing")) {
            RunFragment f;
            f.isImage = true;
            bool done = false;
            while (!done && r.readNext() != QXmlStreamReader::Invalid) {
                if (r.isEndElement() && r.namespaceUri().toString() == NS_W && r.name().toString() == QLatin1String("drawing"))
                    done = true;
                if (!r.isStartElement())
                    continue;
                QString rns = r.namespaceUri().toString();
                QString rn = r.name().toString();
                if (rns == NS_WP && rn == QLatin1String("extent")) {
                    QString cx = r.attributes().value(QLatin1String("cx")).toString();
                    QString cy = r.attributes().value(QLatin1String("cy")).toString();
                    if (!cx.isEmpty()) f.imageWidth = emuToPt(cx.toLongLong());
                    if (!cy.isEmpty()) f.imageHeight = emuToPt(cy.toLongLong());
                }
                if (rns == NS_A && rn == QLatin1String("blip")) {
                    QString embed = r.attributes().value(QLatin1String("r:embed")).toString();
                    if (!embed.isEmpty() && imagesByRelId.contains(embed)) {
                        f.imageData = imagesByRelId[embed].data;
                        f.imageFileName = imagesByRelId[embed].target;
                    }
                }
            }
            if (f.imageData.size() >= 4) {
                fragments.append(f);
            }
        } else {
            skipElement(r);
        }
    }
    return fragments;
}

static void insertRunsIntoCursor(QTextCursor &cursor, const QVector<RunFragment> &runs,
                                 QTextDocument *doc, const QMap<QString, ImageRel> &)
{
    for (const RunFragment &rf : runs) {
        if (rf.isLineBreak) {
            cursor.insertText(QStringLiteral("\n"), rf.charFmt);
        } else if (rf.isTab) {
            cursor.insertText(QStringLiteral("\t"), rf.charFmt);
        } else if (rf.isImage) {
            QImage img;
            if (img.loadFromData(rf.imageData)) {
                QString ext = QFileInfo(rf.imageFileName).suffix().toLower();
                QString name = QStringLiteral("docx_embed_%1.%2").arg(
                    reinterpret_cast<quintptr>(&rf)).arg(ext.isEmpty() ? QStringLiteral("png") : ext);
                doc->addResource(QTextDocument::ImageResource, QUrl(name), img);
                QTextImageFormat imgFmt;
                imgFmt.setName(name);
                if (rf.imageWidth > 0) imgFmt.setWidth(rf.imageWidth);
                if (rf.imageHeight > 0) imgFmt.setHeight(rf.imageHeight);
                cursor.insertImage(imgFmt);
            }
        } else {
            cursor.insertText(rf.text, rf.charFmt);
        }
    }
}

struct ParaBuffer {
    QString xml;
    int numId = -1;
    int ilvl = 0;
    QTextBlockFormat blockFmt;
    Qt::Alignment align = Qt::AlignLeft;
    QVector<RunFragment> runs;
};

static void parseTable(QXmlStreamReader &xml, QTextCursor &cursor,
                       const QMap<QString, ImageRel> &imagesByRelId,
                       QTextDocument *doc,
                       QMap<QPair<int,int>, QTextList*> &activeLists,
                       const QMap<int, NumDef> &numbering)
{
    QString tblXml;
    {
        QXmlStreamWriter pw(&tblXml);
        pw.setAutoFormatting(false);
        pw.writeStartElement(QLatin1String("w:tbl"));
        for (const auto &attr : xml.attributes())
            pw.writeAttribute(attr.qualifiedName().toString(), attr.value().toString());
        pw.writeNamespace(NS_W, QLatin1String("w"));
            pw.writeNamespace(NS_WP, QLatin1String("wp"));
            pw.writeNamespace(NS_A, QLatin1String("a"));
            pw.writeNamespace(NS_R, QLatin1String("r"));

        int depth = 1;
        while (depth > 0 && xml.readNext() != QXmlStreamReader::Invalid) {
            if (xml.isStartElement()) {
                pw.writeStartElement(xml.qualifiedName().toString());
                for (const auto &attr : xml.attributes())
                    pw.writeAttribute(attr.qualifiedName().toString(), attr.value().toString());
                ++depth;
            } else if (xml.isEndElement()) {
                pw.writeEndElement();
                --depth;
            } else if (xml.isCharacters()) {
                pw.writeCharacters(xml.text().toString());
            } else if (xml.isCDATA()) {
                pw.writeCDATA(xml.text().toString());
            } else if (xml.isComment()) {
                pw.writeComment(xml.text().toString());
            }
        }
        pw.writeEndDocument();
    }

    struct CellContent {
        QVector<ParaBuffer> paragraphs;
    };
    QVector<QVector<CellContent>> grid;
    int currentRow = -1;

    QXmlStreamReader tx(tblXml);
    while (!tx.atEnd()) {
        tx.readNext();
        if (!tx.isStartElement()) continue;
        if (tx.name().toString() == QLatin1String("tr")) {
            grid.append(QVector<CellContent>());
            ++currentRow;
        } else if (tx.name().toString() == QLatin1String("tc") && currentRow >= 0) {
            CellContent cell;
            while (!tx.atEnd()) {
                tx.readNext();
                if (tx.isEndElement() && tx.name().toString() == QLatin1String("tc"))
                    break;
                if (!tx.isStartElement()) continue;
                if (tx.name().toString() == QLatin1String("p")) {
                    QByteArray paraBuf;
                    {
                        QXmlStreamWriter pw(&paraBuf);
                        pw.setAutoFormatting(false);
                        pw.writeStartElement(QLatin1String("w:p"));
                        for (const auto &attr : tx.attributes())
                            pw.writeAttribute(attr.qualifiedName().toString(), attr.value().toString());
                        pw.writeNamespace(NS_W, QLatin1String("w"));
            pw.writeNamespace(NS_WP, QLatin1String("wp"));
            pw.writeNamespace(NS_A, QLatin1String("a"));
            pw.writeNamespace(NS_R, QLatin1String("r"));
                        int depth = 1;
                        while (depth > 0 && tx.readNext() != QXmlStreamReader::Invalid) {
                            if (tx.isStartElement()) {
                                pw.writeStartElement(tx.qualifiedName().toString());
                                for (const auto &attr : tx.attributes())
                                    pw.writeAttribute(attr.qualifiedName().toString(), attr.value().toString());
                                ++depth;
                            } else if (tx.isEndElement()) {
                                pw.writeEndElement();
                                --depth;
                            } else if (tx.isCharacters()) {
                                pw.writeCharacters(tx.text().toString());
                            } else if (tx.isCDATA()) {
                                pw.writeCDATA(tx.text().toString());
                            } else if (tx.isComment()) {
                                pw.writeComment(tx.text().toString());
                            }
                        }
                        pw.writeEndDocument();
                    }

                    ParaBuffer pb;
                    QXmlStreamReader px(paraBuf);
                    while (!px.atEnd()) {
                        px.readNext();
                        if (px.isStartElement() && px.name().toString() == QLatin1String("p"))
                            break;
                    }
                    while (!px.atEnd()) {
                        px.readNext();
                        if (px.isEndElement() && px.name().toString() == QLatin1String("p"))
                            break;
                        if (!px.isStartElement()) continue;
                        if (px.namespaceUri().toString() != NS_W) continue;
                        QString pn = px.name().toString();
                        if (pn == QLatin1String("pPr")) {
                            PPrResult ppr = readPPr(px);
                            pb.numId = ppr.numId;
                            pb.ilvl = ppr.ilvl;
                            pb.blockFmt = ppr.blockFmt;
                            pb.align = ppr.align;
                        } else if (pn == QLatin1String("r")) {
                            QVector<RunFragment> runs = collectRuns(px, imagesByRelId);
                            pb.runs = runs;
                        }
                    }
                    cell.paragraphs.append(pb);
                }
            }
            if (currentRow >= 0 && currentRow < grid.size())
                grid[currentRow].append(cell);
        }
    }

    if (grid.isEmpty()) return;

    int rows = grid.size();
    int cols = 0;
    for (const auto &row : grid)
        cols = qMax(cols, row.size());
    if (cols == 0) return;

    QTextTable *table = cursor.insertTable(rows, cols);

    for (int r = 0; r < rows; ++r) {
        for (int c = 0; c < cols && c < grid[r].size(); ++c) {
            QTextTableCell cell = table->cellAt(r, c);
            QTextCursor cellCursor = cell.firstCursorPosition();
            cellCursor.movePosition(QTextCursor::NextBlock, QTextCursor::KeepAnchor);
            cellCursor.removeSelectedText();

            bool firstPara = true;
            for (const ParaBuffer &pb : grid[r][c].paragraphs) {
                if (firstPara) {
                    cellCursor.setBlockFormat(pb.blockFmt);
                    firstPara = false;
                } else {
                    cellCursor.insertBlock(pb.blockFmt);
                }

                QPair<int,int> listKey(pb.numId, pb.ilvl);
                if (pb.numId >= 0 && numbering.contains(pb.numId)) {
                    const NumDef &nd = numbering[pb.numId];
                    int level = qMin(pb.ilvl, nd.levels.size() - 1);
                    const NumLevel &nl = nd.levels[level];

                    if (!activeLists.contains(listKey)) {
                        QTextListFormat lf;
                        lf.setStyle(nl.style);
                        lf.setIndent(nl.indent + level);
                        QTextList *list = cellCursor.createList(lf);
                        activeLists[listKey] = list;
                    }
                    activeLists[listKey]->add(cellCursor.block());
                }

                insertRunsIntoCursor(cellCursor, pb.runs, doc, imagesByRelId);
            }
        }
    }
}

bool DocxConverter::loadFromDocx(const QByteArray &zipData, QTextDocument *doc,
                                  QMarginsF &outMargins, QColor &outPageBackground)
{
    QBuffer buffer;
    buffer.setData(zipData);
    if (!buffer.open(QIODevice::ReadOnly))
        return false;

    QZipReader reader(&buffer);
    if (reader.status() != QZipReader::NoError) {
        qDebug() << "ZIP error:" << reader.status();
        return false;
    }

    QMap<QString, QByteArray> files;
    for (const QZipReader::FileInfo &fi : reader.fileInfoList())
        files[fi.filePath] = reader.fileData(fi.filePath);

    if (!files.contains(QLatin1String("word/document.xml")))
        return false;

    QMap<int, NumDef> numbering;
    if (files.contains(QLatin1String("word/numbering.xml")))
        numbering = parseNumbering(files[QLatin1String("word/numbering.xml")]);

    QMap<QString, ImageRel> imageDataByRelId;
    if (files.contains(QLatin1String("word/_rels/document.xml.rels"))) {
        QXmlStreamReader relR(files[QLatin1String("word/_rels/document.xml.rels")]);
        while (!relR.atEnd()) {
            relR.readNext();
            if (!relR.isStartElement() || relR.name().toString() != QLatin1String("Relationship"))
                continue;
            auto a = relR.attributes();
            QString id = a.value(QLatin1String("Id")).toString();
            QString target = a.value(QLatin1String("Target")).toString();
            QString type = a.value(QLatin1String("Type")).toString();
            if (!id.isEmpty() && !target.isEmpty()) {
                QString mediaPath = QLatin1String("word/") + target;
                if (files.contains(mediaPath)) {
                    ImageRel ir;
                    ir.data = files[mediaPath];
                    ir.target = target;
                    imageDataByRelId[id] = ir;
                }
            }
        }
    }

    doc->clear();
    QTextCursor cursor(doc);
    bool firstBlock = true;

    QMap<QPair<int,int>, QTextList*> activeLists;

    qreal pgW = 816, pgH = 1056;
    QMarginsF margins(96, 96, 96, 96);

    QByteArray docXmlRaw = files[QLatin1String("word/document.xml")];

    QXmlStreamReader xml(docXmlRaw);

    while (!xml.atEnd()) {
        xml.readNext();

        if (xml.isEndElement()) {
            if (xml.name().toString() == QLatin1String("document"))
                break;
            continue;
        }
        if (!xml.isStartElement())
            continue;

        QString n = xml.name().toString();

        if (n == QLatin1String("document") || n == QLatin1String("body"))
            continue;

        if (n == QLatin1String("p")) {
            QString paraXml;
            QXmlStreamWriter pw(&paraXml);
            pw.setAutoFormatting(false);
            pw.writeStartElement(QLatin1String("w:p"));
            for (const auto &attr : xml.attributes())
                pw.writeAttribute(attr.qualifiedName().toString(), attr.value().toString());
            pw.writeNamespace(NS_W, QLatin1String("w"));
            pw.writeNamespace(NS_WP, QLatin1String("wp"));
            pw.writeNamespace(NS_A, QLatin1String("a"));
            pw.writeNamespace(NS_R, QLatin1String("r"));

            int depth = 1;
            while (depth > 0 && xml.readNext() != QXmlStreamReader::Invalid) {
                if (xml.isStartElement()) {
                    pw.writeStartElement(xml.qualifiedName().toString());
                    for (const auto &attr : xml.attributes())
                        pw.writeAttribute(attr.qualifiedName().toString(), attr.value().toString());
                    ++depth;
                } else if (xml.isEndElement()) {
                    pw.writeEndElement();
                    --depth;
                } else if (xml.isCharacters()) {
                    pw.writeCharacters(xml.text().toString());
                } else if (xml.isCDATA()) {
                    pw.writeCDATA(xml.text().toString());
                } else if (xml.isComment()) {
                    pw.writeComment(xml.text().toString());
                }
            }
            pw.writeEndDocument();

            QXmlStreamReader px(paraXml);
            while (!px.atEnd()) {
                px.readNext();
                if (px.isStartElement() && px.name().toString() == QLatin1String("p"))
                    break;
            }

            PPrResult ppr;
            QVector<RunFragment> runs;

            while (!px.atEnd()) {
                px.readNext();
                if (px.isEndElement() && px.name().toString() == QLatin1String("p"))
                    break;
                if (!px.isStartElement())
                    continue;
                if (px.namespaceUri().toString() != NS_W)
                    continue;
                QString pn = px.name().toString();
                if (pn == QLatin1String("pPr")) {
                    ppr = readPPr(px);
                } else if (pn == QLatin1String("r")) {
                    runs = collectRuns(px, imageDataByRelId);
                }
            }

            if (firstBlock) {
                cursor.setBlockFormat(ppr.blockFmt);
                firstBlock = false;
            } else {
                cursor.insertBlock(ppr.blockFmt);
            }

            if (cursor.blockFormat().alignment() != ppr.align) {
                QTextBlockFormat bf = cursor.blockFormat();
                bf.setAlignment(ppr.align);
                cursor.setBlockFormat(bf);
            }

            QPair<int,int> listKey(ppr.numId, ppr.ilvl);
            if (ppr.numId >= 0 && numbering.contains(ppr.numId)) {
                const NumDef &nd = numbering[ppr.numId];
                int level = qMin(ppr.ilvl, nd.levels.size() - 1);
                const NumLevel &nl = nd.levels[level];

                if (!activeLists.contains(listKey)) {
                    QTextListFormat lf;
                    lf.setStyle(nl.style);
                    lf.setIndent(nl.indent + level);
                    QTextList *list = cursor.createList(lf);
                    activeLists[listKey] = list;
                }
                activeLists[listKey]->add(cursor.block());
            }

            insertRunsIntoCursor(cursor, runs, doc, imageDataByRelId);
            continue;
        }

        if (n == QLatin1String("tbl")) {
            parseTable(xml, cursor, imageDataByRelId, doc, activeLists, numbering);
            firstBlock = false;
            continue;
        }

        if (n == QLatin1String("sectPr")) {
            while (xml.readNext() != QXmlStreamReader::Invalid) {
                if (xml.isEndElement() && xml.name().toString() == QLatin1String("sectPr"))
                    break;
                if (!xml.isStartElement()) continue;
                QString sn = xml.name().toString();
                if (sn == QLatin1String("pgSz")) {
                    auto a = xml.attributes();
                    QString w = a.value(QLatin1String("w:w")).toString();
                    QString h = a.value(QLatin1String("w:h")).toString();
                    if (!w.isEmpty()) pgW = twipToPt(w.toInt());
                    if (!h.isEmpty()) pgH = twipToPt(h.toInt());
                    if (!w.isEmpty()) pgW = twipToPt(w.toInt());
                    if (!h.isEmpty()) pgH = twipToPt(h.toInt());
                } else if (sn == QLatin1String("pgMar")) {
                    auto a = xml.attributes();
                    QString top = a.value(QLatin1String("w:top")).toString();
                    QString bottom = a.value(QLatin1String("w:bottom")).toString();
                    QString left = a.value(QLatin1String("w:left")).toString();
                    QString right = a.value(QLatin1String("w:right")).toString();
                    if (!top.isEmpty())    margins.setTop(twipToPt(top.toInt()));
                    if (!bottom.isEmpty()) margins.setBottom(twipToPt(bottom.toInt()));
                    if (!left.isEmpty())   margins.setLeft(twipToPt(left.toInt()));
                    if (!right.isEmpty())  margins.setRight(twipToPt(right.toInt()));
                } else {
                    skipElement(xml);
                }
            }
            continue;
        }

        skipElement(xml);
    }

    if (xml.hasError()) {
        qDebug() << "XML error:" << xml.errorString() << "at line" << xml.lineNumber();
        return false;
    }

    doc->setPageSize(QSizeF(pgW, pgH));
    outMargins = margins;
    outPageBackground = QColor(Qt::white);

    return true;
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

static void writeTableXml(QXmlStreamWriter &w, QTextTable *table,
                          QTextDocument *doc,
                          int &imageCounter, int &relationCounter,
                          QStringList &imageRels, QStringList &imageTargets,
                          QZipWriter &zip)
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
        for (int col = 0; col < table->columns(); ++col) {
            QTextTableCell cell = table->cellAt(row, col);
            if (cell.row() != row || cell.column() != col)
                continue;

            w.writeStartElement(QLatin1String("w:tc"));

            QTextCursor cellCursor = cell.firstCursorPosition();
            QTextBlock block = cellCursor.block();
            int cellEnd = cell.lastPosition();

            while (block.isValid() && block.position() <= cellEnd) {
                QTextBlockFormat bf = block.blockFormat();

                w.writeStartElement(QLatin1String("w:p"));
                w.writeStartElement(QLatin1String("w:pPr"));

                Qt::Alignment align = bf.alignment();
                if (align == Qt::AlignCenter)
                    w.writeTextElement(QLatin1String("w:jc"), QLatin1String("center"));
                else if (align == Qt::AlignRight)
                    w.writeTextElement(QLatin1String("w:jc"), QLatin1String("right"));
                else if (align == Qt::AlignJustify)
                    w.writeTextElement(QLatin1String("w:jc"), QLatin1String("both"));

                if (bf.leftMargin() > 0 || bf.textIndent() != 0) {
                    w.writeStartElement(QLatin1String("w:ind"));
                    if (bf.leftMargin() > 0)
                        w.writeAttribute(QLatin1String("w:left"), QString::number(ptToTwip(bf.leftMargin())));
                    if (bf.textIndent() != 0) {
                        if (bf.textIndent() > 0)
                            w.writeAttribute(QLatin1String("w:firstLine"), QString::number(ptToTwip(bf.textIndent())));
                        else
                            w.writeAttribute(QLatin1String("w:hanging"), QString::number(ptToTwip(-bf.textIndent())));
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
                                imgBuf.open(QIODevice::WriteOnly);
                                img.save(&imgBuf, fmt.toLatin1().constData());
                                imgBuf.close();
                                zip.addFile(QLatin1String("word/") + target, imgData);

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
                                w.writeAttribute(QLatin1String("cx"), QString::number(ptToEmu(iw)));
                                w.writeAttribute(QLatin1String("cy"), QString::number(ptToEmu(ih)));
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
                                w.writeAttribute(QLatin1String("cx"), QString::number(ptToEmu(iw)));
                                w.writeAttribute(QLatin1String("cy"), QString::number(ptToEmu(ih)));
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
                                w.writeEndElement();
                            }
                        }
                        continue;
                    }

                    if (text == QLatin1String("\n")) {
                        w.writeStartElement(QLatin1String("w:r"));
                        w.writeEmptyElement(QLatin1String("w:br"));
                        w.writeEndElement();
                        continue;
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
                    if (cf.verticalAlignment() == QTextCharFormat::AlignSubScript)
                        w.writeTextElement(QLatin1String("w:vertAlign"), QLatin1String("subscript"));
                    else if (cf.verticalAlignment() == QTextCharFormat::AlignSuperScript)
                        w.writeTextElement(QLatin1String("w:vertAlign"), QLatin1String("superscript"));

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
                        w.writeTextElement(QLatin1String("w:sz"), QString::number(half));
                        w.writeTextElement(QLatin1String("w:szCs"), QString::number(half));
                    }

                    QColor fg = cf.foreground().color();
                    if (fg.isValid() && fg != QColor(Qt::black)) {
                        w.writeStartElement(QLatin1String("w:color"));
                        w.writeAttribute(QLatin1String("w:val"), fg.name().mid(1));
                        w.writeEndElement();
                    }

                    w.writeEndElement();
                    w.writeTextElement(QLatin1String("w:t"), text);
                    w.writeEndElement();
                }

                w.writeEndElement();
                block = block.next();
            }

            w.writeEndElement();
        }
        w.writeEndElement();
    }

    w.writeEndElement();
}

QByteArray DocxConverter::saveToDocx(const QTextDocument *doc,
                                      const QMarginsF &margins,
                                      const QColor &pageBackground)
{
    QBuffer buf;
    buf.open(QIODevice::WriteOnly);
    QZipWriter zip(&buf);
    zip.setCompressionPolicy(QZipWriter::AutoCompress);

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
        zip.addFile(QLatin1String("[Content_Types].xml"), xml.toUtf8());
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
        zip.addFile(QLatin1String("_rels/.rels"), xml.toUtf8());
    }

    int imageCounter = 0;
    int relationCounter = 1;
    QStringList imageRels;
    QStringList imageTargets;

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
                                  imageRels, imageTargets, zip);
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

            Qt::Alignment align = bf.alignment();
            if (align == Qt::AlignCenter)
                w.writeTextElement(QLatin1String("w:jc"), QLatin1String("center"));
            else if (align == Qt::AlignRight)
                w.writeTextElement(QLatin1String("w:jc"), QLatin1String("right"));
            else if (align == Qt::AlignJustify)
                w.writeTextElement(QLatin1String("w:jc"), QLatin1String("both"));

            if (bf.lineHeight() > 0 && bf.lineHeight() != 100) {
                w.writeStartElement(QLatin1String("w:spacing"));
                w.writeAttribute(QLatin1String("w:line"), QString::number(propToDocxLine(bf.lineHeight())));
                w.writeAttribute(QLatin1String("w:lineRule"), QLatin1String("auto"));
                w.writeEndElement();
            }

            if (bf.leftMargin() > 0 || bf.textIndent() != 0) {
                w.writeStartElement(QLatin1String("w:ind"));
                if (bf.leftMargin() > 0)
                    w.writeAttribute(QLatin1String("w:left"), QString::number(ptToTwip(bf.leftMargin())));
                if (bf.textIndent() != 0) {
                    if (bf.textIndent() > 0)
                        w.writeAttribute(QLatin1String("w:firstLine"), QString::number(ptToTwip(bf.textIndent())));
                    else
                        w.writeAttribute(QLatin1String("w:hanging"), QString::number(ptToTwip(-bf.textIndent())));
                }
                w.writeEndElement();
            }

            if (list) {
                QTextListFormat lf = list->format();
                int numId = (reinterpret_cast<quintptr>(list) % 10000) + 1;
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
                            imgBuf.open(QIODevice::WriteOnly);
                            img.save(&imgBuf, fmt.toLatin1().constData());
                            imgBuf.close();
                            zip.addFile(QLatin1String("word/") + target, imgData);

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
                            w.writeAttribute(QLatin1String("cx"), QString::number(ptToEmu(iw)));
                            w.writeAttribute(QLatin1String("cy"), QString::number(ptToEmu(ih)));
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
                            w.writeAttribute(QLatin1String("cx"), QString::number(ptToEmu(iw)));
                            w.writeAttribute(QLatin1String("cy"), QString::number(ptToEmu(ih)));
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
                            w.writeEndElement();
                        }
                    }
                    continue;
                }

                if (text == QLatin1String("\n")) {
                    w.writeStartElement(QLatin1String("w:r"));
                    w.writeEmptyElement(QLatin1String("w:br"));
                    w.writeEndElement();
                    continue;
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
                if (cf.verticalAlignment() == QTextCharFormat::AlignSubScript)
                    w.writeTextElement(QLatin1String("w:vertAlign"), QLatin1String("subscript"));
                else if (cf.verticalAlignment() == QTextCharFormat::AlignSuperScript)
                    w.writeTextElement(QLatin1String("w:vertAlign"), QLatin1String("superscript"));

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
                    w.writeTextElement(QLatin1String("w:sz"), QString::number(half));
                    w.writeTextElement(QLatin1String("w:szCs"), QString::number(half));
                }

                QColor fg = cf.foreground().color();
                if (fg.isValid() && fg != QColor(Qt::black)) {
                    w.writeStartElement(QLatin1String("w:color"));
                    w.writeAttribute(QLatin1String("w:val"), fg.name().mid(1));
                    w.writeEndElement();
                }

                w.writeEndElement();
                w.writeTextElement(QLatin1String("w:t"), text);
                w.writeEndElement();
            }

            w.writeEndElement();

            block = block.next();
        }

        w.writeStartElement(QLatin1String("w:sectPr"));
        w.writeStartElement(QLatin1String("w:pgSz"));
        w.writeAttribute(QLatin1String("w:w"), QString::number(ptToTwip(doc->pageSize().width())));
        w.writeAttribute(QLatin1String("w:h"), QString::number(ptToTwip(doc->pageSize().height())));
        w.writeEndElement();
        w.writeStartElement(QLatin1String("w:pgMar"));
        w.writeAttribute(QLatin1String("w:top"), QString::number(ptToTwip(margins.top())));
        w.writeAttribute(QLatin1String("w:right"), QString::number(ptToTwip(margins.right())));
        w.writeAttribute(QLatin1String("w:bottom"), QString::number(ptToTwip(margins.bottom())));
        w.writeAttribute(QLatin1String("w:left"), QString::number(ptToTwip(margins.left())));
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

        zip.addFile(QLatin1String("word/document.xml"), xml.toUtf8());
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
        w.writeEndElement();
        w.writeEndDocument();
        zip.addFile(QLatin1String("word/_rels/document.xml.rels"), xml.toUtf8());
    }

    {
        struct ListEntry {
            QTextListFormat::Style style;
            int indent;
        };
        QMap<quintptr, ListEntry> uniqueLists;
        QTextBlock b = doc->begin();
        while (b.isValid()) {
            QTextList *list = b.textList();
            if (list) {
                quintptr ptr = reinterpret_cast<quintptr>(list);
                if (!uniqueLists.contains(ptr)) {
                    QTextListFormat lf = list->format();
                    uniqueLists[ptr] = { lf.style(), lf.indent() };
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
        for (auto it = uniqueLists.constBegin(); it != uniqueLists.constEnd(); ++it) {
            int numId = (static_cast<int>(it.key() % 10000)) + 1;
            ++abstractNumId;

            w.writeStartElement(QLatin1String("w:abstractNum"));
            w.writeAttribute(QLatin1String("w:abstractNumId"), QString::number(abstractNumId));

            int levels = qMax(1, it.value().indent);
            for (int lvl = 0; lvl < levels; ++lvl) {
                w.writeStartElement(QLatin1String("w:lvl"));
                w.writeAttribute(QLatin1String("w:ilvl"), QString::number(lvl));
                w.writeStartElement(QLatin1String("w:start"));
                w.writeAttribute(QLatin1String("w:val"), QLatin1String("1"));
                w.writeEndElement();
                w.writeStartElement(QLatin1String("w:numFmt"));
                QString fmtVal;
                switch (it.value().style) {
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
                int leftIndent = (it.value().indent + lvl) * 360;
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
        zip.addFile(QLatin1String("word/numbering.xml"), xml.toUtf8());
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
        w.writeEndElement();
        w.writeEndDocument();
        zip.addFile(QLatin1String("word/styles.xml"), xml.toUtf8());
    }

    zip.close();
    return buf.data();
}
