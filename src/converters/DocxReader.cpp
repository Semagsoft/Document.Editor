#include "DocxConverter.h"
#include "DocxCommon.h"

#include "miniz.h"

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
#include <QDir>
#include <QSet>
#include <QDebug>
#include <climits>
#include <new>

// Skips the current element and all its descendants, leaving the reader
// positioned on the matching end element.
static void skipElement(QXmlStreamReader &r)
{
    int depth = 1;
    while (depth > 0 && r.readNext() != QXmlStreamReader::Invalid) {
        if (r.isStartElement())  ++depth;
        if (r.isEndElement())    --depth;
    }
}

// Tracks which namespace (prefix, uri) pairs are in scope while re-serializing
// a fragment, so that prefixes inherited from ancestors of the source document
// (declared on the <w:document> root) stay valid when the fragment is parsed in
// isolation. Because QXmlStreamWriter scopes declarations per element, each
// element (including siblings) must re-declare the prefixes it uses; tracking
// only the current scope stack keeps the output well-formed without emitting
// duplicate xmlns attributes for namespaces already in scope on an ancestor.
struct NsTracker {
    QList<QSet<QPair<QString, QString>>> scopes;

    NsTracker() { scopes.append(QSet<QPair<QString, QString>>()); }

    bool contains(const QString &prefix, const QString &uri) const
    {
        QPair<QString, QString> key(prefix, uri);
        for (const auto &scope : scopes) {
            if (scope.contains(key))
                return true;
        }
        return false;
    }

    void pushScope() { scopes.append(QSet<QPair<QString, QString>>()); }

    void popScope()
    {
        if (scopes.size() > 1)
            scopes.removeLast();
    }

    void declare(QXmlStreamWriter &pw, const QString &prefix, const QString &uri)
    {
        if (prefix.isEmpty() || contains(prefix, uri))
            return;
        pw.writeNamespace(uri, prefix);
        scopes.last().insert(QPair<QString, QString>(prefix, uri));
    }
};

static void declareElementNamespaces(QXmlStreamWriter &pw, const QXmlStreamReader &xml,
                                     NsTracker &ns)
{
    const auto nsDecls = xml.namespaceDeclarations();
    for (const auto &nsd : nsDecls)
        ns.declare(pw, nsd.prefix().toString(), nsd.namespaceUri().toString());
    if (!xml.prefix().isEmpty())
        ns.declare(pw, xml.prefix().toString(), xml.namespaceUri().toString());
    const auto attrs = xml.attributes();
    for (const auto &attr : attrs) {
        if (!attr.prefix().isEmpty())
            ns.declare(pw, attr.prefix().toString(), attr.namespaceUri().toString());
    }
}

static void copyStartElement(QXmlStreamWriter &pw, const QXmlStreamReader &xml, NsTracker &ns)
{
    pw.writeStartElement(xml.qualifiedName().toString());
    declareElementNamespaces(pw, xml, ns);
    for (const auto &attr : xml.attributes())
        pw.writeAttribute(attr.qualifiedName().toString(), attr.value().toString());
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

// Reads the current element's w:val attribute; falls back to its text
// content for OOXML written without it (older versions of this app emitted
// e.g. <w:sz>24</w:sz> / <w:jc>center</w:jc>). For an empty element (<w:b/>)
// readElementText() yields an empty string, which callers treat as "absent".
static QString readValAttrOrText(QXmlStreamReader &r)
{
    QString v = r.attributes().value(QLatin1String("w:val")).toString();
    if (v.isEmpty() && r.tokenType() == QXmlStreamReader::StartElement)
        v = r.readElementText();
    return v;
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
            QString v = readValAttrOrText(r);
            if (v.isEmpty() || v == QLatin1String("true") || v == QLatin1String("1"))
                fmt.setFontWeight(QFont::Bold);
        } else if (n == QLatin1String("i") || n == QLatin1String("iCs")) {
            QString v = readValAttrOrText(r);
            if (v.isEmpty() || v == QLatin1String("true") || v == QLatin1String("1"))
                fmt.setFontItalic(true);
        } else if (n == QLatin1String("u")) {
            QString v = readValAttrOrText(r);
            if (!v.isEmpty() && v != QLatin1String("none"))
                fmt.setFontUnderline(true);
        } else if (n == QLatin1String("strike")) {
            QString v = readValAttrOrText(r);
            if (v.isEmpty() || v == QLatin1String("true") || v == QLatin1String("1"))
                fmt.setFontStrikeOut(true);
        } else if (n == QLatin1String("vertAlign")) {
            QString v = readValAttrOrText(r);
            if (v == QLatin1String("subscript"))
                fmt.setVerticalAlignment(QTextCharFormat::AlignSubScript);
            else if (v == QLatin1String("superscript"))
                fmt.setVerticalAlignment(QTextCharFormat::AlignSuperScript);
        } else if (n == QLatin1String("sz") || n == QLatin1String("szCs")) {
            QString v = readValAttrOrText(r);
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
    QString pStyle;
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

        if (n == QLatin1String("pStyle")) {
            res.pStyle = r.attributes().value(QLatin1String("w:val")).toString();
        } else if (n == QLatin1String("jc")) {
            QString v = readValAttrOrText(r);
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
                res.blockFmt.setLeftMargin(twipToPx(left.toInt()));
            if (!firstLine.isEmpty())
                res.blockFmt.setTextIndent(twipToPx(firstLine.toInt()));
            else if (!hanging.isEmpty())
                res.blockFmt.setTextIndent(-twipToPx(hanging.toInt()));
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

static QVector<RunFragment> collectTextboxParagraphs(QXmlStreamReader &r,
                                                     const QMap<QString, ImageRel> &imagesByRelId,
                                                     const QMap<QString, QString> &hyperlinkTargets);

static QVector<RunFragment> collectRuns(QXmlStreamReader &r,
                                        const QMap<QString, ImageRel> &imagesByRelId,
                                        const QMap<QString, QString> &hyperlinkTargets = {},
                                        const QString &initialHref = {})
{
    QVector<RunFragment> fragments;
    QTextCharFormat currentFmt;
    QString currentHref = initialHref;

    while (r.readNext() != QXmlStreamReader::Invalid) {
        if (r.isEndElement() && r.namespaceUri() == NS_W && r.name().toString() == QLatin1String("r")) {
            currentFmt = QTextCharFormat();
            continue;
        }
        if (r.isEndElement() && r.namespaceUri() == NS_W && r.name().toString() == QLatin1String("p"))
            break;
        if (r.isEndElement() && r.namespaceUri() == NS_W && r.name().toString() == QLatin1String("hyperlink")) {
            currentHref.clear();
            continue;
        }
        if (!r.isStartElement())
            continue;
        QString ns = r.namespaceUri().toString();
        QString n = r.name().toString();

        if (ns == NS_W && n == QLatin1String("hyperlink")) {
            QString relId = r.attributes().value(QLatin1String("r:id")).toString();
            currentHref = hyperlinkTargets.value(relId);
        } else if (ns == NS_W && n == QLatin1String("rPr")) {
            readRPr(r, currentFmt);
        } else if (ns == NS_W && n == QLatin1String("t")) {
            RunFragment f;
            f.text = r.readElementText();
            f.charFmt = currentFmt;
            if (!currentHref.isEmpty()) {
                f.charFmt.setAnchor(true);
                f.charFmt.setAnchorHref(currentHref);
            }
            fragments.append(f);
        } else if (ns == NS_W && n == QLatin1String("r")) {
            currentFmt = QTextCharFormat();
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
                    if (!cx.isEmpty()) f.imageWidth = emuToPx(cx.toLongLong());
                    if (!cy.isEmpty()) f.imageHeight = emuToPx(cy.toLongLong());
                }
                if (rns == NS_A && rn == QLatin1String("blip")) {
                    QString embed = r.attributes().value(QLatin1String("r:embed")).toString();
                    if (!embed.isEmpty() && imagesByRelId.contains(embed)) {
                        f.imageData = imagesByRelId[embed].data;
                        f.imageFileName = imagesByRelId[embed].target;
                    }
                }
                if (rns == NS_W && rn == QLatin1String("txbxContent")) {
                    QVector<RunFragment> txbx = collectTextboxParagraphs(r, imagesByRelId, hyperlinkTargets);
                    fragments.append(txbx);
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
                                 QTextDocument *doc)
{
    static quint64 s_embeddedImageSeq = 0;
    for (const RunFragment &rf : runs) {
        if (rf.isLineBreak) {
            // U+2028 is a within-block line separator; "\n" would split the
            // paragraph and corrupt lists.
            cursor.insertText(QString(QChar::LineSeparator), rf.charFmt);
        } else if (rf.isTab) {
            cursor.insertText(QStringLiteral("\t"), rf.charFmt);
        } else if (rf.isImage) {
            QImage img;
            if (img.loadFromData(rf.imageData)) {
                QString ext = QFileInfo(rf.imageFileName).suffix().toLower();
                QString name = QStringLiteral("docx_embed_%1.%2").arg(++s_embeddedImageSeq)
                    .arg(ext.isEmpty() ? QStringLiteral("png") : ext);
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
    int numId = -1;
    int ilvl = 0;
    QTextBlockFormat blockFmt;
    Qt::Alignment align = Qt::AlignLeft;
    QVector<RunFragment> runs;
};

static QString serializeElementXml(QXmlStreamReader &r)
{
    QString out;
    QXmlStreamWriter pw(&out);
    pw.setAutoFormatting(false);
    NsTracker ns;
    pw.writeStartElement(r.qualifiedName().toString());
    declareElementNamespaces(pw, r, ns);
    ns.declare(pw, QStringLiteral("w"), NS_W);
    ns.declare(pw, QStringLiteral("wp"), NS_WP);
    ns.declare(pw, QStringLiteral("a"), NS_A);
    ns.declare(pw, QStringLiteral("r"), NS_R);
    for (const auto &attr : r.attributes())
        pw.writeAttribute(attr.qualifiedName().toString(), attr.value().toString());

    int depth = 1;
    while (depth > 0 && r.readNext() != QXmlStreamReader::Invalid) {
        if (r.isStartElement()) {
            ns.pushScope();
            copyStartElement(pw, r, ns);
            ++depth;
        } else if (r.isEndElement()) {
            pw.writeEndElement();
            ns.popScope();
            --depth;
        } else if (r.isCharacters()) {
            pw.writeCharacters(r.text().toString());
        } else if (r.isCDATA()) {
            pw.writeCDATA(r.text().toString());
        } else if (r.isComment()) {
            pw.writeComment(r.text().toString());
        }
    }
    pw.writeEndDocument();
    return out;
}

static ParaBuffer parseParagraph(QXmlStreamReader &r,
                                 const QMap<QString, ImageRel> &imagesByRelId,
                                 const QMap<QString, QString> &hyperlinkTargets = {})
{
    ParaBuffer pb;
    const QString xml = serializeElementXml(r);

    QXmlStreamReader px(xml);
    while (!px.atEnd()) {
        px.readNext();
        if (px.isStartElement() && px.name().toString() == QLatin1String("p"))
            break;
    }
    while (!px.atEnd()) {
        px.readNext();
        if (px.isEndElement() && px.name().toString() == QLatin1String("p"))
            break;
        if (!px.isStartElement() || px.namespaceUri().toString() != NS_W)
            continue;
        QString pn = px.name().toString();
        if (pn == QLatin1String("pPr")) {
            PPrResult ppr = readPPr(px);
            pb.numId = ppr.numId;
            pb.ilvl = ppr.ilvl;
            pb.blockFmt = ppr.blockFmt;
            pb.align = ppr.align;
            if (!ppr.pStyle.isEmpty())
                pb.blockFmt.setProperty(QTextFormat::UserProperty, ppr.pStyle);
        } else if (pn == QLatin1String("r")) {
            pb.runs = collectRuns(px, imagesByRelId, hyperlinkTargets);
        } else if (pn == QLatin1String("hyperlink")) {
            QString relId = px.attributes().value(QLatin1String("r:id")).toString();
            pb.runs = collectRuns(px, imagesByRelId, hyperlinkTargets,
                                  hyperlinkTargets.value(relId));
        }
    }
    if (px.hasError())
        qWarning() << "Failed to re-parse paragraph XML:" << px.errorString();
    return pb;
}

// Reads the paragraphs inside a <w:txbxContent> textbox and returns them as a
// run stream. The textbox is a distinct visual box, so its paragraphs are
// flattened into the enclosing body paragraph separated by line breaks.
static QVector<RunFragment> collectTextboxParagraphs(QXmlStreamReader &r,
                                                     const QMap<QString, ImageRel> &imagesByRelId,
                                                     const QMap<QString, QString> &hyperlinkTargets)
{
    QVector<RunFragment> fragments;
    while (r.readNext() != QXmlStreamReader::Invalid) {
        if (r.isEndElement() && r.namespaceUri().toString() == NS_W
            && r.name().toString() == QLatin1String("txbxContent"))
            break;
        if (!r.isStartElement())
            continue;
        if (r.namespaceUri().toString() == NS_W && r.name().toString() == QLatin1String("p")) {
            RunFragment br;
            br.isLineBreak = true;
            fragments.append(br);
            ParaBuffer pb = parseParagraph(r, imagesByRelId, hyperlinkTargets);
            for (const RunFragment &rf : pb.runs)
                fragments.append(rf);
        } else {
            skipElement(r);
        }
    }
    return fragments;
}

static void parseTable(QXmlStreamReader &xml, QTextCursor &cursor,
                       const QMap<QString, ImageRel> &imagesByRelId,
                       const QMap<QString, QString> &hyperlinkTargets,
                       QTextDocument *doc,
                       QMap<QPair<int,int>, QTextList*> &activeLists,
                       const QMap<int, NumDef> &numbering)
{
    QString tblXml;
    {
        QXmlStreamWriter pw(&tblXml);
        pw.setAutoFormatting(false);
        NsTracker ns;
        pw.writeStartElement(QLatin1String("w:tbl"));
        declareElementNamespaces(pw, xml, ns);
        ns.declare(pw, QStringLiteral("w"), NS_W);
        ns.declare(pw, QStringLiteral("wp"), NS_WP);
        ns.declare(pw, QStringLiteral("a"), NS_A);
        ns.declare(pw, QStringLiteral("r"), NS_R);
        for (const auto &attr : xml.attributes())
            pw.writeAttribute(attr.qualifiedName().toString(), attr.value().toString());

        int depth = 1;
        while (depth > 0 && xml.readNext() != QXmlStreamReader::Invalid) {
            if (xml.isStartElement()) {
                ns.pushScope();
                copyStartElement(pw, xml, ns);
                ++depth;
            } else if (xml.isEndElement()) {
                pw.writeEndElement();
                ns.popScope();
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
        int colSpan = 1;
        int gridCol = 0;
        bool vMergeRestart = false;
        bool vMergeContinue = false;
    };
    QVector<QVector<CellContent>> grid;
    int currentRow = -1;
    int rowColCursor = 0;

    QXmlStreamReader tx(tblXml);
    while (!tx.atEnd()) {
        tx.readNext();
        if (!tx.isStartElement()) continue;
        if (tx.name().toString() == QLatin1String("tr")) {
            grid.append(QVector<CellContent>());
            ++currentRow;
            rowColCursor = 0;
        } else if (tx.name().toString() == QLatin1String("tc") && currentRow >= 0) {
            CellContent cell;
            while (!tx.atEnd()) {
                tx.readNext();
                if (tx.isEndElement() && tx.name().toString() == QLatin1String("tc"))
                    break;
                if (!tx.isStartElement()) continue;
                if (tx.name().toString() == QLatin1String("tcPr")) {
                    while (tx.readNext() != QXmlStreamReader::Invalid) {
                        if (tx.isEndElement() && tx.name().toString() == QLatin1String("tcPr"))
                            break;
                        if (!tx.isStartElement()) continue;
                        if (tx.name().toString() == QLatin1String("gridSpan")) {
                            QString v = tx.attributes().value(QLatin1String("w:val")).toString();
                            if (!v.isEmpty()) cell.colSpan = qMax(1, v.toInt());
                        } else if (tx.name().toString() == QLatin1String("vMerge")) {
                            QString v = tx.attributes().value(QLatin1String("w:val")).toString();
                            if (v == QLatin1String("continue") || v.isEmpty())
                                cell.vMergeContinue = true;
                            else
                                cell.vMergeRestart = true;
                        } else {
                            skipElement(tx);
                        }
                    }
                } else if (tx.name().toString() == QLatin1String("p")) {
                    cell.paragraphs.append(parseParagraph(tx, imagesByRelId, hyperlinkTargets));
                }
            }
            if (currentRow >= 0 && currentRow < grid.size()) {
                cell.gridCol = rowColCursor;
                rowColCursor += cell.colSpan;
                grid[currentRow].append(cell);
            }
        }
    }

    if (tx.hasError())
        qWarning() << "Failed to re-parse table XML:" << tx.errorString();

    if (grid.isEmpty()) return;

    int rows = grid.size();
    int cols = 0;
    for (const auto &row : grid)
        cols = qMax(cols, row.size());
    if (cols == 0) return;

    // Cap table dimensions so a crafted document can't force a huge
    // allocation (mirrors XamlConverter's clamping).
    constexpr int kMaxTableRows = 500;
    constexpr int kMaxTableCols = 100;
    rows = qBound(1, rows, kMaxTableRows);
    cols = qBound(1, cols, kMaxTableCols);

    QTextTable *table = cursor.insertTable(rows, cols);

    // Collect vertical-merge ranges (per column) first, then apply each merge
    // exactly once. Merging incrementally as rows are scanned would re-merge
    // ranges that QTextTable has already folded together.
    QVector<QVector<QPair<int, int>>> mergeRanges(cols);
    QVector<int> mergeStartRow(cols, -1);
    for (int r = 0; r < rows; ++r) {
        for (int c = 0; c < grid[r].size(); ++c) {
            const CellContent &cc = grid[r][c];
            const int gc = cc.gridCol;
            if (gc >= cols)
                continue; // cell lies beyond the clamped column count
            if (cc.vMergeContinue) {
                if (mergeStartRow[gc] < 0)
                    mergeStartRow[gc] = r;
            } else if (cc.vMergeRestart) {
                if (mergeStartRow[gc] >= 0 && mergeStartRow[gc] < r)
                    mergeRanges[gc].append(qMakePair(mergeStartRow[gc], r - 1));
                mergeStartRow[gc] = r;
            } else {
                if (mergeStartRow[gc] >= 0 && mergeStartRow[gc] < r)
                    mergeRanges[gc].append(qMakePair(mergeStartRow[gc], r - 1));
                mergeStartRow[gc] = -1;
            }
        }
    }
    for (int c = 0; c < cols; ++c) {
        if (mergeStartRow[c] >= 0 && mergeStartRow[c] < rows)
            mergeRanges[c].append(qMakePair(mergeStartRow[c], rows - 1));
    }
    // Fill every cell (including vMerge continuation cells) BEFORE applying
    // merges: QTextTable::mergeCells moves the content of the merged cells
    // into the target cell, so text in continuation cells is preserved.
    for (int r = 0; r < rows; ++r) {
        for (int i = 0; i < grid[r].size(); ++i) {
            const CellContent &cc = grid[r][i];
            if (cc.gridCol >= cols)
                continue; // cell lies beyond the clamped column count
            QTextTableCell cell = table->cellAt(r, cc.gridCol);
            QTextCursor cellCursor = cell.firstCursorPosition();
            cellCursor.movePosition(QTextCursor::NextBlock, QTextCursor::KeepAnchor);
            cellCursor.removeSelectedText();

            bool firstPara = true;
            for (const ParaBuffer &pb : cc.paragraphs) {
                if (firstPara) {
                    cellCursor.setBlockFormat(pb.blockFmt);
                    firstPara = false;
                } else {
                    cellCursor.insertBlock(pb.blockFmt);
                }

                QPair<int,int> listKey(pb.numId, pb.ilvl);
                if (pb.numId >= 0 && numbering.contains(pb.numId)) {
                    const NumDef &nd = numbering[pb.numId];
                    if (nd.levels.isEmpty())
                        continue;
                    int level = qBound(0, pb.ilvl, nd.levels.size() - 1);
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

                insertRunsIntoCursor(cellCursor, pb.runs, doc);
            }
        }
    }

    for (int c = 0; c < cols; ++c) {
        for (const QPair<int, int> &range : mergeRanges[c]) {
            if (range.second > range.first)
                table->mergeCells(range.first, c, range.second - range.first + 1, 1);
        }
    }

    for (int r = 0; r < rows; ++r) {
        for (int i = 0; i < grid[r].size(); ++i) {
            const CellContent &cc = grid[r][i];
            if (cc.gridCol >= cols)
                continue; // cell lies beyond the clamped column count
            if (cc.colSpan > 1) {
                int span = qMin(cc.colSpan, cols - cc.gridCol);
                if (span > 1)
                    table->mergeCells(r, cc.gridCol, 1, span);
            }
        }
    }
}

bool DocxConverter::loadFromDocx(const QByteArray &zipData, QTextDocument *doc,
                                 QMarginsF &outMargins, QColor &outPageBackground)
{
    if (zipData.isEmpty()) {
        qWarning() << "DOCX data is empty";
        return false;
    }
    
    mz_zip_archive zip;
    mz_zip_zero_struct(&zip);
    if (!mz_zip_reader_init_mem(&zip, zipData.constData(), zipData.size(), 0)) {
        qDebug() << "ZIP error:" << mz_zip_get_last_error(&zip);
        return false;
    }

    // Bound decompression: a tiny .docx can declare huge uncompressed sizes
    // (zip bomb) or exceed INT_MAX, which would truncate the size_t cast below.
    constexpr mz_uint64 kMaxEntrySize = 256ull * 1024 * 1024;
    constexpr mz_uint64 kMaxTotalSize = 512ull * 1024 * 1024;
    mz_uint64 totalSize = 0;
    QMap<QString, QByteArray> files;
    const mz_uint numFiles = mz_zip_reader_get_num_files(&zip);
    for (mz_uint i = 0; i < numFiles; ++i) {
        mz_zip_archive_file_stat st;
        if (!mz_zip_reader_file_stat(&zip, i, &st))
            continue;
        if (st.m_is_directory)
            continue;
        if (st.m_uncomp_size > kMaxEntrySize
            || st.m_uncomp_size > kMaxTotalSize - totalSize) {
            qWarning() << "Skipping oversized ZIP entry:" << st.m_filename;
            continue;
        }
        if (st.m_uncomp_size > static_cast<mz_uint64>(INT_MAX)) {
            qWarning() << "Skipping ZIP entry too large for memory:" << st.m_filename;
            continue;
        }
        const mz_uint size = static_cast<mz_uint>(st.m_uncomp_size);
        QByteArray data;
        try {
            data.resize(static_cast<int>(size));
        } catch (const std::bad_alloc &) {
            qWarning() << "Out of memory extracting ZIP entry:" << st.m_filename;
            mz_zip_reader_end(&zip);
            return false;
        }
        if (!mz_zip_reader_extract_to_mem(&zip, i, data.data(), size, 0)) {
            mz_zip_reader_end(&zip);
            return false;
        }
        totalSize += st.m_uncomp_size;
        files[QString::fromUtf8(st.m_filename)] = data;
    }
    mz_zip_reader_end(&zip);

    if (!files.contains(QLatin1String("word/document.xml")))
        return false;

    QMap<int, NumDef> numbering;
    if (files.contains(QLatin1String("word/numbering.xml")))
        numbering = parseNumbering(files[QLatin1String("word/numbering.xml")]);

    QMap<QString, ImageRel> imageDataByRelId;
    QMap<QString, QString> hyperlinkTargets;
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
                if (type.contains(QLatin1String("image"))) {
                    // Keep the media rel target inside word/ — a "../word/document.xml"
                    // target must not alias another entry as an "image".
                    const QString mediaPath =
                        QDir::cleanPath(QLatin1String("word/") + target);
                    if (mediaPath.startsWith(QStringLiteral("word/")) && files.contains(mediaPath)) {
                        ImageRel ir;
                        ir.data = files[mediaPath];
                        ir.target = target;
                        imageDataByRelId[id] = ir;
                    }
                } else if (type.contains(QLatin1String("hyperlink"))) {
                    // Only allow safe external schemes; a crafted document
                    // could otherwise smuggle javascript:/file: into anchors.
                    const QString scheme = QUrl(target).scheme().toLower();
                    static const QStringList allowed = {
                        QStringLiteral("http"), QStringLiteral("https"),
                        QStringLiteral("ftp"), QStringLiteral("ftps"),
                        QStringLiteral("mailto")
                    };
                    if (allowed.contains(scheme))
                        hyperlinkTargets[id] = target;
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

    // Fail-safe: validate document.xml before mutating the target document so a
    // malformed file cannot leave partial content behind on failure.
    {
        QXmlStreamReader validator(docXmlRaw);
        while (!validator.atEnd())
            validator.readNext();
        if (validator.hasError()) {
            qDebug() << "XML validation error in word/document.xml at line" << validator.lineNumber() 
                     << "column" << validator.columnNumber() << ":" << validator.errorString();
            return false;
        }
    }

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
            ParaBuffer pb = parseParagraph(xml, imageDataByRelId, hyperlinkTargets);

            if (firstBlock) {
                cursor.setBlockFormat(pb.blockFmt);
                firstBlock = false;
            } else {
                cursor.insertBlock(pb.blockFmt);
            }

            if (cursor.blockFormat().alignment() != pb.align) {
                QTextBlockFormat bf = cursor.blockFormat();
                bf.setAlignment(pb.align);
                cursor.setBlockFormat(bf);
            }

            QPair<int,int> listKey(pb.numId, pb.ilvl);
            if (pb.numId >= 0 && numbering.contains(pb.numId)) {
                const NumDef &nd = numbering[pb.numId];
                if (nd.levels.isEmpty())
                    continue;
                int level = qBound(0, pb.ilvl, nd.levels.size() - 1);
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

            insertRunsIntoCursor(cursor, pb.runs, doc);
            continue;
        }

        if (n == QLatin1String("tbl")) {
            parseTable(xml, cursor, imageDataByRelId, hyperlinkTargets, doc,
                       activeLists, numbering);
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
                    if (!w.isEmpty()) pgW = twipToPx(w.toInt());
                    if (!h.isEmpty()) pgH = twipToPx(h.toInt());
                } else if (sn == QLatin1String("pgMar")) {
                    auto a = xml.attributes();
                    QString top = a.value(QLatin1String("w:top")).toString();
                    QString bottom = a.value(QLatin1String("w:bottom")).toString();
                    QString left = a.value(QLatin1String("w:left")).toString();
                    QString right = a.value(QLatin1String("w:right")).toString();
                    if (!top.isEmpty())    margins.setTop(twipToPx(top.toInt()));
                    if (!bottom.isEmpty()) margins.setBottom(twipToPx(bottom.toInt()));
                    if (!left.isEmpty())   margins.setLeft(twipToPx(left.toInt()));
                    if (!right.isEmpty())  margins.setRight(twipToPx(right.toInt()));
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
        doc->clear();
        return false;
    }

    doc->setPageSize(QSizeF(pgW, pgH));
    outMargins = margins;
    outPageBackground = QColor(Qt::white);

    return true;
}
