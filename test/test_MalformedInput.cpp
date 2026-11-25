// Malformed-input / fuzz-style tests for the document converters.
//
// Feeds adversarial, truncated, and malformed inputs to each converter and
// verifies two things:
//   1. The loader never crashes (no segfault, no abort, no uncaught exception).
//   2. The loader returns false gracefully instead of producing partial state.
//
// These tests exercise the security surface: DOCX parses untrusted ZIP + OOXML,
// XAML parses untrusted XML, RTF parses untrusted control-word streams.
#include <QTest>
#include <QTextDocument>
#include <QTextCursor>
#include <QTextBlock>
#include <QTextCharFormat>
#include <QTextFrame>
#include <QTextTable>
#include <QMarginsF>
#include <QColor>
#include <QByteArray>
#include <QXmlStreamWriter>
#include <QMap>

#include "converters/DocxConverter.h"
#include "converters/XamlConverter.h"
#include "converters/RtfConverter.h"

#include "miniz.h"

// ---------------------------------------------------------------------------
// ZIP helpers (miniz is linked publicly via de_objects)
// ---------------------------------------------------------------------------

static size_t zipWriteCallback(void *pOpaque, mz_uint64 /*fileOfs*/,
                               const void *pBuf, size_t n)
{
    QByteArray *buffer = static_cast<QByteArray *>(pOpaque);
    buffer->append(static_cast<const char *>(pBuf), static_cast<int>(n));
    return n;
}

// Build an in-memory ZIP containing the given files. Returns an empty
// QByteArray if any step fails (callers should check !result.isEmpty()).
static QByteArray makeZip(const QMap<QString, QByteArray> &files)
{
    mz_zip_archive zip;
    mz_zip_zero_struct(&zip);
    QByteArray buffer;
    zip.m_pWrite = zipWriteCallback;
    zip.m_pIO_opaque = &buffer;
    if (!mz_zip_writer_init_v2(&zip, 0, 0))
        return {};
    for (auto it = files.cbegin(); it != files.cend(); ++it) {
        if (!mz_zip_writer_add_mem(&zip, it.key().toUtf8().constData(),
                                   it.value().constData(), it.value().size(),
                                   MZ_BEST_COMPRESSION))
            return {};
    }
    if (!mz_zip_writer_finalize_archive(&zip))
        return {};
    mz_zip_writer_end(&zip);
    return buffer;
}

// A minimal, well-formed word/document.xml with a single paragraph.
static QByteArray minimalDocXml()
{
    return QStringLiteral(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\n"
        "  <w:body>\n"
        "    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>\n"
        "  </w:body>\n"
        "</w:document>\n").toUtf8();
}

// A minimal _rels/document.xml.rels declaring one image relationship.
static QByteArray relsXml()
{
    return QStringLiteral(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
        "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\"/>\n"
        "  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com\"/>\n"
        "</Relationships>\n").toUtf8();
}

// A minimal [Content_Types].xml.
static QByteArray contentTypesXml()
{
    return QStringLiteral(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
        "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
        "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
        "  <Default Extension=\"png\" ContentType=\"image/png\"/>\n"
        "</Types>\n").toUtf8();
}

// A minimal valid DOCX (document.xml + content types + rels + a tiny PNG image).
static QByteArray validDocx()
{
    QMap<QString, QByteArray> files;
    files["word/document.xml"] = minimalDocXml();
    files["word/_rels/document.xml.rels"] = relsXml();
    files["[Content_Types].xml"] = contentTypesXml();
    // Tiny valid 1x1 PNG.
    files["word/media/image1.png"] = QByteArray::fromHex(
        "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
        "0000000c4944415408d7c3f9cf000000020001e221bc350000000049454e44ae42"
        "6082");
    return makeZip(files);
}

class TestMalformedInput : public QObject
{
    Q_OBJECT

private:
    static bool loadDocx(const QByteArray &data)
    {
        QTextDocument doc;
        QMarginsF margins;
        QColor bg;
        return DocxConverter::loadFromDocx(data, &doc, margins, bg);
    }

    static bool loadXaml(const QString &xml)
    {
        QTextDocument doc;
        QMarginsF margins;
        QColor bg;
        return XamlConverter::loadFromXaml(xml, &doc, margins, bg);
    }

    static bool loadRtf(const QByteArray &rtf)
    {
        QTextDocument doc;
        return RtfConverter::loadFromRtf(rtf, &doc);
    }

private slots:
    // =====================================================================
    // DOCX: empty / non-ZIP / truncated
    // =====================================================================
    void docxEmptyInput()
    {
        QVERIFY(!loadDocx(QByteArray()));
    }

    void docxNotAZip()
    {
        QVERIFY(!loadDocx(QByteArrayLiteral("not a zip file")));
        QVERIFY(!loadDocx(QByteArrayLiteral("PK\x03\x04"))); // zip signature, nothing else
        QVERIFY(!loadDocx(QByteArray::fromHex("504b0304140000000000")));
    }

    void docxTruncatedZip()
    {
        // A valid ZIP, then cut off mid-stream.
        QByteArray zip = validDocx();
        QVERIFY(!zip.isEmpty());
        QVERIFY(!loadDocx(zip.left(zip.size() / 2)));
        QVERIFY(!loadDocx(zip.left(10)));
    }

    // =====================================================================
    // DOCX: valid ZIP, missing or unreadable document.xml
    // =====================================================================
    void docxZipWithoutDocumentXml()
    {
        QMap<QString, QByteArray> files;
        files["word/other.xml"] = minimalDocXml();
        files["[Content_Types].xml"] = contentTypesXml();
        QVERIFY(!loadDocx(makeZip(files)));
    }

    void docxZipWithInvalidXml()
    {
        QMap<QString, QByteArray> files;
        files["word/document.xml"] = QByteArrayLiteral("<w:document><w:body><w:p><w:r><w:t>unclosed");
        files["[Content_Types].xml"] = contentTypesXml();
        QVERIFY(!loadDocx(makeZip(files)));
    }

    void docxZipWithMalformedXml()
    {
        QMap<QString, QByteArray> files;
        // Mismatched tags, illegal characters, deep nesting.
        QString xml = QStringLiteral(
            "<?xml version=\"1.0\"?>\n"
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\n"
            "<w:body>");
        for (int i = 0; i < 10000; ++i)
            xml += QStringLiteral("<w:p>");
        xml += QStringLiteral("</w:body></w:document>");
        files["word/document.xml"] = xml.toUtf8();
        files["[Content_Types].xml"] = contentTypesXml();
        QVERIFY(!loadDocx(makeZip(files)));
    }

    // =====================================================================
    // DOCX: zip bomb (declared uncompressed size far exceeds the cap)
    // =====================================================================
    void docxZipBombHugeDeclaredSize()
    {
        // Build a ZIP whose local header claims an enormous uncompressed size
        // (4 GiB) but holds almost no data. The loader must skip this entry
        // (it exceeds kMaxEntrySize) rather than allocate 4 GiB.
        QByteArray localHeader;
        localHeader.append(QByteArray::fromHex("504b0304")); // local file signature
        localHeader.append(QByteArray::fromHex("1400"));     // version needed
        localHeader.append(QByteArray::fromHex("0000"));     // general purpose bit flag
        localHeader.append(QByteArray::fromHex("0000"));     // compression method (none)
        localHeader.append(QByteArray::fromHex("5a4b184f")); // mod time/date (dummy)
        localHeader.append(QByteArray::fromHex("12345678")); // crc32 (dummy)
        localHeader.append(QByteArray::fromHex("ffffffff")); // compressed size (4 GiB)
        localHeader.append(QByteArray::fromHex("ffffffff")); // uncompressed size (4 GiB)
        localHeader.append(QByteArray::fromHex("0500"));     // filename length
        localHeader.append(QByteArray::fromHex("0000"));     // extra field length
        localHeader.append("word/big.bin");

        QByteArray centralHeader;
        centralHeader.append(QByteArray::fromHex("504b0102")); // central dir signature
        centralHeader.append(QByteArray::fromHex("3e00"));     // version made by
        centralHeader.append(localHeader.mid(4, 26));          // version..extra length
        centralHeader.append(QByteArray::fromHex("0000"));     // comment length
        centralHeader.append(QByteArray::fromHex("0000"));     // disk start
        centralHeader.append(QByteArray::fromHex("0000"));     // internal attr
        centralHeader.append(QByteArray::fromHex("0000"));     // external attr
        centralHeader.append(QByteArray::fromHex("32000000")); // local header offset (50)
        centralHeader.append("word/big.bin");

        quint32 cdSize = static_cast<quint32>(centralHeader.size());
        QByteArray endOfCentral;
        endOfCentral.append(QByteArray::fromHex("504b0506")); // end of central dir signature
        endOfCentral.append(QByteArray::fromHex("0000"));     // disk number
        endOfCentral.append(QByteArray::fromHex("0000"));     // dir disk
        endOfCentral.append(QByteArray::fromHex("0100"));     // num entries this disk
        endOfCentral.append(QByteArray::fromHex("0100"));     // num entries total
        endOfCentral.append(QByteArray::fromHex("0000"));     // dir size (4 bytes, little-endian)
        endOfCentral.append(QByteArray::fromHex("0000"));
        endOfCentral.append(QByteArray::fromHex("0000"));
        endOfCentral.append(QByteArray::fromHex("0000"));
        // Write the central directory size as little-endian uint32.
        endOfCentral[12] = static_cast<char>(cdSize & 0xff);
        endOfCentral[13] = static_cast<char>((cdSize >> 8) & 0xff);
        endOfCentral[14] = static_cast<char>((cdSize >> 16) & 0xff);
        endOfCentral[15] = static_cast<char>((cdSize >> 24) & 0xff);
        endOfCentral.append(QByteArray::fromHex("32000000")); // offset of central dir (50)
        endOfCentral.append(QByteArray::fromHex("0000"));     // comment length

        QByteArray zip = localHeader + centralHeader + endOfCentral;
        QVERIFY(!loadDocx(zip));
    }

    // =====================================================================
    // DOCX: malicious relationship targets
    // =====================================================================
    void docxJavascriptHyperlinkRejected()
    {
        // A hyperlink relationship with a javascript: URL must be refused;
        // the resulting document must not carry an anchor with that href.
        QString rels = QStringLiteral(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
            "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"javascript:alert(1)\"/>\n"
            "</Relationships>\n").toUtf8();

        QMap<QString, QByteArray> files;
        files["word/document.xml"] = QStringLiteral(
            "<?xml version=\"1.0\"?>\n"
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
            "<w:body>\n"
            "  <w:p><w:r><w:hyperlink r:id=\"rId1\"><w:t>Click</w:t></w:hyperlink></w:r></w:p>\n"
            "</w:body>\n"
            "</w:document>\n").toUtf8();
        files["word/_rels/document.xml.rels"] = rels.toUtf8();
        files["[Content_Types].xml"] = contentTypesXml();

        QTextDocument doc;
        QMarginsF margins;
        QColor bg;
        QVERIFY(DocxConverter::loadFromDocx(makeZip(files), &doc, margins, bg));

        bool foundAnchor = false;
        for (QTextBlock block = doc.begin(); block.isValid(); block = block.next()) {
            for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                QTextCharFormat cf = it.fragment().charFormat();
                if (cf.isAnchor() && cf.anchorHref().contains(QStringLiteral("javascript")))
                    foundAnchor = true;
            }
        }
        QVERIFY(!foundAnchor);
    }

    void docxPathTraversalImageRelRejected()
    {
        // An image relationship whose target escapes word/ must be refused.
        QString rels = QStringLiteral(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
            "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../../../etc/passwd\"/>\n"
            "</Relationships>\n").toUtf8();

        QMap<QString, QByteArray> files;
        files["word/document.xml"] = minimalDocXml();
        files["word/_rels/document.xml.rels"] = rels.toUtf8();
        files["[Content_Types].xml"] = contentTypesXml();

        QTextDocument doc;
        QMarginsF margins;
        QColor bg;
        QVERIFY(DocxConverter::loadFromDocx(makeZip(files), &doc, margins, bg));
        // Document loads, but the image must not have been fetched: no image
        // fragments should exist in the document.
        int imageCount = 0;
        for (QTextBlock block = doc.begin(); block.isValid(); block = block.next()) {
            for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                if (it.fragment().charFormat().isImageFormat())
                    ++imageCount;
            }
        }
        QCOMPARE(imageCount, 0);
    }

    // =====================================================================
    // DOCX: extreme table dimensions (DoS) - the hardening fix
    // =====================================================================
    void docxExtremeTableDimensions()
    {
        // A grid claiming 600 rows x 200 cols exceeds the clamping limits
        // (500 rows / 100 cols) but is small enough to parse without OOM.
        const int claimedRows = 600;
        const int claimedCols = 200;

        QString tbl = QStringLiteral(
            "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>"
            "<w:tblGrid>");
        for (int c = 0; c < claimedCols; ++c)
            tbl += QStringLiteral("<w:gridCol w:w=\"1000\"/>");
        tbl += QStringLiteral("</w:tblGrid>");
        QString rowsXml;
        for (int r = 0; r < claimedRows; ++r) {
            rowsXml += QStringLiteral("<w:tr>");
            for (int c = 0; c < claimedCols; ++c)
                rowsXml += QStringLiteral("<w:tc><w:tcPr/><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>");
            rowsXml += QStringLiteral("</w:tr>");
        }
        tbl += rowsXml + QStringLiteral("</w:tbl>");

        QString docXml = QStringLiteral(
            "<?xml version=\"1.0\"?>\n"
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\n"
            "<w:body>") + tbl + QStringLiteral("</w:body></w:document>");

        QMap<QString, QByteArray> files;
        files["word/document.xml"] = docXml.toUtf8();
        files["[Content_Types].xml"] = contentTypesXml();

        QTextDocument doc;
        QMarginsF margins;
        QColor bg;
        QVERIFY(DocxConverter::loadFromDocx(makeZip(files), &doc, margins, bg));

        // The table must exist but be clamped to sane bounds (500 rows, 100 cols).
        int tableCount = 0;
        int loadedRows = 0, loadedCols = 0;
        for (QTextBlock block = doc.begin(); block.isValid(); block = block.next()) {
            QTextFrame *frame = doc.frameAt(block.position());
            while (frame) {
                if (QTextTable *t = qobject_cast<QTextTable *>(frame)) {
                    ++tableCount;
                    loadedRows = qMax(loadedRows, t->rows());
                    loadedCols = qMax(loadedCols, t->columns());
                }
                frame = frame->parentFrame();
            }
        }
        QVERIFY(tableCount >= 1);
        QCOMPARE(loadedRows, 500);
        QCOMPARE(loadedCols, 100);
    }

    // =====================================================================
    // XAML: malformed XML
    // =====================================================================
    void xamlNotXml()
    {
        QVERIFY(!loadXaml(QStringLiteral("not xml")));
        QVERIFY(!loadXaml(QStringLiteral("")));
        QVERIFY(!loadXaml(QStringLiteral("<Document>unclosed")));
    }

    void xamlInvalidXml()
    {
        QVERIFY(!loadXaml(QStringLiteral("<Document><Table Rows=\"abc\" Cols=\"def\">")));
        QVERIFY(!loadXaml(QStringLiteral("<?xml version='1.0'?><broken><unmatched>")));
    }

    void xamlDeeplyNested()
    {
        QString xml = QStringLiteral("<Document>");
        for (int i = 0; i < 100000; ++i)
            xml += QStringLiteral("<Run>");
        xml += QStringLiteral("</Document>");
        // Deeply nested input must not crash or stack-overflow.
        QVERIFY(loadXaml(xml) || true); // either accepted or rejected, but no crash
    }

    void xamlExtremeTableDimensions()
    {
        // Already clamped in code; verify it still works and doesn't crash.
        QString xaml = QStringLiteral(
            "<Document><Table Rows=\"2000000000\" Cols=\"2000000000\">"
            "<TableRow><TableCell><Paragraph><Run>x</Run></Paragraph></TableCell></TableRow>"
            "</Table></Document>");
        QTextDocument doc;
        QMarginsF margins;
        QColor bg;
        QVERIFY(loadXaml(xaml));
    }

    // =====================================================================
    // RTF: malformed control streams
    // =====================================================================
    void rtfNotRtf()
    {
        QVERIFY(!loadRtf(QStringLiteral("not rtf").toUtf8()));
        QVERIFY(!loadRtf(QByteArray()));
        QVERIFY(!loadRtf(QStringLiteral("{\\rtf1").toUtf8())); // unclosed group
    }

    void rtfTruncatedBraces()
    {
        QVERIFY(!loadRtf(QStringLiteral("{\\rtf1\\ansi\\deff0 \\b unclosed").toUtf8()));
        QVERIFY(!loadRtf(QStringLiteral("{\\rtf1\\ansi\\deff0 {\\colortbl;\\red0\\green0\\blue0;}").toUtf8()));
        // Extreme \fs values are handled gracefully (covered by rtfExtremeControlValues).
    }

    void rtfDeeplyNested()
    {
        QByteArray rtf = QStringLiteral("{\\rtf1\\ansi\\deff0 ").toUtf8();
        for (int i = 0; i < 100000; ++i)
            rtf += "{";
        rtf += "x";
        for (int i = 0; i < 100000; ++i)
            rtf += "}";
        // Deeply nested RTF groups must not crash.
        QVERIFY(loadRtf(rtf) || true);
    }

    void rtfExtremeControlValues()
    {
        // Extreme \fs, \u, \red, etc. values must not crash.
        QVERIFY(loadRtf(QStringLiteral("{\\rtf1\\ansi\\deff0 \\fs999999999 text}").toUtf8()));
        QVERIFY(loadRtf(QStringLiteral("{\\rtf1\\ansi\\deff0 \\u999999999?}").toUtf8()));
        QVERIFY(loadRtf(QStringLiteral("{\\rtf1\\ansi\\deff0 {\\colortbl;\\red999999\\green999999\\blue999999;}\\cf1 text}").toUtf8()));
    }

    // =====================================================================
    // All converters: valid input still works (regression guard)
    // =====================================================================
    void validDocxStillLoads()
    {
        QVERIFY(loadDocx(validDocx()));
    }

    void validXamlStillLoads()
    {
        QVERIFY(loadXaml(QStringLiteral("<Document><Paragraph>Hello</Paragraph></Document>")));
    }

    void validRtfStillLoads()
    {
        QVERIFY(loadRtf(QByteArrayLiteral("{\\rtf1\\ansi\\deff0 Hello World}")));
    }
};

QTEST_MAIN(TestMalformedInput)
#include "test_MalformedInput.moc"
