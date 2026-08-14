#include "RtfConverter.h"

#include <QTextCursor>
#include <QTextBlockFormat>
#include <QTextCharFormat>
#include <QTextBlock>
#include <QTextList>
#include <QTextFrame>
#include <QStack>
#include <QXmlStreamWriter>
#include <QColor>
#include <QFont>

// Windows-1252 mapping for the 0x80..0x9F range (everything else matches
// Latin-1 1:1). RTF files are typically ANSI/Windows-1252, not UTF-8.
static const QChar kCp1252High[32] = {
    QChar(0x20AC), QChar(0x0081), QChar(0x201A), QChar(0x0192),
    QChar(0x201E), QChar(0x2026), QChar(0x2020), QChar(0x2021),
    QChar(0x02C6), QChar(0x2030), QChar(0x0160), QChar(0x2039),
    QChar(0x0152), QChar(0x008D), QChar(0x017D), QChar(0x008F),
    QChar(0x0090), QChar(0x2018), QChar(0x2019), QChar(0x201C),
    QChar(0x201D), QChar(0x2022), QChar(0x2013), QChar(0x2014),
    QChar(0x02DC), QChar(0x2122), QChar(0x0161), QChar(0x203A),
    QChar(0x0153), QChar(0x009D), QChar(0x017E), QChar(0x0178),
};

static QChar fromCp1252Byte(int code)
{
    if (code >= 0x80 && code < 0xA0)
        return kCp1252High[code - 0x80];
    return QChar(static_cast<ushort>(code));
}

static QString decodeWindows1252(const QByteArray &data)
{
    QString out;
    out.reserve(data.size());
    for (char byte : data) {
        const uchar b = static_cast<uchar>(byte);
        out += fromCp1252Byte(b);
    }
    return out;
}

struct RtfFormatState {
    bool bold = false;
    bool italic = false;
    bool underline = false;
    bool strikethrough = false;
    bool subscript = false;
    bool superscript = false;
    int fontSizeHalfPoints = 0;
    QString fontFamily;
    QColor foreground;
    QColor background;
    Qt::Alignment alignment = Qt::AlignLeft;
    int firstLineIndent = 0;
    int leftIndent = 0;
    int rightIndent = 0;

    void reset()
    {
        bold = false;
        italic = false;
        underline = false;
        strikethrough = false;
        subscript = false;
        superscript = false;
        fontSizeHalfPoints = 0;
        fontFamily.clear();
        foreground = QColor();
        background = QColor();
        alignment = Qt::AlignLeft;
        firstLineIndent = 0;
        leftIndent = 0;
        rightIndent = 0;
    }
};

static void applyCharFormat(QTextCursor &cursor, const RtfFormatState &state)
{
    QTextCharFormat fmt;
    if (state.bold)
        fmt.setFontWeight(QFont::Bold);
    if (state.italic)
        fmt.setFontItalic(true);
    if (state.underline)
        fmt.setFontUnderline(true);
    if (state.strikethrough)
        fmt.setFontStrikeOut(true);
    if (state.subscript)
        fmt.setVerticalAlignment(QTextCharFormat::AlignSubScript);
    if (state.superscript)
        fmt.setVerticalAlignment(QTextCharFormat::AlignSuperScript);
    if (state.fontSizeHalfPoints > 0)
        fmt.setFontPointSize(state.fontSizeHalfPoints / 2.0);
    if (!state.fontFamily.isEmpty())
        fmt.setFontFamilies({state.fontFamily});
    if (state.foreground.isValid())
        fmt.setForeground(state.foreground);
    if (state.background.isValid())
        fmt.setBackground(state.background);
    cursor.setCharFormat(fmt);
}

static void applyBlockFormat(QTextCursor &cursor, const RtfFormatState &state)
{
    QTextBlockFormat fmt;
    fmt.setAlignment(state.alignment);
    if (state.firstLineIndent != 0)
        fmt.setTextIndent(state.firstLineIndent / 15.0);
    if (state.leftIndent > 0)
        fmt.setLeftMargin(state.leftIndent / 15.0);
    if (state.rightIndent > 0)
        fmt.setRightMargin(state.rightIndent / 15.0);
    cursor.setBlockFormat(fmt);
}

static int parseColorTable(const QString &rtf, int i, QVector<QColor> &colorTable)
{
    // Parses {\colortbl;\redN\greenN\blueN;\redN\greenN\blueN;} starting just
    // past the "colortbl" control word. Returns the index of the closing '}'.
    QColor current;
    bool haveColor = false;
    while (i < rtf.length()) {
        const QChar c = rtf[i];
        if (c == QLatin1Char('{')) {
            // Nested destination (e.g. \* ...) — skip it.
            int depth = 1;
            i++;
            while (i < rtf.length() && depth > 0) {
                if (rtf[i] == QLatin1Char('{')) depth++;
                else if (rtf[i] == QLatin1Char('}')) depth--;
                i++;
            }
        } else if (c == QLatin1Char('}')) {
            break;
        } else if (c == QLatin1Char(';')) {
            if (haveColor)
                colorTable.append(current);
            current = QColor();
            haveColor = false;
            i++;
        } else if (c == QLatin1Char('\\')) {
            i++;
            QString word;
            while (i < rtf.length() && rtf[i].isLetter()) {
                word += rtf[i];
                i++;
            }
            int arg = 0;
            bool negative = false;
            if (i < rtf.length() && rtf[i] == QLatin1Char('-')) {
                negative = true;
                i++;
            }
            while (i < rtf.length() && rtf[i].isDigit()) {
                arg = arg * 10 + (rtf[i].unicode() - QLatin1Char('0').unicode());
                i++;
            }
            if (negative)
                arg = -arg;
            if (i < rtf.length() && rtf[i] == QLatin1Char(' '))
                i++;
            if (word == QLatin1String("red")) {
                current.setRed(qBound(0, arg, 255));
                haveColor = true;
            } else if (word == QLatin1String("green")) {
                current.setGreen(qBound(0, arg, 255));
                haveColor = true;
            } else if (word == QLatin1String("blue")) {
                current.setBlue(qBound(0, arg, 255));
                haveColor = true;
            }
            // cyan/magenta/yellow/black/tint/shade/font are ignored.
        } else {
            i++;
        }
    }
    return i;
}

static bool loadRtfToDocument(const QString &rtf, QTextDocument *doc)
{
    QTextCursor cursor(doc);
    QStack<RtfFormatState> stateStack;
    RtfFormatState currentState;
    stateStack.push(currentState);

    // Color table parsing
    QVector<QColor> colorTable;
    colorTable.append(QColor(Qt::black)); // index 0 = auto

    int i = 0;
    bool inColorTable = false;
    QColor parsingColor;
    bool parsingColorTable = false;

    while (i < rtf.length()) {
        QChar c = rtf[i];

        if (c == QLatin1Char('{')) {
            stateStack.push(currentState);
            i++;
        } else if (c == QLatin1Char('}')) {
            if (stateStack.size() > 1) {
                currentState = stateStack.pop();
                applyCharFormat(cursor, currentState);
            } else {
                stateStack.top() = currentState;
            }
            i++;
        } else if (c == QLatin1Char('\\')) {
            i++;
            if (i >= rtf.length())
                break;

            QChar next = rtf[i];

            // Escaped special characters
            if (next == QLatin1Char('{') || next == QLatin1Char('}') || next == QLatin1Char('\\')) {
                cursor.insertText(QString(next));
                i++;
                continue;
            }

            // Line break
            if (next == QLatin1Char('\n') || next == QLatin1Char('\r')) {
                i++;
                continue;
            }

            // Hex-escaped byte: \'hh  (' is not a letter, so it is not caught
            // by the control-word reader below). 0xE9 -> "é", 0x93 -> " in
            // Windows-1252.
            if (next == QLatin1Char('\'')) {
                i++;
                if (i + 2 <= rtf.length()) {
                    QString hex = rtf.mid(i, 2);
                    bool ok = false;
                    int code = hex.toInt(&ok, 16);
                    if (ok) {
                        QChar ch = fromCp1252Byte(code);
                        applyCharFormat(cursor, currentState);
                        cursor.insertText(QString(ch));
                    }
                    i += 2;
                }
                continue;
            }

            // Control word
            QString controlWord;
            bool hasNumeric = false;
            int numericArg = 0;
            bool negative = false;

            while (i < rtf.length() && rtf[i].isLetter()) {
                controlWord += rtf[i];
                i++;
            }

            if (i < rtf.length() && rtf[i] == QLatin1Char('-')) {
                negative = true;
                i++;
            }

            while (i < rtf.length() && rtf[i].isDigit()) {
                hasNumeric = true;
                numericArg = numericArg * 10 + (rtf[i].unicode() - QLatin1Char('0').unicode());
                i++;
            }
            if (negative)
                numericArg = -numericArg;

            // Skip space delimiter after control word
            if (i < rtf.length() && rtf[i] == QLatin1Char(' '))
                i++;

            // === Character formatting ===
            if (controlWord == QStringLiteral("b")) {
                currentState.bold = !hasNumeric || numericArg != 0;
            } else if (controlWord == QStringLiteral("i")) {
                currentState.italic = !hasNumeric || numericArg != 0;
            } else if (controlWord == QStringLiteral("ul") || controlWord == QStringLiteral("uld")) {
                currentState.underline = !hasNumeric || numericArg != 0;
            } else if (controlWord == QStringLiteral("ulnone") || controlWord == QStringLiteral("ul0")) {
                currentState.underline = false;
            } else if (controlWord == QStringLiteral("strike") || controlWord == QStringLiteral("striked")) {
                currentState.strikethrough = !hasNumeric || numericArg != 0;
            } else if (controlWord == QStringLiteral("strike0")) {
                currentState.strikethrough = false;
            } else if (controlWord == QStringLiteral("sub")) {
                currentState.subscript = true;
                currentState.superscript = false;
            } else if (controlWord == QStringLiteral("super")) {
                currentState.superscript = true;
                currentState.subscript = false;
            } else if (controlWord == QStringLiteral("nosupersub")) {
                currentState.subscript = false;
                currentState.superscript = false;
            } else if (controlWord == QStringLiteral("fs") && hasNumeric) {
                currentState.fontSizeHalfPoints = numericArg;
            } else if (controlWord == QStringLiteral("f") && hasNumeric) {
                // Font index - requires font table parsing; use default for now
            } else if (controlWord == QStringLiteral("cf") && hasNumeric) {
                if (numericArg >= 0 && numericArg < colorTable.size())
                    currentState.foreground = colorTable[numericArg];
            } else if (controlWord == QStringLiteral("cb") && hasNumeric) {
                if (numericArg >= 0 && numericArg < colorTable.size())
                    currentState.background = colorTable[numericArg];
            } else if (controlWord == QStringLiteral("highlight") && hasNumeric) {
                if (numericArg >= 0 && numericArg < colorTable.size())
                    currentState.background = colorTable[numericArg];
            } else if (controlWord == QStringLiteral("plain")) {
                currentState.reset();
            }

            // === Paragraph formatting ===
            else if (controlWord == QStringLiteral("par")) {
                applyBlockFormat(cursor, currentState);
                cursor.insertBlock();
            } else if (controlWord == QStringLiteral("pard")) {
                currentState.alignment = Qt::AlignLeft;
                currentState.firstLineIndent = 0;
                currentState.leftIndent = 0;
                currentState.rightIndent = 0;
            } else if (controlWord == QStringLiteral("ql")) {
                currentState.alignment = Qt::AlignLeft;
            } else if (controlWord == QStringLiteral("qc")) {
                currentState.alignment = Qt::AlignCenter;
            } else if (controlWord == QStringLiteral("qr")) {
                currentState.alignment = Qt::AlignRight;
            } else if (controlWord == QStringLiteral("qj")) {
                currentState.alignment = Qt::AlignJustify;
            } else if (controlWord == QStringLiteral("fi") && hasNumeric) {
                currentState.firstLineIndent = numericArg;
            } else if (controlWord == QStringLiteral("li") && hasNumeric) {
                currentState.leftIndent = numericArg;
            } else if (controlWord == QStringLiteral("ri") && hasNumeric) {
                currentState.rightIndent = numericArg;
            }

            // === Special characters ===
            else if (controlWord == QStringLiteral("tab")) {
                cursor.insertText(QStringLiteral("\t"));
            } else if (controlWord == QStringLiteral("line")) {
                cursor.insertText(QString(QChar::LineSeparator));
            } else if (controlWord == QStringLiteral("cell")) {
                cursor.insertText(QStringLiteral("\t"));
            } else if (controlWord == QStringLiteral("row")) {
                cursor.insertBlock();
            } else if (controlWord == QStringLiteral("u") && hasNumeric) {
                QChar ch(static_cast<ushort>(numericArg));
                applyCharFormat(cursor, currentState);
                cursor.insertText(QString(ch));
                // Skip the replacement character that follows \uN
                if (i < rtf.length())
                    i++;
            }

            // === Tables and groups ===
            else if (controlWord == QStringLiteral("colortbl")) {
                i = parseColorTable(rtf, i, colorTable);
            } else if (controlWord == QStringLiteral("fonttbl")
                     || controlWord == QStringLiteral("stylesheet")
                     || controlWord == QStringLiteral("header")
                     || controlWord == QStringLiteral("footer")
                     || controlWord == QStringLiteral("headerf")
                     || controlWord == QStringLiteral("footerf")
                     || controlWord == QStringLiteral("headerl")
                     || controlWord == QStringLiteral("footerl")
                     || controlWord == QStringLiteral("headerr")
                     || controlWord == QStringLiteral("footerr")
                     || controlWord == QStringLiteral("pict")) {
                // Skip these groups - fast-forward to matching close brace
                int depth = 1;
                while (i < rtf.length() && depth > 0) {
                    if (rtf[i] == QLatin1Char('{')) depth++;
                    else if (rtf[i] == QLatin1Char('}')) depth--;
                    i++;
                }
            } else if (controlWord == QStringLiteral("*")) {
                // \* - skip this destination group
                int depth = 1;
                while (i < rtf.length() && depth > 0) {
                    if (rtf[i] == QLatin1Char('{')) depth++;
                    else if (rtf[i] == QLatin1Char('}')) depth--;
                    i++;
                }
            }

            applyCharFormat(cursor, currentState);

        } else if (c == QLatin1Char('\r') || c == QLatin1Char('\n')) {
            i++;
        } else {
            applyCharFormat(cursor, currentState);
            cursor.insertText(QString(c));
            i++;
        }
    }

    return true;
}

static QString escapeRtfText(const QString &text)
{
    QString result;
    for (const QChar &c : text) {
        if (c == QLatin1Char('\\'))
            result += QStringLiteral("\\\\");
        else if (c == QLatin1Char('{'))
            result += QStringLiteral("\\{");
        else if (c == QLatin1Char('}'))
            result += QStringLiteral("\\}");
        else if (c == QLatin1Char('\n'))
            result += QStringLiteral("\\par ");
        else if (c == QChar::LineSeparator)
            result += QStringLiteral("\\line ");
        else if (c == QLatin1Char('\t'))
            result += QStringLiteral("\\tab ");
        else if (c.unicode() > 127)
            result += QStringLiteral("\\u%1?").arg(static_cast<int>(c.unicode()));
        else
            result += c;
    }
    return result;
}

bool RtfConverter::loadFromRtf(const QByteArray &rtfData, QTextDocument *doc)
{
    // RTF is ANSI/Windows-1252 (or another code page) — not UTF-8. Decoding as
    // UTF-8 corrupts every non-ASCII byte, breaking literal text and \'hh
    // escapes that the parser below relies on.
    QString text = decodeWindows1252(rtfData);

    // Verify it looks like RTF
    if (!text.startsWith(QLatin1Char('{')) && !text.contains(QStringLiteral("\\rtf")))
        return false;

    doc->clear();
    return loadRtfToDocument(text, doc);
}

QByteArray RtfConverter::saveToRtf(const QTextDocument *doc)
{
    QString rtf;
    rtf += QStringLiteral("{\\rtf1\\ansi\\deff0\n");

    // Font table
    rtf += QStringLiteral("{\\fonttbl{\\f0\\fnil\\fcharset0 ");
    rtf += escapeRtfText(doc->defaultFont().family());
    rtf += QStringLiteral(";}}\n");

    // Color table - just black and white for minimal output
    rtf += QStringLiteral("{\\colortbl;\\red0\\green0\\blue0;\\red255\\green255\\blue255;}\n");

    // Defaults
    rtf += QStringLiteral("\\viewkind4\\uc1\\pard\\lang1033\\f0\\fs");
    rtf += QString::number(qRound(doc->defaultFont().pointSizeF() * 2));
    rtf += QStringLiteral("\n");

    QTextBlock block = doc->begin();
    while (block.isValid()) {
        QTextBlockFormat blockFmt = block.blockFormat();

        // Paragraph alignment
        if (blockFmt.alignment() == Qt::AlignCenter)
            rtf += QStringLiteral("\\qc ");
        else if (blockFmt.alignment() == Qt::AlignRight)
            rtf += QStringLiteral("\\qr ");
        else if (blockFmt.alignment() == Qt::AlignJustify)
            rtf += QStringLiteral("\\qj ");

        // Indent
        if (blockFmt.textIndent() != 0)
            rtf += QStringLiteral("\\fi%1 ").arg(qRound(blockFmt.textIndent() * 15));
        if (blockFmt.leftMargin() > 0)
            rtf += QStringLiteral("\\li%1 ").arg(static_cast<int>(blockFmt.leftMargin() * 15));
        if (blockFmt.rightMargin() > 0)
            rtf += QStringLiteral("\\ri%1 ").arg(static_cast<int>(blockFmt.rightMargin() * 15));

        for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
            QTextFragment fragment = it.fragment();
            if (!fragment.isValid())
                continue;

            QTextCharFormat charFmt = fragment.charFormat();
            QString text = fragment.text();

            // Open formatting
            bool bold = charFmt.fontWeight() >= QFont::Bold;
            bool italic = charFmt.fontItalic();
            bool underline = charFmt.fontUnderline();

            if (bold) rtf += QStringLiteral("\\b ");
            if (italic) rtf += QStringLiteral("\\i ");
            if (underline) rtf += QStringLiteral("\\ul ");

            if (charFmt.fontPointSize() > 0)
                rtf += QStringLiteral("\\fs%1 ").arg(qRound(charFmt.fontPointSize() * 2));

            rtf += escapeRtfText(text);

            // Close formatting
            if (bold) rtf += QStringLiteral("\\b0 ");
            if (italic) rtf += QStringLiteral("\\i0 ");
            if (underline) rtf += QStringLiteral("\\ulnone ");
        }

        rtf += QStringLiteral("\\par\n");
        block = block.next();
    }

    rtf += QStringLiteral("}\n");
    return rtf.toUtf8();
}
