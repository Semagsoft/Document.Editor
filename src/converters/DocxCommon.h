#pragma once

#include <QString>
#include <QtNumeric>

// Namespace URIs for OOXML (OpenXML wordprocessingml)
#define NS_W  QStringLiteral("http://schemas.openxmlformats.org/wordprocessingml/2006/main")
#define NS_REL QStringLiteral("http://schemas.openxmlformats.org/package/2006/relationships")
#define NS_A  QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/main")
#define NS_WP QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
#define NS_CT QStringLiteral("http://schemas.openxmlformats.org/package/2006/content-types")
#define NS_R  QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships")

// Unit conversion constants shared by the DOCX reader and writer.
static constexpr qreal EMU_PER_PX    = 9525.0;    // 914400 EMU per inch / 96 px per inch
static constexpr qreal TWIP_PER_PX   = 15.0;      // 1440 twips per inch / 96 px per inch
static constexpr qreal HALFPT_PER_PT = 2.0;

static inline qreal emuToPx(qlonglong emu)   { return static_cast<qreal>(emu) / EMU_PER_PX; }
static inline qlonglong pxToEmu(qreal px)        { return qRound64(px * EMU_PER_PX); }
static inline qreal twipToPx(int twip)       { return static_cast<qreal>(twip) / TWIP_PER_PX; }
static inline int pxToTwip(qreal px)       { return qRound(px * TWIP_PER_PX); }
static inline qreal halfPtToPt(int half)     { return static_cast<qreal>(half) / HALFPT_PER_PT; }
static inline int ptToHalfPt(qreal pt)     { return qRound(pt * HALFPT_PER_PT); }
static inline int propToDocxLine(int pct)  { return qRound(pct * 240.0 / 100.0); }
static inline int docxLineToProportional(int line)
{
    if (line <= 0) return 100;
    return qRound(line * 100.0 / 240.0);
}
