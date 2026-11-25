// DOCX converter implementation is split across two translation units:
//
//   DocxReader.cpp  — loadFromDocx() and all OOXML parsing helpers
//   DocxWriter.cpp  — saveToDocx() and all OOXML serialization helpers
//
// Shared constants, unit conversions, and namespace URIs live in DocxCommon.h.
// This file is intentionally empty; the public interface is declared in
// DocxConverter.h and implemented by the two files above.

#include "DocxConverter.h"
