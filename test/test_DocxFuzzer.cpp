// libFuzzer target for DocxConverter::loadFromDocx.
//
// Build with:
//   cmake -B build -DCMAKE_BUILD_TYPE=Debug \
//     -DCMAKE_CXX_FLAGS="-fsanitize=fuzzer,address,undefined" \
//     -DCMAKE_EXE_LINKER_FLAGS="-fsanitize=fuzzer,address,undefined"
//   cmake --build build --target test_DocxFuzzer
//
// Run:
//   ./build/test/test_DocxFuzzer -max_len=100000 -max_total_time=60 ./corpus
//
// Feeds arbitrary bytes to the DOCX loader. AddressSanitizer and
// UndefinedBehaviorSanitizer catch memory safety issues; libFuzzer explores
// the input space to find crashes that unit tests miss.
#include <cstdint>
#include <cstddef>

#include <QTextDocument>
#include <QMarginsF>
#include <QColor>

#include "converters/DocxConverter.h"

extern "C" int LLVMFuzzerTestOneInput(const uint8_t *data, size_t size)
{
    QTextDocument doc;
    QMarginsF margins;
    QColor bg;
    DocxConverter::loadFromDocx(
        QByteArray(reinterpret_cast<const char *>(data), size),
        &doc, margins, bg);
    return 0;
}
