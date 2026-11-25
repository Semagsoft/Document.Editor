find_package(Qt6 REQUIRED COMPONENTS Test)

function(add_document_editor_test test_name)
    add_executable(${test_name} ${ARGN})
    target_link_libraries(${test_name} PRIVATE de_objects Qt6::Test)
    add_test(NAME ${test_name} COMMAND ${test_name})
endfunction()

# Standalone utility tests
add_document_editor_test(test_TextUtils test_TextUtils.cpp)

add_document_editor_test(test_DpiUtils test_DpiUtils.cpp)

add_document_editor_test(test_SpellChecker test_SpellChecker.cpp)

add_document_editor_test(test_NetworkUtils test_NetworkUtils.cpp)

# Converter tests
add_document_editor_test(test_XamlConverter test_XamlConverter.cpp)

add_document_editor_test(test_RtfConverter test_RtfConverter.cpp)

add_document_editor_test(test_DocxConverter test_DocxConverter.cpp)
target_compile_definitions(test_DocxConverter PRIVATE
    DOCX_FIXTURE_DIR="${CMAKE_CURRENT_SOURCE_DIR}/fixtures")

# Malformed-input / fuzz-style tests for all converters (security hardening).
add_document_editor_test(test_MalformedInput test_MalformedInput.cpp)

add_document_editor_test(test_FtpClient test_FtpClient.cpp)

# Editor tests
add_document_editor_test(test_DocumentEditor test_DocumentEditor.cpp)

add_document_editor_test(test_DocumentManager test_DocumentManager.cpp)

# Plugin tests
add_document_editor_test(test_PluginManager test_PluginManager.cpp)
target_compile_definitions(test_PluginManager PRIVATE
    PLUGIN_FIXTURE_DIR="${CMAKE_BINARY_DIR}/src/plugins")
# The test loads the real example plugin; make sure it is built and copied
# into the fixture directory first (no declared dependency existed before).
if(TARGET ExamplePlugin)
    add_dependencies(test_PluginManager ExamplePlugin)
endif()

# Settings test
add_document_editor_test(test_Settings test_Settings.cpp)

# ThemeManager test
add_document_editor_test(test_ThemeManager test_ThemeManager.cpp)

# DocumentService test
add_document_editor_test(test_DocumentService test_DocumentService.cpp)

# Roundtrip/regression tests (converter roundtrips, isSafeFilePath)
add_document_editor_test(test_RoundTrip test_RoundTrip.cpp)

# ---------------------------------------------------------------------------
# libFuzzer target for the DOCX converter (security hardening).
#
# Built with AddressSanitizer + UndefinedBehaviorSanitizer + libFuzzer.
# Requires clang++ (g++ does not support -fsanitize=fuzzer). Not registered
# with add_test() — it runs until it finds a crash or hits a time limit,
# which is incompatible with ctest's per-test timeout model.
#
# Build manually:
#   CXX=clang++ cmake -B build -DCMAKE_BUILD_TYPE=Debug
#   cmake --build build --target test_DocxFuzzer
#
# Run manually:
#   ./build/test/test_DocxFuzzer -max_len=100000 -max_total_time=60 ./corpus
# ---------------------------------------------------------------------------
if(CMAKE_CXX_COMPILER_ID MATCHES "Clang")
    add_executable(test_DocxFuzzer test_DocxFuzzer.cpp)
    target_link_libraries(test_DocxFuzzer PRIVATE de_objects)
    target_compile_options(test_DocxFuzzer PRIVATE
        -fsanitize=fuzzer,address,undefined
        -Wno-literal-conversion)
    target_link_options(test_DocxFuzzer PRIVATE
        -fsanitize=fuzzer,address,undefined
        -fno-omit-frame-pointer)
else()
    message(STATUS "test_DocxFuzzer skipped: requires clang++ (current CXX: ${CMAKE_CXX_COMPILER_ID})")
endif()
