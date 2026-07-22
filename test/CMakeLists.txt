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

add_document_editor_test(test_FtpClient test_FtpClient.cpp)

# Editor tests
add_document_editor_test(test_DocumentEditor test_DocumentEditor.cpp)

add_document_editor_test(test_DocumentManager test_DocumentManager.cpp)

# Plugin tests
add_document_editor_test(test_PluginManager test_PluginManager.cpp)

# Settings test
add_document_editor_test(test_Settings test_Settings.cpp)

# ThemeManager test
add_document_editor_test(test_ThemeManager test_ThemeManager.cpp)

# DocumentService test
add_document_editor_test(test_DocumentService test_DocumentService.cpp)
