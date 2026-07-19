find_package(Qt6 REQUIRED COMPONENTS Test)

function(add_document_editor_test test_name)
    add_executable(${test_name} ${ARGN})
    target_include_directories(${test_name} PRIVATE ${CMAKE_SOURCE_DIR}/src)
    target_link_libraries(${test_name} PRIVATE Qt6::Test)
    add_test(NAME ${test_name} COMMAND ${test_name})
endfunction()

# Standalone utility tests
add_document_editor_test(test_TextUtils test_TextUtils.cpp
    ${CMAKE_SOURCE_DIR}/src/utils/TextUtils.cpp
)

add_document_editor_test(test_DpiUtils test_DpiUtils.cpp
    ${CMAKE_SOURCE_DIR}/src/utils/DpiUtils.cpp
)
target_link_libraries(test_DpiUtils PRIVATE Qt6::Gui Qt6::Widgets)

add_document_editor_test(test_SpellChecker test_SpellChecker.cpp
    ${CMAKE_SOURCE_DIR}/src/utils/SpellChecker.cpp
)

add_document_editor_test(test_NetworkUtils test_NetworkUtils.cpp
    ${CMAKE_SOURCE_DIR}/src/utils/NetworkUtils.cpp
)
target_link_libraries(test_NetworkUtils PRIVATE Qt6::Network)

# Converter tests
add_document_editor_test(test_XamlConverter test_XamlConverter.cpp
    ${CMAKE_SOURCE_DIR}/src/converters/XamlConverter.cpp
)
target_link_libraries(test_XamlConverter PRIVATE Qt6::Gui Qt6::Xml)

add_document_editor_test(test_RtfConverter test_RtfConverter.cpp
    ${CMAKE_SOURCE_DIR}/src/converters/RtfConverter.cpp
)
target_link_libraries(test_RtfConverter PRIVATE Qt6::Gui)

add_document_editor_test(test_FtpClient test_FtpClient.cpp
    ${CMAKE_SOURCE_DIR}/src/converters/FtpClient.cpp
)
target_link_libraries(test_FtpClient PRIVATE Qt6::Network)

# Editor tests
add_document_editor_test(test_DocumentEditor test_DocumentEditor.cpp
    ${CMAKE_SOURCE_DIR}/src/editor/DocumentEditor.cpp
    ${CMAKE_SOURCE_DIR}/src/services/SpellCheckHighlighter.cpp
    ${CMAKE_SOURCE_DIR}/src/utils/SpellChecker.cpp
    ${CMAKE_SOURCE_DIR}/src/utils/TextUtils.cpp
    ${CMAKE_SOURCE_DIR}/src/converters/XamlConverter.cpp
    ${CMAKE_SOURCE_DIR}/src/converters/RtfConverter.cpp
)
target_link_libraries(test_DocumentEditor PRIVATE Qt6::Gui Qt6::Widgets Qt6::Xml)

set(WIDGET_SOURCES
    ${CMAKE_SOURCE_DIR}/src/widgets/TableGridPicker.cpp
    ${CMAKE_SOURCE_DIR}/src/widgets/SymbolPicker.cpp
)

set(DIALOG_SOURCES
    ${CMAKE_SOURCE_DIR}/src/dialogs/InsertTableDialog.cpp
    ${CMAKE_SOURCE_DIR}/src/dialogs/InsertImageDialog.cpp
    ${CMAKE_SOURCE_DIR}/src/dialogs/InsertShapeDialog.cpp
    ${CMAKE_SOURCE_DIR}/src/dialogs/InsertChartDialog.cpp
    ${CMAKE_SOURCE_DIR}/src/dialogs/InsertLinkDialog.cpp
    ${CMAKE_SOURCE_DIR}/src/dialogs/InsertSymbolDialog.cpp
    ${CMAKE_SOURCE_DIR}/src/dialogs/InsertDateDialog.cpp
    ${CMAKE_SOURCE_DIR}/src/dialogs/InsertTimeDialog.cpp
    ${CMAKE_SOURCE_DIR}/src/dialogs/InsertVideoDialog.cpp
    ${CMAKE_SOURCE_DIR}/src/dialogs/InsertLineDialog.cpp
    ${WIDGET_SOURCES}
)

add_document_editor_test(test_DocumentManager test_DocumentManager.cpp
    ${CMAKE_SOURCE_DIR}/src/editor/DocumentManager.cpp
    ${CMAKE_SOURCE_DIR}/src/editor/DocumentEditor.cpp
    ${CMAKE_SOURCE_DIR}/src/editor/DocumentTab.cpp
    ${CMAKE_SOURCE_DIR}/src/services/SpellCheckHighlighter.cpp
    ${CMAKE_SOURCE_DIR}/src/utils/SpellChecker.cpp
    ${CMAKE_SOURCE_DIR}/src/utils/TextUtils.cpp
    ${CMAKE_SOURCE_DIR}/src/widgets/RulerWidget.cpp
    ${CMAKE_SOURCE_DIR}/src/utils/DpiUtils.cpp
    ${CMAKE_SOURCE_DIR}/src/services/DocumentService.cpp
    ${CMAKE_SOURCE_DIR}/src/converters/FtpClient.cpp
    ${CMAKE_SOURCE_DIR}/src/converters/XamlConverter.cpp
    ${CMAKE_SOURCE_DIR}/src/converters/RtfConverter.cpp
    ${DIALOG_SOURCES}
)
target_link_libraries(test_DocumentManager PRIVATE Qt6::Gui Qt6::Widgets Qt6::PrintSupport Qt6::Xml Qt6::Charts Qt6::Network)

# Plugin tests
add_document_editor_test(test_PluginManager test_PluginManager.cpp
    ${CMAKE_SOURCE_DIR}/src/plugins/PluginManager.cpp
    ${CMAKE_SOURCE_DIR}/src/plugins/PluginContext.cpp
)

# Settings test
add_document_editor_test(test_Settings test_Settings.cpp
    ${CMAKE_SOURCE_DIR}/src/app/Settings.cpp
)
target_link_libraries(test_Settings PRIVATE Qt6::Gui)

# ThemeManager test
add_document_editor_test(test_ThemeManager test_ThemeManager.cpp
    ${CMAKE_SOURCE_DIR}/src/mainwindow/ThemeManager.cpp
)
target_link_libraries(test_ThemeManager PRIVATE Qt6::Gui Qt6::Widgets)

# DocumentService test
add_document_editor_test(test_DocumentService test_DocumentService.cpp
    ${CMAKE_SOURCE_DIR}/src/services/DocumentService.cpp
    ${CMAKE_SOURCE_DIR}/src/editor/DocumentManager.cpp
    ${CMAKE_SOURCE_DIR}/src/editor/DocumentEditor.cpp
    ${CMAKE_SOURCE_DIR}/src/editor/DocumentTab.cpp
    ${CMAKE_SOURCE_DIR}/src/services/SpellCheckHighlighter.cpp
    ${CMAKE_SOURCE_DIR}/src/utils/SpellChecker.cpp
    ${CMAKE_SOURCE_DIR}/src/widgets/RulerWidget.cpp
    ${CMAKE_SOURCE_DIR}/src/utils/DpiUtils.cpp
    ${CMAKE_SOURCE_DIR}/src/utils/TextUtils.cpp
    ${CMAKE_SOURCE_DIR}/src/converters/FtpClient.cpp
    ${CMAKE_SOURCE_DIR}/src/converters/XamlConverter.cpp
    ${CMAKE_SOURCE_DIR}/src/converters/RtfConverter.cpp
    ${CMAKE_SOURCE_DIR}/src/app/Settings.cpp
    ${CMAKE_SOURCE_DIR}/src/mainwindow/StatusBarManager.cpp
    ${DIALOG_SOURCES}
)
target_link_libraries(test_DocumentService PRIVATE Qt6::Gui Qt6::Widgets Qt6::PrintSupport Qt6::Xml Qt6::Network Qt6::Charts)
