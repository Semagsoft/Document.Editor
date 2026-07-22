set(SOURCES
    app/Application.cpp
    app/Settings.cpp
    mainwindow/ActionManager.cpp
    mainwindow/MainWindow.cpp
    mainwindow/ThemeManager.cpp
    mainwindow/StatusBarManager.cpp
    editor/DocumentEditor.cpp
    editor/DocumentTab.cpp
    editor/DocumentManager.cpp
    widgets/RulerWidget.cpp
    widgets/TableGridPicker.cpp
    widgets/SymbolPicker.cpp

    utils/DpiUtils.cpp
    dialogs/AboutDialog.cpp
    dialogs/GoToLineDialog.cpp
    dialogs/LineSpacingDialog.cpp
    dialogs/InsertTableDialog.cpp
    dialogs/InsertSymbolDialog.cpp
    dialogs/InsertLinkDialog.cpp
    dialogs/InsertDateDialog.cpp
    dialogs/InsertTimeDialog.cpp
    dialogs/InsertImageDialog.cpp
    dialogs/InsertShapeDialog.cpp
    dialogs/InsertChartDialog.cpp
    dialogs/InsertVideoDialog.cpp
    dialogs/InsertLineDialog.cpp
    dialogs/FindDialog.cpp
    dialogs/ReplaceDialog.cpp
    dialogs/OptionsDialog.cpp
    dialogs/SpellCheckDialog.cpp
    utils/TextUtils.cpp
    utils/NetworkUtils.cpp
    utils/SpellChecker.cpp
    converters/FtpClient.cpp
    converters/XamlConverter.cpp
    converters/RtfConverter.cpp
    converters/DocxConverter.cpp
    plugins/PluginManager.cpp
    plugins/PluginContext.cpp
    services/DocumentService.cpp
    services/SpellCheckHighlighter.cpp
)

set(HEADERS
    app/Application.h
    app/Settings.h
    mainwindow/ActionManager.h
    mainwindow/MainWindow.h
    mainwindow/ThemeManager.h
    mainwindow/StatusBarManager.h
    editor/DocumentEditor.h
    editor/DocumentTab.h
    editor/DocumentManager.h
    widgets/RulerWidget.h
    widgets/TableGridPicker.h
    widgets/SymbolPicker.h

    utils/DpiUtils.h
    dialogs/AboutDialog.h
    dialogs/GoToLineDialog.h
    dialogs/LineSpacingDialog.h
    dialogs/InsertTableDialog.h
    dialogs/InsertSymbolDialog.h
    dialogs/InsertLinkDialog.h
    dialogs/InsertDateDialog.h
    dialogs/InsertTimeDialog.h
    dialogs/InsertImageDialog.h
    dialogs/InsertShapeDialog.h
    dialogs/InsertChartDialog.h
    dialogs/InsertVideoDialog.h
    dialogs/InsertLineDialog.h
    dialogs/FindDialog.h
    dialogs/ReplaceDialog.h
    dialogs/OptionsDialog.h
    dialogs/SpellCheckDialog.h
    utils/TextUtils.h
    utils/NetworkUtils.h
    converters/FtpClient.h
    converters/XamlConverter.h
    converters/RtfConverter.h
    converters/DocxConverter.h
    plugins/IPlugin.h
    plugins/PluginManager.h
    plugins/PluginContext.h
    services/DocumentService.h
    services/SpellCheckHighlighter.h
)

set(RESOURCES
    ${CMAKE_SOURCE_DIR}/resources/resources.qrc
)

find_package(Qt6CorePrivate REQUIRED)

add_library(de_objects OBJECT ${SOURCES} ${HEADERS})
target_include_directories(de_objects PUBLIC ${CMAKE_CURRENT_SOURCE_DIR})
target_link_libraries(de_objects PUBLIC Qt6::CorePrivate)
target_link_libraries(de_objects PUBLIC
    Qt6::Core
    Qt6::Gui
    Qt6::Widgets
    Qt6::PrintSupport
    Qt6::Network
    Qt6::TextToSpeech
    Qt6::Xml
    Qt6::Charts
)

add_executable(Document.Editor main.cpp ${RESOURCES} $<TARGET_OBJECTS:de_objects>)
target_link_libraries(Document.Editor PRIVATE de_objects)
