cmake_minimum_required(VERSION 3.20)
project(Document.Editor VERSION 1.0 LANGUAGES C CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

# Must come before any add_subdirectory() so nested projects (and the
# example plugin) can rely on CMAKE_INSTALL_BINDIR etc. being defined.
include(GNUInstallDirs)

find_package(Qt6 REQUIRED COMPONENTS
    Core
    Gui
    Widgets
    PrintSupport
    Network
    TextToSpeech
    Xml
    Charts
    Concurrent
)

# Enable compiler warnings so issues are surfaced in CI. -Werror keeps the
# tree warning-free; -Wsfinae-incomplete is disabled because Qt's moc output
# trips it on GCC (it forward-declares the class being defined). This flag is
# GCC-specific and is not recognized by clang.
if(MSVC)
    add_compile_options(/W4)
else()
    add_compile_options(-Wall -Wextra -Werror)
    if(CMAKE_CXX_COMPILER_ID MATCHES "GNU")
        add_compile_options($<$<COMPILE_LANGUAGE:CXX>:-Wno-sfinae-incomplete>)
    endif()
endif()

add_compile_definitions(APP_VERSION="${PROJECT_VERSION}")

add_subdirectory(third_party/miniz)
add_subdirectory(src)

# Example plugin — built so the plugin loading path is exercised.
add_subdirectory(plugins/example)

enable_testing()
add_subdirectory(test)

# Install
install(TARGETS Document.Editor
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
)
install(FILES README.md LICENSE
    DESTINATION ${CMAKE_INSTALL_DOCDIR}
)
# Linux desktop integration: icon + .desktop entry so the app shows up in
# application menus.
install(FILES resources/icons/app.png
    DESTINATION ${CMAKE_INSTALL_DATADIR}/icons/hicolor/64x64/apps
    RENAME document-editor.png
)
install(FILES packaging/document-editor.desktop
    DESTINATION ${CMAKE_INSTALL_DATADIR}/applications
)

# CPack
set(CPACK_PACKAGE_NAME "Document.Editor")
set(CPACK_PACKAGE_VERSION "${PROJECT_VERSION}")
set(CPACK_PACKAGE_VENDOR "Semagsoft")
set(CPACK_PACKAGE_DESCRIPTION_SUMMARY "A cross-platform rich text document editor built with Qt6 and C++17")
set(CPACK_RESOURCE_FILE_LICENSE "${CMAKE_SOURCE_DIR}/LICENSE")
set(CPACK_RESOURCE_FILE_README "${CMAKE_SOURCE_DIR}/README.md")

if(WIN32)
    set(CPACK_GENERATOR "NSIS")
elseif(APPLE)
    set(CPACK_GENERATOR "DMG")
else()
    set(CPACK_GENERATOR "DEB;TGZ")
endif()

# DEB metadata: without dependencies and a maintainer the package installs
# broken on a clean machine.
set(CPACK_DEBIAN_PACKAGE_MAINTAINER "Semagsoft")
set(CPACK_DEBIAN_PACKAGE_SECTION "editors")
set(CPACK_DEBIAN_PACKAGE_PRIORITY "optional")
set(CPACK_DEBIAN_PACKAGE_HOMEPAGE "https://documenteditor.net/")
set(CPACK_DEBIAN_PACKAGE_DESCRIPTION
    "A cross-platform rich text document editor built with Qt6.\n\
 Supports XAML, DOCX, HTML, RTF and plain text, multi-tab editing,\n\
 plugins, themes, page layout and print support.")
set(CPACK_DEBIAN_PACKAGE_DEPENDS
    "libqt6core6, libqt6gui6, libqt6widgets6, libqt6network6, libqt6printsupport6, libqt6texttospeech6, libqt6xml6, libqt6charts6, libqt6concurrent6")
set(CPACK_PACKAGE_CONTACT "Semagsoft")

include(CPack)
