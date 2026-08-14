cmake_minimum_required(VERSION 3.20)
project(Document.Editor VERSION 1.0 LANGUAGES C CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

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

add_compile_definitions(APP_VERSION="${PROJECT_VERSION}")

add_subdirectory(third_party/miniz)
add_subdirectory(src)

# Example plugin — built so the plugin loading path is exercised.
add_subdirectory(plugins/example)

enable_testing()
add_subdirectory(test)

# Install
include(GNUInstallDirs)
install(TARGETS Document.Editor
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
)
install(FILES README.md LICENSE
    DESTINATION ${CMAKE_INSTALL_DOCDIR}
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

include(CPack)
