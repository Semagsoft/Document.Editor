cmake_minimum_required(VERSION 3.20)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)

find_package(Qt6 REQUIRED COMPONENTS Core)

add_library(ExamplePlugin SHARED
    ExamplePlugin.h
    ExamplePlugin.cpp
)

target_include_directories(ExamplePlugin PRIVATE
    ${CMAKE_SOURCE_DIR}/src
)

target_link_libraries(ExamplePlugin PRIVATE
    Qt6::Core
)

# Ensure the plugin is placed in a discoverable location
install(TARGETS ExamplePlugin
    LIBRARY DESTINATION ${CMAKE_INSTALL_BINDIR}/plugins
)

# During development, also drop the plugin next to the app binary so the
# PluginManager can discover it without a separate install step.
if(TARGET Document.Editor)
    add_custom_command(TARGET ExamplePlugin POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E make_directory
                $<TARGET_FILE_DIR:Document.Editor>/plugins
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
                $<TARGET_FILE:ExamplePlugin>
                $<TARGET_FILE_DIR:Document.Editor>/plugins/
        COMMENT "Copying ExamplePlugin next to Document.Editor"
    )
endif()
