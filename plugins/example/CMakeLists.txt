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
