cmake_minimum_required(VERSION 3.20)

project(miniz VERSION 3.0.2 LANGUAGES C)

add_library(miniz STATIC
    miniz.c
    miniz_tdef.c
    miniz_tinfl.c
    miniz_zip.c
)

target_include_directories(miniz PUBLIC ${CMAKE_CURRENT_SOURCE_DIR})
set_target_properties(miniz PROPERTIES
    POSITION_INDEPENDENT_CODE ON
)
