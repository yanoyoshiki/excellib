cmake_minimum_required(VERSION 3.16)
project(case1 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# ============================================================
#  excellib (sibling repo) を取り込む
# ============================================================
if(NOT TARGET excellib)
    add_subdirectory(${CMAKE_CURRENT_SOURCE_DIR}/../../excellib_v2
                     ${CMAKE_CURRENT_BINARY_DIR}/excellib_build)
endif()

# ============================================================
#  Win32 GUI 実行ファイル
# ============================================================
if(NOT WIN32)
    message(WARNING "case1 is a Win32 GUI app. Configure on Windows or with a Win32 cross compiler.")
endif()

set(CASE1_SOURCES
    main.cpp
    workflow.cpp
    ui_theme.cpp
    ui_app.cpp
    ui_stage1_input.cpp
    ui_stage2_master.cpp
    ui_stage3_jobs.cpp
    ui_stage4_run.cpp
)

set(CASE1_HEADERS
    config.hpp
    workflow.hpp
    ui_theme.hpp
    ui_app.hpp
    ui_stages.hpp
)

if(WIN32)
    add_executable(case1 WIN32
        ${CASE1_SOURCES}
        ${CASE1_HEADERS}
        app.rc
    )
else()
    # macOS / Linux ではコンパイルだけ確認できるよう、ダミー的にビルドする
    add_executable(case1 ${CASE1_SOURCES} ${CASE1_HEADERS})
endif()

target_link_libraries(case1 PRIVATE excellib)

if(WIN32)
    target_link_libraries(case1 PRIVATE comctl32 gdi32 user32 uxtheme shell32)
endif()

# MSVC 警告レベル
if(MSVC)
    target_compile_options(case1 PRIVATE /W4 /utf-8 /wd4100 /wd4505)
else()
    target_compile_options(case1 PRIVATE -Wall -Wextra -Wno-unused-parameter)
endif()

# MinGW 静的リンク（スタンドアロン exe）
if(MINGW)
    target_link_options(case1 PRIVATE -static -static-libgcc -static-libstdc++ -mwindows)
endif()

target_include_directories(case1 PRIVATE ${CMAKE_CURRENT_SOURCE_DIR})

message(STATUS "case1: GUI app for batch printing via excellib")
