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
    ui_stages.cpp
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
target_include_directories(case1 PRIVATE ${CMAKE_CURRENT_SOURCE_DIR})

# ============================================================
#  Windows 固有の定義・ライブラリ・コンパイルオプション
# ============================================================
if(WIN32)
    # Win10 (0x0A00) を最低ターゲットとして API を露出する。
    # SetProcessDpiAwarenessContext / GetDpiForSystem / EM_SETCUEBANNER 等を使うため。
    target_compile_definitions(case1 PRIVATE
        _WIN32_WINNT=0x0A00
        WINVER=0x0A00
        UNICODE _UNICODE          # 念のため (W API を明示しているが整合性のため)
        NOMINMAX                  # windows.h の min/max マクロを抑止
        WIN32_LEAN_AND_MEAN       # 各 .cpp の重複定義防止
    )

    # 必要な Win32 ライブラリを明示
    #   ole32     : COM (CoTaskMemFree)
    #   comctl32  : Common Controls (ListView / Progress / TabControl / EM_SETCUEBANNER)
    #   comdlg32  : GetSaveFileNameW / OPENFILENAMEW
    #   shell32   : SHBrowseForFolderW / SHGetPathFromIDListW
    #   gdi32     : フォント・描画
    #   user32    : ウィンドウ・メッセージ
    #   uxtheme   : Visual Styles (manifest 連携)
    target_link_libraries(case1 PRIVATE
        comctl32 comdlg32 shell32 ole32 gdi32 user32 uxtheme
    )
endif()

# MSVC 警告レベル
if(MSVC)
    target_compile_options(case1 PRIVATE /W4 /utf-8 /wd4100 /wd4505)
endif()

# GCC / Clang 共通
if(NOT MSVC)
    target_compile_options(case1 PRIVATE -Wall -Wextra -Wno-unused-parameter)
endif()

# ============================================================
#  MinGW / TDM-GCC 専用設定
# ============================================================
if(MINGW)
    # スタンドアロン exe (ランタイム DLL なし) + WIN32 サブシステム
    target_link_options(case1 PRIVATE
        -static -static-libgcc -static-libstdc++ -mwindows
    )

    # ソースは UTF-8、wide 文字リテラルは UTF-16LE で出力
    # これがないと L"日本語" が壊れた wchar_t 列になる
    target_compile_options(case1 PRIVATE
        -finput-charset=UTF-8
        -fexec-charset=UTF-8
        -fwide-exec-charset=UTF-16LE
    )

    # MinGW で過剰に出る警告を case1 だけ抑制
    target_compile_options(case1 PRIVATE
        -Wno-cast-function-type      # GetProcAddress の戻りキャスト
        -Wno-missing-field-initializers  # Win32 構造体の {} 初期化
        -Wno-pedantic-ms-format      # printf 拡張系
    )

    # windres も UTF-8 で読み込む
    if(CMAKE_RC_COMPILER)
        set(CMAKE_RC_FLAGS "${CMAKE_RC_FLAGS} -c65001")
    endif()
endif()

message(STATUS "case1: GUI app for batch printing via excellib")
