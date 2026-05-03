cmake_minimum_required(VERSION 3.16)
project(excellib VERSION 2.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# ============================================================
#  コンパイラ設定
# ============================================================
if(MSVC)
    add_compile_options(/W4 /WX /permissive- /wd4068 /utf-8)
else()
    add_compile_options(-Wall -Wextra -Wno-trigraphs -Wno-unused-parameter)
endif()

# MinGW: ランタイム DLL なしで動くスタンドアロン exe を生成する
if(MINGW)
    add_link_options(-static -static-libgcc -static-libstdc++)
endif()

# ============================================================
#  ライブラリ本体 (外部依存なし)
# ============================================================
add_library(excellib STATIC
    src/common/common.cpp
    src/common/factory.cpp
    src/common/deflate.cpp
    src/xls/xls_parser.cpp
    src/xlsx/xlsx_parser.cpp
    src/print/xlsx_page_setup.cpp
    src/print/excel_printer.cpp
    src/print/batch_printer.cpp
)

target_include_directories(excellib
    PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
        $<INSTALL_INTERFACE:include>
)

# Windows: Excel COM に必要なライブラリ
if(WIN32)
    target_link_libraries(excellib PUBLIC ole32 oleaut32 winspool)
endif()

# ============================================================
#  テスト
# ============================================================
enable_testing()

add_executable(test_all tests/test_all.cpp)
target_link_libraries(test_all PRIVATE excellib)
add_test(NAME excellib_tests COMMAND test_all)

# ============================================================
#  サンプル
# ============================================================
foreach(n 01_read 02_write 03_print 04_batch_print 05_error_handling)
    add_executable(${n} examples/${n}.cpp)
    target_link_libraries(${n} PRIVATE excellib)
endforeach()

# ============================================================
#  インストール
# ============================================================
include(GNUInstallDirs)

install(TARGETS excellib
    ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
)
install(DIRECTORY include/
    DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
)
install(DIRECTORY docs/
    DESTINATION ${CMAKE_INSTALL_DOCDIR}
)

# ============================================================
#  ビルド情報
# ============================================================
message(STATUS "excellib v${PROJECT_VERSION}")
message(STATUS "  Build type : ${CMAKE_BUILD_TYPE}")
message(STATUS "  C++ std    : C++${CMAKE_CXX_STANDARD}")
message(STATUS "  Compiler   : ${CMAKE_CXX_COMPILER_ID}")
message(STATUS "  外部依存   : なし (DEFLATE/CRC32は内部実装)")
