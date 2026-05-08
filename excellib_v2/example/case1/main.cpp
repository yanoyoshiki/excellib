#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <cstdint>
#include "ui_app.hpp"

int APIENTRY WinMain(HINSTANCE hInst, HINSTANCE, LPSTR, int) {
    // DPI awareness は manifest 側で宣言済み。
    // 古い OS / SDK でもビルドできるよう、コード側は GetProcAddress で動的解決し、
    // 値は HANDLE 互換の数値で扱う (DPI_AWARENESS_CONTEXT 型に依存しない)。
    typedef BOOL (WINAPI *PFN_SPDAC)(HANDLE);
    typedef BOOL (WINAPI *PFN_SPDA)(void);
    if (HMODULE u32 = GetModuleHandleW(L"user32.dll")) {
        auto fn  = reinterpret_cast<PFN_SPDAC>(
            reinterpret_cast<void(*)()>(GetProcAddress(u32, "SetProcessDpiAwarenessContext")));
        auto fn2 = reinterpret_cast<PFN_SPDA>(
            reinterpret_cast<void(*)()>(GetProcAddress(u32, "SetProcessDPIAware")));
        if (fn) {
            // -4 = DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
            fn(reinterpret_cast<HANDLE>(static_cast<intptr_t>(-4)));
        } else if (fn2) {
            fn2();
        }
    }

    case1::AppState state;
    case1::MainWindow app(hInst, state);
    if (!app.create()) return 1;
    return app.run_message_loop();
}
#else
// macOS / Linux ではビルドだけ通す（実行時は何もしない）
#include <iostream>
int main() {
    std::cout << "case1 is a Win32 GUI app. Build & run on Windows.\n";
    return 0;
}
#endif
