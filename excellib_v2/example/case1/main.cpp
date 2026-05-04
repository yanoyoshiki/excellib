#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include "ui_app.hpp"

int APIENTRY WinMain(HINSTANCE hInst, HINSTANCE, LPSTR, int) {
    // DPI awareness は manifest で宣言済み。
    // コード側からも段階的にフォールバック設定 (古い OS 向け)
#ifndef DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
    #define DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 ((HANDLE)-4)
#endif
    typedef BOOL (WINAPI *PFN_SPDAC)(HANDLE);
    typedef BOOL (WINAPI *PFN_SPDA)();
    if (HMODULE u32 = GetModuleHandleW(L"user32.dll")) {
        if (auto fn = (PFN_SPDAC)GetProcAddress(u32, "SetProcessDpiAwarenessContext")) {
            fn(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
        } else if (auto fn2 = (PFN_SPDA)GetProcAddress(u32, "SetProcessDPIAware")) {
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
