#include <napi.h>

#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#define _WIN32_WINNT 0x0601
#include <windows.h>
#include <dwmapi.h>
#include <psapi.h>
#include <unknwn.h>
#include <objbase.h>
#include <shellapi.h>
#include <shobjidl.h>
#include <propsys.h>
#include <propkey.h>
#include <vector>
#include <map>
#include <string>
#include <algorithm>
#include <cmath>
// GDI+ requires min/max; with NOMINMAX define we provide temporary macros.
#ifndef min
#define min(a,b) ((a) < (b) ? (a) : (b))
#endif
#ifndef max
#define max(a,b) ((a) > (b) ? (a) : (b))
#endif
#include <gdiplus.h>
#undef min
#undef max
#ifdef ENABLE_WGC
#include <d3d11.h>
#include <dxgi1_2.h>
// C++/WinRT (Windows Graphics Capture & Imaging)
#include <winrt/base.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Graphics.Capture.h>
#include <winrt/Windows.Graphics.DirectX.h>
#include <winrt/Windows.Graphics.DirectX.Direct3D11.h>
#include <winrt/Windows.Graphics.Imaging.h>
#include <winrt/Windows.Storage.Streams.h>
// Interop to create GraphicsCaptureItem from HWND
#include <windows.graphics.capture.interop.h>
#include <windows.graphics.directx.direct3d11.interop.h>
#endif
#include <sstream>
#include <iomanip>
#include <unordered_map>
#include <mutex>
#include <functional>

// PrintWindow Flags definieren falls nicht verfügbar
#ifndef PW_RENDERFULLCONTENT
#define PW_RENDERFULLCONTENT 0x00000002
#endif

#ifndef PW_CLIENTONLY
#define PW_CLIENTONLY 0x00000001
#endif

// WinEvent constants for cloaking were introduced in newer SDKs; define if missing
#ifndef EVENT_OBJECT_CLOAKED
#define EVENT_OBJECT_CLOAKED 0x8017
#endif
#ifndef EVENT_OBJECT_UNCLOAKED
#define EVENT_OBJECT_UNCLOAKED 0x8018
#endif

#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "psapi.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "propsys.lib")
#pragma comment(lib, "gdiplus.lib")
#ifdef ENABLE_WGC
#pragma comment(lib, "d3d11.lib")
#pragma comment(lib, "dxgi.lib")
#pragma comment(lib, "windowsapp.lib")
#endif

using namespace Napi;

// ---------------- UTF-8 / UTF-16 helpers ----------------
// Forward declarations for helpers used by event hooks
static std::string GetWindowTitle(HWND hwnd);
static std::string GetExecutablePath(HWND hwnd);
static std::string WideToUtf8(const std::wstring& w) {
    if (w.empty()) return std::string();
    int sizeNeeded = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), nullptr, 0, nullptr, nullptr);
    if (sizeNeeded <= 0) return std::string();
    std::string result((size_t)sizeNeeded, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), result.data(), sizeNeeded, nullptr, nullptr);
    return result;
}

static std::wstring Utf8ToWide(const std::string& s) {
    if (s.empty()) return std::wstring();
    int sizeNeeded = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), nullptr, 0);
    if (sizeNeeded <= 0) return std::wstring();
    std::wstring result((size_t)sizeNeeded, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), result.data(), sizeNeeded);
    return result;
}

struct WindowInfo {
    HWND hwnd;
    std::string title;
    std::string executablePath;
    bool isVisible;
};


// GDI+ global token
static ULONG_PTR g_gdiplusToken = 0;
static void GdiplusCleanup(void* /*arg*/) {
    if (g_gdiplusToken) {
        Gdiplus::GdiplusShutdown(g_gdiplusToken);
        g_gdiplusToken = 0;
    }
}
// WinRT apartment lifetime
#ifdef ENABLE_WGC
static bool g_winrtInited = false;
static void WinrtInitOnce() {
    if (!g_winrtInited) {
        try { winrt::init_apartment(winrt::apartment_type::single_threaded); g_winrtInited = true; } catch (...) { }
    }
}
static void WinrtCleanup(void* /*arg*/) {
    if (g_winrtInited) {
        try { winrt::uninit_apartment(); } catch (...) {}
        g_winrtInited = false;
    }
}
#endif

// Caches
struct ThumbCacheEntry {
    std::string base64;
    RECT rect;
    ULONGLONG ts;
    int w;
    int h;
};
static std::unordered_map<HWND, ThumbCacheEntry> g_thumbCache;
static std::unordered_map<HWND, std::string> g_iconCache;
static const ULONGLONG THUMB_TTL_MS = 800; // Cache thumbnails for ~0.8s
static std::mutex g_cacheMutex; // Protects g_thumbCache and g_iconCache

// ---------------- Window Event Hooks (Create/Destroy/Focus) ----------------
static HWINEVENTHOOK g_hookCreate = nullptr;
static HWINEVENTHOOK g_hookDestroy = nullptr;
static HWINEVENTHOOK g_hookForeground = nullptr;
static HWINEVENTHOOK g_hookShow = nullptr;
static HWINEVENTHOOK g_hookHide = nullptr;
static HWINEVENTHOOK g_hookCloaked = nullptr;     // EVENT_OBJECT_CLOAKED
static HWINEVENTHOOK g_hookUncloaked = nullptr;   // EVENT_OBJECT_UNCLOAKED
static HWINEVENTHOOK g_hookMinStart = nullptr;    // EVENT_SYSTEM_MINIMIZESTART
static HWINEVENTHOOK g_hookMinEnd = nullptr;      // EVENT_SYSTEM_MINIMIZEEND
static HWINEVENTHOOK g_hookState = nullptr;       // EVENT_OBJECT_STATECHANGE

struct WindowEventPayload {
    HWND hwnd{};
    std::string title;
    std::string exePath;
    bool isVisible{};
};

static ThreadSafeFunction g_tsfnCreated;
static ThreadSafeFunction g_tsfnClosed;
static ThreadSafeFunction g_tsfnFocused;
static ThreadSafeFunction g_tsfnMinimized;
static ThreadSafeFunction g_tsfnRestored;
static ThreadSafeFunction g_tsfnChange; // unified event: type in payload

// Fallback poller when WinEvent hooks are unavailable in some environments
#include <thread>
#include <atomic>
#include <unordered_set>
static std::thread g_eventPollerThread;
static std::atomic<bool> g_eventPollerRunning{ false };
static std::atomic<bool> g_usingFallbackEvents{ false };
static std::atomic<ULONGLONG> g_lastHookEventTick{ 0 };

static void StartFallbackEventPoller();
static void StopFallbackEventPoller();

static bool IsTopLevelWindow(HWND hwnd) {
    if (!hwnd) return false;
    return GetAncestor(hwnd, GA_ROOT) == hwnd;
}

static WindowEventPayload MakePayload(HWND hwnd) {
    WindowEventPayload p{};
    p.hwnd = hwnd;
    p.title = GetWindowTitle(hwnd);
    p.exePath = GetExecutablePath(hwnd);
    p.isVisible = IsWindowVisible(hwnd) ? true : false;
    return p;
}

static void CALLBACK WinEventProcCB(HWINEVENTHOOK hook, DWORD event, HWND hwnd, LONG idObject, LONG idChild, DWORD, DWORD) {
    g_lastHookEventTick.store(GetTickCount64());
    // Filter only window object events or foreground changes
    if (event == EVENT_SYSTEM_FOREGROUND) {
        if (!IsWindow(hwnd)) return;
        // Map to top-level root to normalize hosted/UWP cases
        HWND top = GetAncestor(hwnd, GA_ROOT);
        if (top) hwnd = top;
        if (g_tsfnFocused || g_tsfnChange) {
            WindowEventPayload payload = MakePayload(hwnd);
            auto* heap = new WindowEventPayload(std::move(payload));
            if (g_tsfnFocused) {
                g_tsfnFocused.BlockingCall(new WindowEventPayload(*heap), [](Env env, Function cb, WindowEventPayload* data){
                    Object o = Object::New(env);
                    o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("title", String::New(env, data->title));
                    o.Set("executablePath", String::New(env, data->exePath));
                    o.Set("isVisible", Boolean::New(env, data->isVisible));
                    o.Set("type", String::New(env, "focused"));
                    cb.Call({ o });
                    delete data;
                });
            }
            if (g_tsfnChange) {
                g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                    Object o = Object::New(env);
                    o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("title", String::New(env, data->title));
                    o.Set("executablePath", String::New(env, data->exePath));
                    o.Set("isVisible", Boolean::New(env, data->isVisible));
                    o.Set("type", String::New(env, "focused"));
                    cb.Call({ o });
                    delete data;
                });
            } else {
                delete heap;
            }
        }
        return;
    }
    // Handle UWP-style minimize/restore via cloaking
    if (event == EVENT_OBJECT_CLOAKED) {
        if (idObject != OBJID_WINDOW) return;
        // Map child to top-level window to ensure hosted/UWP apps surface events
        HWND top = hwnd ? GetAncestor(hwnd, GA_ROOT) : NULL;
        if (top) hwnd = top;
        if (!IsWindow(hwnd) || !IsTopLevelWindow(hwnd)) return;
        if (g_tsfnMinimized || g_tsfnChange) {
            WindowEventPayload payload = MakePayload(hwnd);
            auto* heap = new WindowEventPayload(std::move(payload));
            if (g_tsfnMinimized) {
                g_tsfnMinimized.BlockingCall(new WindowEventPayload(*heap), [](Env env, Function cb, WindowEventPayload* data){
                    Object o = Object::New(env);
                    o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("title", String::New(env, data->title));
                    o.Set("executablePath", String::New(env, data->exePath));
                    o.Set("isVisible", Boolean::New(env, data->isVisible));
                    o.Set("type", String::New(env, "minimized"));
                    cb.Call({ o });
                    delete data;
                });
            }
            if (g_tsfnChange) {
                g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                    Object o = Object::New(env);
                    o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("title", String::New(env, data->title));
                    o.Set("executablePath", String::New(env, data->exePath));
                    o.Set("isVisible", Boolean::New(env, data->isVisible));
                    o.Set("type", String::New(env, "minimized"));
                    cb.Call({ o });
                    delete data;
                });
            } else {
                delete heap;
            }
        }
        return;
    }
    // Fallback: state changes can indicate minimized/restored transitions
    if (event == EVENT_OBJECT_STATECHANGE) {
        if (idObject != OBJID_WINDOW && idObject != OBJID_CLIENT) return;
        if (!IsWindow(hwnd)) return;
        HWND top = GetAncestor(hwnd, GA_ROOT);
        if (top) hwnd = top;
        if (!IsWindow(hwnd) || !IsTopLevelWindow(hwnd)) return;
        if (IsIconic(hwnd)) {
            if (g_tsfnMinimized || g_tsfnChange) {
                auto payload = MakePayload(hwnd);
                auto* heap = new WindowEventPayload(payload);
                if (g_tsfnMinimized) g_tsfnMinimized.BlockingCall(new WindowEventPayload(payload), [](Env env, Function cb, WindowEventPayload* data){
                    Object o = Object::New(env);
                    o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("title", String::New(env, data->title));
                    o.Set("executablePath", String::New(env, data->exePath));
                    o.Set("isVisible", Boolean::New(env, data->isVisible));
                    o.Set("type", String::New(env, "minimized"));
                    cb.Call({ o });
                    delete data;
                });
                if (g_tsfnChange) g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                    Object o = Object::New(env);
                    o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("title", String::New(env, data->title));
                    o.Set("executablePath", String::New(env, data->exePath));
                    o.Set("isVisible", Boolean::New(env, data->isVisible));
                    o.Set("type", String::New(env, "minimized"));
                    cb.Call({ o });
                    delete data;
                }); else delete heap;
            }
        } else {
            if ((g_tsfnRestored || g_tsfnChange) && IsWindowVisible(hwnd)) {
                auto payload = MakePayload(hwnd);
                auto* heap = new WindowEventPayload(payload);
                if (g_tsfnRestored) g_tsfnRestored.BlockingCall(new WindowEventPayload(payload), [](Env env, Function cb, WindowEventPayload* data){
                    Object o = Object::New(env);
                    o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("title", String::New(env, data->title));
                    o.Set("executablePath", String::New(env, data->exePath));
                    o.Set("isVisible", Boolean::New(env, data->isVisible));
                    o.Set("type", String::New(env, "restored"));
                    cb.Call({ o });
                    delete data;
                });
                if (g_tsfnChange) g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                    Object o = Object::New(env);
                    o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("title", String::New(env, data->title));
                    o.Set("executablePath", String::New(env, data->exePath));
                    o.Set("isVisible", Boolean::New(env, data->isVisible));
                    o.Set("type", String::New(env, "restored"));
                    cb.Call({ o });
                    delete data;
                }); else delete heap;
            }
        }
        // don't return; allow further specific handlers if they apply
    }
    if (event == EVENT_OBJECT_UNCLOAKED) {
        if (idObject != OBJID_WINDOW) return;
        HWND top = hwnd ? GetAncestor(hwnd, GA_ROOT) : NULL;
        if (top) hwnd = top;
        if (!IsWindow(hwnd) || !IsTopLevelWindow(hwnd)) return;
        if (g_tsfnRestored || g_tsfnChange) {
            WindowEventPayload payload = MakePayload(hwnd);
            auto* heap = new WindowEventPayload(payload);
            if (g_tsfnRestored) g_tsfnRestored.BlockingCall(new WindowEventPayload(payload), [](Env env, Function cb, WindowEventPayload* data){
                Object o = Object::New(env);
                o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("title", String::New(env, data->title));
                o.Set("executablePath", String::New(env, data->exePath));
                o.Set("isVisible", Boolean::New(env, data->isVisible));
                o.Set("type", String::New(env, "restored"));
                cb.Call({ o });
                delete data;
            });
            if (g_tsfnChange) g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                Object o = Object::New(env);
                o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("title", String::New(env, data->title));
                o.Set("executablePath", String::New(env, data->exePath));
                o.Set("isVisible", Boolean::New(env, data->isVisible));
                o.Set("type", String::New(env, "restored"));
                cb.Call({ o });
                delete data;
            }); else delete heap;
        }
        return;
    }
    // System-wide minimize start/end as hints (best-effort)
    if (event == EVENT_SYSTEM_MINIMIZESTART) {
        if ((!IsWindow(hwnd) || !IsTopLevelWindow(hwnd))) {
            HWND fg = GetForegroundWindow();
            if (fg) hwnd = GetAncestor(fg, GA_ROOT);
        }
        if (IsWindow(hwnd) && IsTopLevelWindow(hwnd) && (g_tsfnMinimized || g_tsfnChange)) {
            WindowEventPayload payload = MakePayload(hwnd);
            auto* heap = new WindowEventPayload(payload);
            if (g_tsfnMinimized) g_tsfnMinimized.BlockingCall(new WindowEventPayload(payload), [](Env env, Function cb, WindowEventPayload* data){
                Object o = Object::New(env);
                o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("title", String::New(env, data->title));
                o.Set("executablePath", String::New(env, data->exePath));
                o.Set("isVisible", Boolean::New(env, data->isVisible));
                o.Set("type", String::New(env, "minimized"));
                cb.Call({ o });
                delete data;
            });
            if (g_tsfnChange) g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                Object o = Object::New(env);
                o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("title", String::New(env, data->title));
                o.Set("executablePath", String::New(env, data->exePath));
                o.Set("isVisible", Boolean::New(env, data->isVisible));
                o.Set("type", String::New(env, "minimized"));
                cb.Call({ o });
                delete data;
            }); else delete heap;
        }
        return;
    }
    if (event == EVENT_SYSTEM_MINIMIZEEND) {
        if ((!IsWindow(hwnd) || !IsTopLevelWindow(hwnd))) {
            HWND fg = GetForegroundWindow();
            if (fg) hwnd = GetAncestor(fg, GA_ROOT);
        }
        if (IsWindow(hwnd) && IsTopLevelWindow(hwnd) && (g_tsfnRestored || g_tsfnChange)) {
            WindowEventPayload payload = MakePayload(hwnd);
            auto* heap = new WindowEventPayload(payload);
            if (g_tsfnRestored) g_tsfnRestored.BlockingCall(new WindowEventPayload(payload), [](Env env, Function cb, WindowEventPayload* data){
                Object o = Object::New(env);
                o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("title", String::New(env, data->title));
                o.Set("executablePath", String::New(env, data->exePath));
                o.Set("isVisible", Boolean::New(env, data->isVisible));
                o.Set("type", String::New(env, "restored"));
                cb.Call({ o });
                delete data;
            });
            if (g_tsfnChange) g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                Object o = Object::New(env);
                o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("title", String::New(env, data->title));
                o.Set("executablePath", String::New(env, data->exePath));
                o.Set("isVisible", Boolean::New(env, data->isVisible));
                o.Set("type", String::New(env, "restored"));
                cb.Call({ o });
                delete data;
            }); else delete heap;
        }
        return;
    }
    // For show/hide/minimize/restore, require a real window object (accept WINDOW or CLIENT)
    if (idObject != OBJID_WINDOW && idObject != OBJID_CLIENT) return;
    if (!IsWindow(hwnd) || !IsTopLevelWindow(hwnd)) return;
    if (event == EVENT_OBJECT_HIDE) {
        HWND top = hwnd ? GetAncestor(hwnd, GA_ROOT) : NULL;
        if (top) hwnd = top;
        // Treat as minimized/hidden
        if (g_tsfnMinimized || g_tsfnChange) {
            WindowEventPayload payload = MakePayload(hwnd);
            auto* heap = new WindowEventPayload(payload);
            if (g_tsfnMinimized) g_tsfnMinimized.BlockingCall(new WindowEventPayload(payload), [](Env env, Function cb, WindowEventPayload* data){
                Object o = Object::New(env);
                o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("title", String::New(env, data->title));
                o.Set("executablePath", String::New(env, data->exePath));
                o.Set("isVisible", Boolean::New(env, data->isVisible));
                o.Set("type", String::New(env, "minimized"));
                cb.Call({ o });
                delete data;
            });
            if (g_tsfnChange) g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                Object o = Object::New(env);
                o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("title", String::New(env, data->title));
                o.Set("executablePath", String::New(env, data->exePath));
                o.Set("isVisible", Boolean::New(env, data->isVisible));
                o.Set("type", String::New(env, "minimized"));
                cb.Call({ o });
                delete data;
            }); else delete heap;
        }
        return;
    }
    if (event == EVENT_OBJECT_SHOW) {
        HWND top = hwnd ? GetAncestor(hwnd, GA_ROOT) : NULL;
        if (top) hwnd = top;
        // Treat as restored/shown
        if (g_tsfnRestored || g_tsfnChange) {
            WindowEventPayload payload = MakePayload(hwnd);
            auto* heap = new WindowEventPayload(payload);
            if (g_tsfnRestored) g_tsfnRestored.BlockingCall(new WindowEventPayload(payload), [](Env env, Function cb, WindowEventPayload* data){
                Object o = Object::New(env);
                o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("title", String::New(env, data->title));
                o.Set("executablePath", String::New(env, data->exePath));
                o.Set("isVisible", Boolean::New(env, data->isVisible));
                o.Set("type", String::New(env, "restored"));
                cb.Call({ o });
                delete data;
            });
            if (g_tsfnChange) g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                Object o = Object::New(env);
                o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("title", String::New(env, data->title));
                o.Set("executablePath", String::New(env, data->exePath));
                o.Set("isVisible", Boolean::New(env, data->isVisible));
                o.Set("type", String::New(env, "restored"));
                cb.Call({ o });
                delete data;
            }); else delete heap;
        }
        // continue to also let CREATE/other handlers run if applicable
    }
    if (event == EVENT_OBJECT_CREATE) {
        if (g_tsfnCreated || g_tsfnChange) {
            WindowEventPayload payload = MakePayload(hwnd);
            auto* heap = new WindowEventPayload(payload);
            if (g_tsfnCreated) g_tsfnCreated.BlockingCall(new WindowEventPayload(payload), [](Env env, Function cb, WindowEventPayload* data){
                Object o = Object::New(env);
                o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("title", String::New(env, data->title));
                o.Set("executablePath", String::New(env, data->exePath));
                o.Set("isVisible", Boolean::New(env, data->isVisible));
                o.Set("type", String::New(env, "created"));
                cb.Call({ o });
                delete data;
            });
            if (g_tsfnChange) g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                Object o = Object::New(env);
                o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("title", String::New(env, data->title));
                o.Set("executablePath", String::New(env, data->exePath));
                o.Set("isVisible", Boolean::New(env, data->isVisible));
                o.Set("type", String::New(env, "created"));
                cb.Call({ o });
                delete data;
            }); else delete heap;
        }
    } else if (event == EVENT_OBJECT_DESTROY) {
        if (g_tsfnClosed || g_tsfnChange) {
            // For destroyed windows, title/path may be inaccessible; send minimal info
            WindowEventPayload payload{}; payload.hwnd = hwnd; payload.isVisible = false;
            auto* heap = new WindowEventPayload(payload);
            if (g_tsfnClosed) g_tsfnClosed.BlockingCall(new WindowEventPayload(payload), [](Env env, Function cb, WindowEventPayload* data){
                Object o = Object::New(env);
                o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("title", String::New(env, data->title));
                o.Set("executablePath", String::New(env, data->exePath));
                o.Set("isVisible", Boolean::New(env, data->isVisible));
                o.Set("type", String::New(env, "closed"));
                cb.Call({ o });
                delete data;
            });
            if (g_tsfnChange) g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                Object o = Object::New(env);
                o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                o.Set("title", String::New(env, data->title));
                o.Set("executablePath", String::New(env, data->exePath));
                o.Set("isVisible", Boolean::New(env, data->isVisible));
                o.Set("type", String::New(env, "closed"));
                cb.Call({ o });
                delete data;
            }); else delete heap;
        }
    }
}

static void EnsureHooksInstalled() {
    if (g_hookCreate || g_hookDestroy || g_hookForeground || g_hookShow || g_hookHide || g_hookCloaked || g_hookUncloaked || g_hookMinStart || g_hookMinEnd || g_hookState) return;
    DWORD flags = WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS;
    g_hookCreate = SetWinEventHook(EVENT_OBJECT_CREATE, EVENT_OBJECT_CREATE, NULL, WinEventProcCB, 0, 0, flags);
    g_hookDestroy = SetWinEventHook(EVENT_OBJECT_DESTROY, EVENT_OBJECT_DESTROY, NULL, WinEventProcCB, 0, 0, flags);
    g_hookForeground = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, NULL, WinEventProcCB, 0, 0, flags);
    g_hookShow = SetWinEventHook(EVENT_OBJECT_SHOW, EVENT_OBJECT_SHOW, NULL, WinEventProcCB, 0, 0, flags);
    g_hookHide = SetWinEventHook(EVENT_OBJECT_HIDE, EVENT_OBJECT_HIDE, NULL, WinEventProcCB, 0, 0, flags);
    // Cloak/uncloak and minimize start/end
    g_hookCloaked = SetWinEventHook(EVENT_OBJECT_CLOAKED, EVENT_OBJECT_CLOAKED, NULL, WinEventProcCB, 0, 0, flags);
    g_hookUncloaked = SetWinEventHook(EVENT_OBJECT_UNCLOAKED, EVENT_OBJECT_UNCLOAKED, NULL, WinEventProcCB, 0, 0, flags);
    g_hookMinStart = SetWinEventHook(EVENT_SYSTEM_MINIMIZESTART, EVENT_SYSTEM_MINIMIZESTART, NULL, WinEventProcCB, 0, 0, flags);
    g_hookMinEnd = SetWinEventHook(EVENT_SYSTEM_MINIMIZEEND, EVENT_SYSTEM_MINIMIZEEND, NULL, WinEventProcCB, 0, 0, flags);
    g_hookState = SetWinEventHook(EVENT_OBJECT_STATECHANGE, EVENT_OBJECT_STATECHANGE, NULL, WinEventProcCB, 0, 0, flags);
    // Always start a lightweight fallback poller; it auto-suppresses when hook events are flowing
    StartFallbackEventPoller();
}

static void UninstallHooks() {
    if (g_hookCreate) { UnhookWinEvent(g_hookCreate); g_hookCreate = nullptr; }
    if (g_hookDestroy) { UnhookWinEvent(g_hookDestroy); g_hookDestroy = nullptr; }
    if (g_hookForeground) { UnhookWinEvent(g_hookForeground); g_hookForeground = nullptr; }
    if (g_hookShow) { UnhookWinEvent(g_hookShow); g_hookShow = nullptr; }
    if (g_hookHide) { UnhookWinEvent(g_hookHide); g_hookHide = nullptr; }
    if (g_hookCloaked) { UnhookWinEvent(g_hookCloaked); g_hookCloaked = nullptr; }
    if (g_hookUncloaked) { UnhookWinEvent(g_hookUncloaked); g_hookUncloaked = nullptr; }
    if (g_hookMinStart) { UnhookWinEvent(g_hookMinStart); g_hookMinStart = nullptr; }
    if (g_hookMinEnd) { UnhookWinEvent(g_hookMinEnd); g_hookMinEnd = nullptr; }
    if (g_hookState) { UnhookWinEvent(g_hookState); g_hookState = nullptr; }
    StopFallbackEventPoller();
}

// --------------------- Fallback poller implementation ---------------------
static void EnumTopLevelWindows(std::vector<HWND>& out) {
    out.clear();
    EnumWindows([](HWND hwnd, LPARAM lParam) -> BOOL {
        // Top-level visible and Alt-Tab eligible-ish (fast checks)
        if (!IsWindow(hwnd) || !IsTopLevelWindow(hwnd) || !IsWindowVisible(hwnd)) return TRUE;
        LONG exStyle = GetWindowLongPtr(hwnd, GWL_EXSTYLE);
        if (exStyle & WS_EX_TOOLWINDOW) return TRUE;
        if (GetWindow(hwnd, GW_OWNER)) return TRUE;
        // Basic size check
        RECT rc{}; if (!GetWindowRect(hwnd, &rc)) return TRUE;
        if ((rc.right - rc.left) < 50 || (rc.bottom - rc.top) < 50) return TRUE;
        reinterpret_cast<std::vector<HWND>*>(lParam)->push_back(hwnd);
        return TRUE;
    }, reinterpret_cast<LPARAM>(&out));
}

static void StartFallbackEventPoller() {
    bool expected = false;
    if (!g_eventPollerRunning.compare_exchange_strong(expected, true)) return; // already running
    g_usingFallbackEvents = true;
    g_eventPollerThread = std::thread([](){
        HWND lastFg = nullptr;
        std::unordered_set<HWND> known;
        std::unordered_map<HWND, bool> minimized;
        std::vector<HWND> current;
        while (g_eventPollerRunning.load()) {
            // Suppress if hooks have been active recently
            ULONGLONG nowTick = GetTickCount64();
            ULONGLONG lastHook = g_lastHookEventTick.load();
            bool hooksActive = (lastHook != 0 && (nowTick - lastHook) < 1000);

            // Foreground change detection
            HWND fg = GetForegroundWindow();
            if (fg) {
                HWND top = GetAncestor(fg, GA_ROOT);
                if (top) fg = top;
            }
            if (!hooksActive && fg && fg != lastFg && (g_tsfnFocused || g_tsfnChange)) {
                auto* heap = new WindowEventPayload(MakePayload(fg));
                if (g_tsfnFocused) g_tsfnFocused.BlockingCall(new WindowEventPayload(*heap), [](Env env, Function cb, WindowEventPayload* data){
                    Object o = Object::New(env);
                    o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("title", String::New(env, data->title));
                    o.Set("executablePath", String::New(env, data->exePath));
                    o.Set("isVisible", Boolean::New(env, data->isVisible));
                    o.Set("type", String::New(env, "focused"));
                    cb.Call({ o });
                    delete data;
                });
                if (g_tsfnChange) g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                    Object o = Object::New(env);
                    o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                    o.Set("title", String::New(env, data->title));
                    o.Set("executablePath", String::New(env, data->exePath));
                    o.Set("isVisible", Boolean::New(env, data->isVisible));
                    o.Set("type", String::New(env, "focused"));
                    cb.Call({ o });
                    delete data;
                }); else delete heap;
            }
            lastFg = fg;

            // Window set diff for created/closed + minimized/restored state
            EnumTopLevelWindows(current);
            std::unordered_set<HWND> currSet(current.begin(), current.end());
            // Created
            if (!hooksActive && (g_tsfnCreated || g_tsfnChange)) {
                for (HWND h : current) {
                    if (known.find(h) == known.end()) {
                        known.insert(h);
                        auto payload = MakePayload(h);
                        auto* heap = new WindowEventPayload(payload);
                        if (g_tsfnCreated) g_tsfnCreated.BlockingCall(new WindowEventPayload(payload), [](Env env, Function cb, WindowEventPayload* data){
                            Object o = Object::New(env);
                            o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                            o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                            o.Set("title", String::New(env, data->title));
                            o.Set("executablePath", String::New(env, data->exePath));
                            o.Set("isVisible", Boolean::New(env, data->isVisible));
                            o.Set("type", String::New(env, "created"));
                            cb.Call({ o });
                            delete data;
                        });
                        if (g_tsfnChange) g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                            Object o = Object::New(env);
                            o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                            o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                            o.Set("title", String::New(env, data->title));
                            o.Set("executablePath", String::New(env, data->exePath));
                            o.Set("isVisible", Boolean::New(env, data->isVisible));
                            o.Set("type", String::New(env, "created"));
                            cb.Call({ o });
                            delete data;
                        }); else delete heap;
                    }
                }
            } else {
                // keep known in sync even if no created callback set
                for (HWND h : current) known.insert(h);
            }
            // Closed
            if (!hooksActive && (g_tsfnClosed || g_tsfnChange)) {
                std::vector<HWND> toRemove;
                for (auto h : known) if (currSet.find(h) == currSet.end()) toRemove.push_back(h);
                for (auto h : toRemove) {
                    known.erase(h);
                    WindowEventPayload payload{}; payload.hwnd = h; payload.isVisible = false;
                    auto* heap = new WindowEventPayload(payload);
                    if (g_tsfnClosed) g_tsfnClosed.BlockingCall(new WindowEventPayload(payload), [](Env env, Function cb, WindowEventPayload* data){
                        Object o = Object::New(env);
                        o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                        o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                        o.Set("title", String::New(env, data->title));
                        o.Set("executablePath", String::New(env, data->exePath));
                        o.Set("isVisible", Boolean::New(env, data->isVisible));
                        o.Set("type", String::New(env, "closed"));
                        cb.Call({ o });
                        delete data;
                    });
                    if (g_tsfnChange) g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                        Object o = Object::New(env);
                        o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                        o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                        o.Set("title", String::New(env, data->title));
                        o.Set("executablePath", String::New(env, data->exePath));
                        o.Set("isVisible", Boolean::New(env, data->isVisible));
                        o.Set("type", String::New(env, "closed"));
                        cb.Call({ o });
                        delete data;
                    }); else delete heap;
                }
            } else {
                // purge known of disappeared
                std::vector<HWND> toRemove;
                for (auto h : known) if (currSet.find(h) == currSet.end()) toRemove.push_back(h);
                for (auto h : toRemove) known.erase(h);
            }

            // Minimized/restored transitions
        for (HWND h : current) {
                bool nowMin = IsIconic(h) ? true : false;
                auto it = minimized.find(h);
                bool wasKnown = (it != minimized.end());
                if (!wasKnown) minimized[h] = nowMin;
                else if (it->second != nowMin) {
                    it->second = nowMin;
                    if (!hooksActive && nowMin && (g_tsfnMinimized || g_tsfnChange)) {
                        auto* heap = new WindowEventPayload(MakePayload(h));
                        if (g_tsfnMinimized) g_tsfnMinimized.BlockingCall(new WindowEventPayload(*heap), [](Env env, Function cb, WindowEventPayload* data){
                            Object o = Object::New(env);
                            o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                            o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                            o.Set("title", String::New(env, data->title));
                            o.Set("executablePath", String::New(env, data->exePath));
                            o.Set("isVisible", Boolean::New(env, data->isVisible));
                            o.Set("type", String::New(env, "minimized"));
                            cb.Call({ o });
                            delete data;
                        });
                        if (g_tsfnChange) g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                            Object o = Object::New(env);
                            o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                            o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                            o.Set("title", String::New(env, data->title));
                            o.Set("executablePath", String::New(env, data->exePath));
                            o.Set("isVisible", Boolean::New(env, data->isVisible));
                            o.Set("type", String::New(env, "minimized"));
                            cb.Call({ o });
                            delete data;
                        }); else delete heap;
                    } else if (!hooksActive && !nowMin && (g_tsfnRestored || g_tsfnChange)) {
                        auto* heap = new WindowEventPayload(MakePayload(h));
                        if (g_tsfnRestored) g_tsfnRestored.BlockingCall(new WindowEventPayload(*heap), [](Env env, Function cb, WindowEventPayload* data){
                            Object o = Object::New(env);
                            o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                            o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                            o.Set("title", String::New(env, data->title));
                            o.Set("executablePath", String::New(env, data->exePath));
                            o.Set("isVisible", Boolean::New(env, data->isVisible));
                            o.Set("type", String::New(env, "restored"));
                            cb.Call({ o });
                            delete data;
                        });
                        if (g_tsfnChange) g_tsfnChange.BlockingCall(heap, [](Env env, Function cb, WindowEventPayload* data){
                            Object o = Object::New(env);
                            o.Set("id", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                            o.Set("hwnd", Number::New(env, (uint64_t)(uintptr_t)data->hwnd));
                            o.Set("title", String::New(env, data->title));
                            o.Set("executablePath", String::New(env, data->exePath));
                            o.Set("isVisible", Boolean::New(env, data->isVisible));
                            o.Set("type", String::New(env, "restored"));
                            cb.Call({ o });
                            delete data;
                        }); else delete heap;
                    }
                }
            }

            Sleep(250);
        }
    });
}

static void StopFallbackEventPoller() {
    if (!g_eventPollerRunning.load()) return;
    g_eventPollerRunning.store(false);
    if (g_eventPollerThread.joinable()) {
        try { g_eventPollerThread.join(); } catch (...) {}
    }
    g_usingFallbackEvents = false;
}

// Registered windows mapping and id counter (shared with async workers)
// No longer need a global map of IDs; id is the HWND itself

// Minimal IVirtualDesktopManager Definition, um Header-Abhängigkeiten zu vermeiden
// {AA509086-5CA9-4C25-8F95-589D3C07B48A}
static const CLSID CLSID_VirtualDesktopManager = {0xAA509086, 0x5CA9, 0x4C25, {0x8F, 0x95, 0x58, 0x9D, 0x3C, 0x07, 0xB4, 0x8A}};
// {A5CD92FF-29BE-454C-8D04-D82879FB3F1B}
static const IID IID_IVirtualDesktopManager = {0xA5CD92FF, 0x29BE, 0x454C, {0x8D, 0x04, 0xD8, 0x28, 0x79, 0xFB, 0x3F, 0x1B}};

struct IVirtualDesktopManager : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE IsWindowOnCurrentVirtualDesktop(HWND topLevelWindow, BOOL* onCurrentDesktop) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetWindowDesktopId(HWND topLevelWindow, GUID* desktopId) = 0;
    virtual HRESULT STDMETHODCALLTYPE MoveWindowToDesktop(HWND topLevelWindow, REFGUID desktopId) = 0;
};

// Helper to locate encoder CLSID (e.g., image/png)
int GetEncoderClsid(const WCHAR* format, CLSID* pClsid) {
    UINT num = 0, size = 0;
    Gdiplus::GetImageEncodersSize(&num, &size);
    if (size == 0) return -1;
    std::vector<BYTE> buffer(size);
    Gdiplus::ImageCodecInfo* pImageCodecInfo = reinterpret_cast<Gdiplus::ImageCodecInfo*>(buffer.data());
    if (Gdiplus::GetImageEncoders(num, size, pImageCodecInfo) != Gdiplus::Ok) return -1;
    for (UINT j = 0; j < num; ++j) {
        if (wcscmp(pImageCodecInfo[j].MimeType, format) == 0) {
            *pClsid = pImageCodecInfo[j].Clsid;
            return (int)j;
        }
    }
    return -1;
}

// Base64 Encoding
std::string base64_encode(unsigned char const* bytes_to_encode, unsigned int in_len) {
    std::string ret;
    int i = 0;
    int j = 0;
    unsigned char char_array_3[3];
    unsigned char char_array_4[4];
    
    const std::string chars_to_encode = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

    while (in_len--) {
        char_array_3[i++] = *(bytes_to_encode++);
        if (i == 3) {
            char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
            char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
            char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
            char_array_4[3] = char_array_3[2] & 0x3f;

            for(i = 0; (i < 4); i++)
                ret += chars_to_encode[char_array_4[i]];
            i = 0;
        }
    }

    if (i) {
        for(j = i; j < 3; j++)
            char_array_3[j] = '\0';

        char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
        char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
        char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
        char_array_4[3] = char_array_3[2] & 0x3f;

        for (j = 0; (j < i + 1); j++)
            ret += chars_to_encode[char_array_4[j]];

        while((i++ < 3))
            ret += '=';
    }

    return ret;
}

// BMP Header Strukturen
#pragma pack(push, 1)
struct BMPFileHeader {
    uint16_t bfType;
    uint32_t bfSize;
    uint16_t bfReserved1;
    uint16_t bfReserved2;
    uint32_t bfOffBits;
};

struct BMPInfoHeader {
    uint32_t biSize;
    int32_t biWidth;
    int32_t biHeight;
    uint16_t biPlanes;
    uint16_t biBitCount;
    uint32_t biCompression;
    uint32_t biSizeImage;
    int32_t biXPelsPerMeter;
    int32_t biYPelsPerMeter;
    uint32_t biClrUsed;
    uint32_t biClrImportant;
};
#pragma pack(pop)

// Executable Path ermitteln (mit Fallbacks und minimalen Rechten)
std::string GetExecutablePath(HWND hwnd) {
    DWORD processId = 0;
    GetWindowThreadProcessId(hwnd, &processId);
    if (!processId) return "";

    // Zuerst: minimale Rechte, funktioniert häufig auch bei erhöhten Prozessen
    HANDLE hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, processId);
    if (hProcess) {
        WCHAR pathW[MAX_PATH] = {0};
        DWORD size = MAX_PATH;
        if (QueryFullProcessImageNameW(hProcess, 0, pathW, &size)) {
            std::wstring w(pathW, size);
            CloseHandle(hProcess);
            return WideToUtf8(w);
        }
        CloseHandle(hProcess);
    }

    // Fallback: erweiterte Rechte (kann bei fehlenden Rechten fehlschlagen)
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processId);
    if (hProcess) {
        WCHAR pathW[MAX_PATH] = {0};
        if (GetModuleFileNameExW(hProcess, NULL, pathW, MAX_PATH)) {
            CloseHandle(hProcess);
            return WideToUtf8(std::wstring(pathW));
        }
        CloseHandle(hProcess);
    }
    return "";
}

// Fenster-Titel ermitteln
std::string GetWindowTitle(HWND hwnd) {
    int len = GetWindowTextLengthW(hwnd);
    if (len <= 0) return "";
    std::wstring wtitle;
    wtitle.resize((size_t)len);
    int got = GetWindowTextW(hwnd, &wtitle[0], len + 1);
    if (got <= 0) return "";
    // In case GetWindowTextW wrote a null terminator inside
    if ((size_t)got < wtitle.size()) wtitle.resize((size_t)got);
    return WideToUtf8(wtitle);
}

// Fensterklasse ermitteln
std::string GetWindowClassName(HWND hwnd) {
    WCHAR className[256] = {0};
    int len = GetClassNameW(hwnd, className, (int)(sizeof(className)/sizeof(className[0])));
    if (len > 0) return WideToUtf8(std::wstring(className, len));
    return "";
}

// Versuche, einen Kindfenster-Titel zu finden (für ApplicationFrameWindow u.ä.)
std::string GetFirstChildTitle(HWND hwnd) {
    HWND child = GetWindow(hwnd, GW_CHILD);
    while (child) {
        int len = GetWindowTextLengthW(child);
        if (len > 0) {
            std::string t = GetWindowTitle(child);
            if (!t.empty()) return t;
        }
        child = GetWindow(child, GW_HWNDNEXT);
    }
    return "";
}

// Prüfe, ob ein sichtbares Kindfenster existiert (für AppFrame-Hosts)
bool HasVisibleChildWindow(HWND hwnd) {
    HWND child = GetWindow(hwnd, GW_CHILD);
    while (child) {
        if (IsWindowVisible(child)) return true;
        child = GetWindow(child, GW_HWNDNEXT);
    }
    return false;
}

// Erkenne PowerToys Command Palette anhand Titel oder Prozesspfad
bool IsPowerToysCommandPalette(HWND hwnd) {
    std::string title = GetWindowTitle(hwnd);
    std::string path = GetExecutablePath(hwnd);
    auto toLower = [](std::string s){
        std::transform(s.begin(), s.end(), s.begin(), [](unsigned char c){ return (char)std::tolower(c); });
        return s;
    };
    std::string tl = toLower(title);
    std::string pl = toLower(path);
    if (tl.find("befehlspalette") != std::string::npos || tl.find("command palette") != std::string::npos) return true;
    if (pl.find("microsoft.cmdpal.ui.exe") != std::string::npos) return true;
    return false;
}

// Erkenne Datei-Explorer Fenster
bool IsExplorerWindow(HWND hwnd) {
    std::string cls = GetWindowClassName(hwnd);
    if (_stricmp(cls.c_str(), "CabinetWClass") == 0) return true;
    std::string path = GetExecutablePath(hwnd);
    if (!path.empty()) {
        std::string pl = path;
        std::transform(pl.begin(), pl.end(), pl.begin(), [](unsigned char c){ return (char)std::tolower(c); });
        if (pl.rfind("explorer.exe") != std::string::npos) return true;
    }
    return false;
}

// Erkenne WhatsApp anhand Titel oder Prozesspfad
bool IsWhatsAppWindow(HWND hwnd) {
    std::string title = GetWindowTitle(hwnd);
    if (!title.empty()) {
        std::string tl = title;
        std::transform(tl.begin(), tl.end(), tl.begin(), [](unsigned char c){ return (char)std::tolower(c); });
        if (tl.find("whatsapp") != std::string::npos) return true;
    }
    std::string path = GetExecutablePath(hwnd);
    if (!path.empty()) {
        std::string pl = path;
        std::transform(pl.begin(), pl.end(), pl.begin(), [](unsigned char c){ return (char)std::tolower(c); });
        if (pl.find("\\whatsapp.exe") != std::string::npos || pl.rfind("whatsapp.exe") == pl.size() - std::string("whatsapp.exe").size()) {
            return true;
        }
    }
    return false;
}

// Encode HBITMAP -> PNG (base64 data URL)
std::string BitmapToPngBase64(HBITMAP hBitmap, int width, int height) {
    HDC hdcMem = CreateCompatibleDC(NULL);
    if (!hdcMem) return "data:image/png;base64,";
    
    HBITMAP hOldBitmap = (HBITMAP)SelectObject(hdcMem, hBitmap);

    BITMAPINFO bmi = {};
    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth = width;
    bmi.bmiHeader.biHeight = -height; // Negative für Top-Down
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 24;
    bmi.bmiHeader.biCompression = BI_RGB;

    // Berechne Größe der Bilddaten
    int rowSize = ((width * 3 + 3) / 4) * 4; // 4-Byte-Alignment
    int imageSize = rowSize * height;
    
    // Allokiere Speicher für Bilddaten
    std::vector<unsigned char> pixelData(imageSize);
    
    // Bitmap-Daten abrufen
    if (GetDIBits(hdcMem, hBitmap, 0, height, pixelData.data(), &bmi, DIB_RGB_COLORS) == 0) {
        SelectObject(hdcMem, hOldBitmap);
        DeleteDC(hdcMem);
        return "data:image/png;base64,";
    }

    // Build a GDI+ Bitmap from our 24bpp RGB buffer
    Gdiplus::Bitmap gdiBitmap(width, height, rowSize, PixelFormat24bppRGB, pixelData.data());

    // Save to IStream as PNG
    IStream* stream = nullptr;
    if (FAILED(CreateStreamOnHGlobal(NULL, TRUE, &stream))) {
        SelectObject(hdcMem, hOldBitmap);
        DeleteDC(hdcMem);
        return "data:image/png;base64,";
    }
    CLSID pngClsid{};
    if (GetEncoderClsid(L"image/png", &pngClsid) < 0) {
        stream->Release();
        SelectObject(hdcMem, hOldBitmap);
        DeleteDC(hdcMem);
        return "data:image/png;base64,";
    }
    if (gdiBitmap.Save(stream, &pngClsid, NULL) != Gdiplus::Ok) {
        stream->Release();
        SelectObject(hdcMem, hOldBitmap);
        DeleteDC(hdcMem);
        return "data:image/png;base64,";
    }
    HGLOBAL hMem = NULL;
    if (FAILED(GetHGlobalFromStream(stream, &hMem)) || !hMem) {
        stream->Release();
        SelectObject(hdcMem, hOldBitmap);
        DeleteDC(hdcMem);
        return "data:image/png;base64,";
    }
    SIZE_T sz = GlobalSize(hMem);
    void* pData = GlobalLock(hMem);
    std::string base64 = pData && sz ? base64_encode(reinterpret_cast<unsigned char*>(pData), (unsigned int)sz) : std::string();
    if (pData) GlobalUnlock(hMem);
    stream->Release();

    SelectObject(hdcMem, hOldBitmap);
    DeleteDC(hdcMem);

    return "data:image/png;base64," + base64;
}

// Get size of HBITMAP
static bool GetBitmapSize(HBITMAP hbm, int& w, int& h) {
    if (!hbm) return false;
    BITMAP bm{};
    if (GetObject(hbm, sizeof(bm), &bm) == 0) return false;
    w = bm.bmWidth; h = bm.bmHeight;
    return (w > 0 && h > 0);
}

// Resize an HBITMAP to destW x destH using StretchBlt; returns new HBITMAP (caller deletes)
static HBITMAP ResizeBitmap(HBITMAP src, int srcW, int srcH, int destW, int destH) {
    HDC hdcScreen = GetDC(NULL);
    if (!hdcScreen) return NULL;
    HDC srcDC = CreateCompatibleDC(hdcScreen);
    HDC dstDC = CreateCompatibleDC(hdcScreen);
    HBITMAP dst = CreateCompatibleBitmap(hdcScreen, destW, destH);
    if (!srcDC || !dstDC || !dst) {
        if (dst) DeleteObject(dst);
        if (srcDC) DeleteDC(srcDC);
        if (dstDC) DeleteDC(dstDC);
        ReleaseDC(NULL, hdcScreen);
        return NULL;
    }
    HGDIOBJ oldSrc = SelectObject(srcDC, src);
    HGDIOBJ oldDst = SelectObject(dstDC, dst);
    SetStretchBltMode(dstDC, HALFTONE);
    StretchBlt(dstDC, 0, 0, destW, destH, srcDC, 0, 0, srcW, srcH, SRCCOPY);
    SelectObject(srcDC, oldSrc);
    SelectObject(dstDC, oldDst);
    DeleteDC(srcDC);
    DeleteDC(dstDC);
    ReleaseDC(NULL, hdcScreen);
    return dst;
}

// Icon (HICON) in Base64-BMP umwandeln
std::string IconToBase64(HICON hIcon, int size) {
    if (!hIcon) return "data:image/png;base64,";
    HDC hdc = CreateCompatibleDC(NULL);
    HDC hdcScreen = GetDC(NULL);
    HBITMAP hbm = hdcScreen ? CreateCompatibleBitmap(hdcScreen, size, size) : NULL;
    if (!hdc || !hbm) {
        if (hdc) DeleteDC(hdc);
        if (hbm) DeleteObject(hbm);
        if (hdcScreen) ReleaseDC(NULL, hdcScreen);
        return "data:image/png;base64,";
    }
    HGDIOBJ old = SelectObject(hdc, hbm);
    RECT rc{0,0,size,size};
    HBRUSH brush = (HBRUSH)GetStockObject(WHITE_BRUSH);
    FillRect(hdc, &rc, brush);
    DrawIconEx(hdc, 0, 0, hIcon, size, size, 0, NULL, DI_NORMAL);
    std::string b64 = BitmapToPngBase64(hbm, size, size);
    SelectObject(hdc, old);
    DeleteObject(hbm);
    DeleteDC(hdc);
    ReleaseDC(NULL, hdcScreen);
    return b64;
}

// Get a reasonable HICON for a window (caller may need to DestroyIcon if extracted)
static HICON GetBestIconHandle(HWND hwnd, const std::string& exePath, int desired) {
    HICON hIcon = (HICON)SendMessage(hwnd, WM_GETICON, ICON_BIG, 0);
    if (!hIcon) hIcon = (HICON)SendMessage(hwnd, WM_GETICON, ICON_SMALL2, 0);
    if (!hIcon) hIcon = (HICON)SendMessage(hwnd, WM_GETICON, ICON_SMALL, 0);
    if (!hIcon) hIcon = (HICON)GetClassLongPtr(hwnd, GCLP_HICON);
    if (!hIcon) hIcon = (HICON)GetClassLongPtr(hwnd, GCLP_HICONSM);
    if (!hIcon && !exePath.empty()) {
        std::wstring wpath = Utf8ToWide(exePath);
        HICON extracted = NULL;
        ExtractIconExW(wpath.c_str(), 0, &extracted, NULL, 1);
        if (extracted) return extracted; // caller will DestroyIcon
    }
    return hIcon;
}

// Create a placeholder thumbnail (w x h) with centered app icon
static std::string CreateIconPlaceholderThumbnail(HWND hwnd, const std::string& exePath, int w, int h) {
    HDC hdcScreen = GetDC(NULL);
    if (!hdcScreen) return "data:image/png;base64,";
    HDC hdc = CreateCompatibleDC(hdcScreen);
    HBITMAP hbm = CreateCompatibleBitmap(hdcScreen, w, h);
    std::string result = "data:image/png;base64,";
    if (hdc && hbm) {
        HGDIOBJ old = SelectObject(hdc, hbm);
        // Background: use window color or white
        HBRUSH bg = (HBRUSH)GetSysColorBrush(COLOR_WINDOW);
        RECT rc{0,0,w,h};
        FillRect(hdc, &rc, bg);
        int iconSize = (int)std::min<double>(std::min(w, h) * 0.6, 128.0);
        HICON icon = GetBestIconHandle(hwnd, exePath, iconSize);
        int x = (w - iconSize) / 2;
        int y = (h - iconSize) / 2;
        if (icon) {
            DrawIconEx(hdc, x, y, icon, iconSize, iconSize, 0, NULL, DI_NORMAL);
        }
        result = BitmapToPngBase64(hbm, w, h);
        SelectObject(hdc, old);
        if (icon) DestroyIcon(icon);
    }
    if (hbm) DeleteObject(hbm);
    if (hdc) DeleteDC(hdc);
    ReleaseDC(NULL, hdcScreen);
    return result;
}

// -------------- DWM Thumbnail off-screen capture helpers --------------
static ATOM g_CaptureWndClass = 0;
static const wchar_t* kCaptureWndClassName = L"DwmWin_CaptureWnd";

static LRESULT CALLBACK CaptureWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_ERASEBKGND: return 1; // no flicker
        default: return DefWindowProc(hwnd, msg, wParam, lParam);
    }
}

static bool EnsureCaptureWindowClass() {
    if (g_CaptureWndClass) return true;
    WNDCLASSEXW wc{}; wc.cbSize = sizeof(wc);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = CaptureWndProc;
    wc.hInstance = GetModuleHandleW(NULL);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.lpszClassName = kCaptureWndClassName;
    g_CaptureWndClass = RegisterClassExW(&wc);
    return g_CaptureWndClass != 0;
}

static std::string CaptureWithDwmThumbnail(HWND srcHwnd, int maxWidth, int maxHeight) {
    if (!EnsureCaptureWindowClass()) return "data:image/png;base64,";
    // Create off-screen toolwindow
    HWND dest = CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE, kCaptureWndClassName, L"",
        WS_POPUP, 0, 0, maxWidth, maxHeight, NULL, NULL, GetModuleHandleW(NULL), NULL);
    if (!dest) return "data:image/png;base64,";

    // Register DWM thumbnail
    HTHUMBNAIL hThumb = NULL;
    HRESULT hr = DwmRegisterThumbnail(dest, srcHwnd, &hThumb);
    if (FAILED(hr) || !hThumb) {
        DestroyWindow(dest);
        return "data:image/png;base64,";
    }

    SIZE srcSize{};
    if (FAILED(DwmQueryThumbnailSourceSize(hThumb, &srcSize)) || srcSize.cx <= 0 || srcSize.cy <= 0) {
        DwmUnregisterThumbnail(hThumb);
        DestroyWindow(dest);
        return "data:image/png;base64,";
    }

    // Compute destination rect preserving aspect ratio
    double sx = (double)maxWidth / (double)srcSize.cx;
    double sy = (double)maxHeight / (double)srcSize.cy;
    double s = std::min(sx, sy);
    int outW = std::max(1, (int)std::round(srcSize.cx * s));
    int outH = std::max(1, (int)std::round(srcSize.cy * s));
    RECT destRect{ 0, 0, outW, outH };
    // Ensure window size matches destination and place on-screen at (0,0) without activating
    SetWindowPos(dest, HWND_BOTTOM, 0, 0, outW, outH, SWP_NOACTIVATE);

    DWM_THUMBNAIL_PROPERTIES props{};
    props.dwFlags = DWM_TNP_VISIBLE | DWM_TNP_RECTDESTINATION | DWM_TNP_SOURCECLIENTAREAONLY;
    props.fSourceClientAreaOnly = FALSE;
    props.rcDestination = destRect;
    DwmUpdateThumbnailProperties(hThumb, &props);

    // Show (off-screen) and force paint
    ShowWindow(dest, SW_SHOWNOACTIVATE);
    UpdateWindow(dest);
    DwmFlush();
    // Allow DWM to compose; slightly longer delay for minimized sources
    Sleep(50);

    // Capture the composed thumbnail from the dest window
    HDC hdcWindow = GetDC(dest);
    HDC hdcMem = CreateCompatibleDC(hdcWindow);
    HBITMAP hbm = CreateCompatibleBitmap(hdcWindow, outW, outH);
    std::string result = "data:image/png;base64,";
    if (hdcWindow && hdcMem && hbm) {
        HGDIOBJ old = SelectObject(hdcMem, hbm);
    BitBlt(hdcMem, 0, 0, outW, outH, hdcWindow, 0, 0, SRCCOPY);
        SelectObject(hdcMem, old);
        result = BitmapToPngBase64(hbm, outW, outH);
    }
    if (hbm) DeleteObject(hbm);
    if (hdcMem) DeleteDC(hdcMem);
    if (hdcWindow) ReleaseDC(dest, hdcWindow);

    ShowWindow(dest, SW_HIDE);
    DwmUnregisterThumbnail(hThumb);
    DestroyWindow(dest);
    return result;
}

// Ensure COM for the current thread; returns true if we initialized and must uninit later
static bool EnsureComApartment() {
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (SUCCEEDED(hr)) return true;
    // If already initialized on this thread with different model, do not uninit here
    if (hr == RPC_E_CHANGED_MODE) return false;
    return false;
}

// Get AppUserModelID (AUMID) for a window
static std::wstring GetAumidForWindow(HWND hwnd) {
    std::wstring aumid;
    IPropertyStore* store = nullptr;
    HRESULT hr = SHGetPropertyStoreForWindow(hwnd, IID_PPV_ARGS(&store));
    if (SUCCEEDED(hr) && store) {
        PROPVARIANT pv; PropVariantInit(&pv);
        hr = store->GetValue(PKEY_AppUserModel_ID, &pv);
        if (SUCCEEDED(hr) && pv.vt == VT_LPWSTR && pv.pwszVal) {
            aumid.assign(pv.pwszVal);
        }
        PropVariantClear(&pv);
        store->Release();
    }
    return aumid;
}

// Load UWP icon via shell:Appsfolder using AUMID and return as PNG data URL
static std::string GetUwpIconFromAumid(const std::wstring& aumid, int size) {
    if (aumid.empty()) return "data:image/png;base64,";
    bool didInit = EnsureComApartment();
    std::wstring parsing = L"shell:Appsfolder\\" + aumid;
    IShellItem* psi = nullptr;
    HRESULT hr = SHCreateItemFromParsingName(parsing.c_str(), nullptr, IID_PPV_ARGS(&psi));
    if (FAILED(hr) || !psi) {
        if (didInit) CoUninitialize();
        return "data:image/png;base64,";
    }
    IShellItemImageFactory* pif = nullptr;
    hr = psi->QueryInterface(IID_PPV_ARGS(&pif));
    if (FAILED(hr) || !pif) {
        psi->Release();
        if (didInit) CoUninitialize();
        return "data:image/png;base64,";
    }

    SIZE sz{ size, size };
    HBITMAP hbmp = nullptr;
    hr = pif->GetImage(sz, SIIGBF_RESIZETOFIT | SIIGBF_BIGGERSIZEOK, &hbmp);
    std::string result = "data:image/png;base64,";
    if (SUCCEEDED(hr) && hbmp) {
        result = BitmapToPngBase64(hbmp, size, size);
        DeleteObject(hbmp);
    }
    pif->Release();
    psi->Release();
    if (didInit) CoUninitialize();
    return result;
}

// Icon für ein Fenster abrufen (mit Cache)
std::string GetWindowIconBase64(HWND hwnd, const std::string& exePath, int size = 32) {
    {
        std::lock_guard<std::mutex> lock(g_cacheMutex);
        auto it = g_iconCache.find(hwnd);
        if (it != g_iconCache.end()) return it->second;
    }

    // Special handling for UWP-hosted windows (ApplicationFrameHost, WhatsApp, etc.)
    std::string cls = GetWindowClassName(hwnd);
    std::string clsLower = cls; std::transform(clsLower.begin(), clsLower.end(), clsLower.begin(), [](unsigned char c){ return (char)std::tolower(c); });
    std::string exeLower = exePath; std::transform(exeLower.begin(), exeLower.end(), exeLower.begin(), [](unsigned char c){ return (char)std::tolower(c); });
    bool isAppFrame = (clsLower == "applicationframewindow");
    bool isFrameHost = exeLower.find("\\applicationframehost.exe") != std::string::npos ||
                       (exeLower.size() >= 24 && exeLower.rfind("applicationframehost.exe") == exeLower.size() - strlen("applicationframehost.exe"));
    bool uwpCandidate = isAppFrame || isFrameHost || IsWhatsAppWindow(hwnd);
    if (uwpCandidate) {
        std::wstring aumid = GetAumidForWindow(hwnd);
        if (!aumid.empty()) {
            std::string uwpIcon = GetUwpIconFromAumid(aumid, size);
            if (uwpIcon.size() > strlen("data:image/png;base64,")) {
                std::lock_guard<std::mutex> lock(g_cacheMutex);
                g_iconCache[hwnd] = uwpIcon;
                return uwpIcon;
            }
        }
    }

    HICON hIcon = (HICON)SendMessage(hwnd, WM_GETICON, ICON_SMALL2, 0);
    if (!hIcon) hIcon = (HICON)SendMessage(hwnd, WM_GETICON, ICON_SMALL, 0);
    if (!hIcon) hIcon = (HICON)SendMessage(hwnd, WM_GETICON, ICON_BIG, 0);
    if (!hIcon) hIcon = (HICON)GetClassLongPtr(hwnd, GCLP_HICONSM);
    if (!hIcon) hIcon = (HICON)GetClassLongPtr(hwnd, GCLP_HICON);

    HICON extracted = NULL;
    if (!hIcon && !exePath.empty()) {
        // Let Windows pick the best icon from the file
        std::wstring wpath = Utf8ToWide(exePath);
        ExtractIconExW(wpath.c_str(), 0, NULL, &extracted, 1);
        if (extracted) hIcon = extracted;
    }

    std::string b64 = IconToBase64(hIcon, size);
    if (extracted) DestroyIcon(extracted);
    {
        std::lock_guard<std::mutex> lock(g_cacheMutex);
        g_iconCache[hwnd] = b64;
    }
    return b64;
}

#ifdef ENABLE_WGC
// Windows Graphics Capture (WinRT) one-shot capture for an HWND
// Returns PNG data URL or empty PNG URL if unsupported/failure.
std::string CaptureWindowScreenshotWGC(HWND hwnd, int maxWidth = 200, int maxHeight = 150) {
    using namespace winrt;
    namespace WG = winrt::Windows::Graphics;
    namespace WGC = winrt::Windows::Graphics::Capture;
    namespace WGD = winrt::Windows::Graphics::DirectX;
    namespace WGD11 = winrt::Windows::Graphics::DirectX::Direct3D11;
    namespace WGI = winrt::Windows::Graphics::Imaging;
    namespace WSS = winrt::Windows::Storage::Streams;

    WinrtInitOnce();
    if (!g_winrtInited) return "data:image/png;base64,";

    // Check support
    if (!WGC::GraphicsCaptureSession::IsSupported()) {
        return "data:image/png;base64,";
    }

    // Create D3D11 device with BGRA support
    winrt::com_ptr<ID3D11Device> d3dDevice;
    winrt::com_ptr<ID3D11DeviceContext> d3dContext;
    UINT flags = D3D11_CREATE_DEVICE_BGRA_SUPPORT;
    D3D_FEATURE_LEVEL fl;
    HRESULT hr = D3D11CreateDevice(nullptr, D3D_DRIVER_TYPE_HARDWARE, nullptr, flags, nullptr, 0, D3D11_SDK_VERSION, d3dDevice.put(), &fl, d3dContext.put());
    if (FAILED(hr)) return "data:image/png;base64,";

    // Wrap as WinRT Direct3D device
    winrt::com_ptr<IDXGIDevice> dxgiDevice;
    hr = d3dDevice->QueryInterface(__uuidof(IDXGIDevice), dxgiDevice.put_void());
    if (FAILED(hr) || !dxgiDevice) return "data:image/png;base64,";
    WGD11::IDirect3DDevice winrtDevice{ nullptr };
    {
        winrt::com_ptr<IInspectable> inspectable;
        HRESULT hrDev = CreateDirect3D11DeviceFromDXGIDevice(dxgiDevice.get(), inspectable.put());
        if (FAILED(hrDev) || !inspectable) return "data:image/png;base64,";
    winrtDevice = inspectable.as<WGD11::IDirect3DDevice>();
    }

    // Create GraphicsCaptureItem from HWND via interop
    auto factory = winrt::get_activation_factory<WGC::GraphicsCaptureItem>();
    auto interop = factory.as<IGraphicsCaptureItemInterop>();
    WGC::GraphicsCaptureItem item{ nullptr };
    hr = interop->CreateForWindow(hwnd, winrt::guid_of<WGC::GraphicsCaptureItem>(), reinterpret_cast<void**>(winrt::put_abi(item)));
    if (FAILED(hr) || !item) return "data:image/png;base64,";

    // Determine size
    auto sz = item.Size();
    if (sz.Width <= 0 || sz.Height <= 0) return "data:image/png;base64,";

    // Frame pool and session
    WGC::Direct3D11CaptureFramePool pool = WGC::Direct3D11CaptureFramePool::Create(winrtDevice, WGD::DirectXPixelFormat::B8G8R8A8UIntNormalized, 1, sz);
    WGC::GraphicsCaptureSession session = pool.CreateCaptureSession(item);
    // Optional: Hide capture border if available
    try { session.IsBorderRequired(false); } catch (...) {}
    session.StartCapture();

    WGC::Direct3D11CaptureFrame frame{ nullptr };
    // Poll a few times for a frame
    for (int i = 0; i < 30; ++i) {
        frame = pool.TryGetNextFrame();
        if (frame) break;
        Sleep(10);
    }
    if (!frame) {
        try { session.Close(); pool.Close(); } catch (...) {}
        return "data:image/png;base64,";
    }

    WGD11::IDirect3DSurface surface = frame.Surface();

    // Copy to SoftwareBitmap
    auto op = WGI::SoftwareBitmap::CreateCopyFromSurfaceAsync(surface, WGI::BitmapAlphaMode::Premultiplied);
    WGI::SoftwareBitmap sbmp = nullptr;
    try { sbmp = op.get(); } catch (...) { sbmp = nullptr; }
    if (!sbmp) {
        try { session.Close(); pool.Close(); } catch (...) {}
        return "data:image/png;base64,";
    }

    // Scale if larger than requested
    int32_t w = sbmp.PixelWidth();
    int32_t h = sbmp.PixelHeight();
    double sx = (double)maxWidth / (double)w;
    double sy = (double)maxHeight / (double)h;
    double s = std::min(sx, sy);
    uint32_t outW = (uint32_t)std::max(1, (int)std::round(w * s));
    uint32_t outH = (uint32_t)std::max(1, (int)std::round(h * s));

    WSS::InMemoryRandomAccessStream stream;
    WGI::BitmapEncoder encoder = nullptr;
    try { encoder = WGI::BitmapEncoder::CreateAsync(WGI::BitmapEncoder::PngEncoderId(), stream).get(); } catch (...) { encoder = nullptr; }
    if (!encoder) {
        try { session.Close(); pool.Close(); } catch (...) {}
        return "data:image/png;base64,";
    }
    encoder.SetSoftwareBitmap(sbmp);
    auto transform = encoder.BitmapTransform();
    transform.ScaledWidth(outW);
    transform.ScaledHeight(outH);
    try { encoder.FlushAsync().get(); } catch (...) {}

    // Read back bytes
    uint64_t size = stream.Size();
    std::string base64;
    try {
        auto input = stream.GetInputStreamAt(0);
        WSS::DataReader reader(input);
        reader.LoadAsync((uint32_t)size).get();
        std::vector<uint8_t> bytes((size_t)size);
        reader.ReadBytes(winrt::array_view<uint8_t>(bytes));
        base64 = base64_encode(bytes.data(), (unsigned int)bytes.size());
    } catch (...) {
        base64.clear();
    }

    try { session.Close(); pool.Close(); } catch (...) {}

    return std::string("data:image/png;base64,") + base64;
}
#endif // ENABLE_WGC

// Screenshot eines Fensters erstellen
std::string CaptureWindowScreenshot(HWND hwnd, int maxWidth = 200, int maxHeight = 150) {
    if (!IsWindow(hwnd)) {
        return "data:image/png;base64,";
    }

    auto isGoodPngSize = [](const std::string& data) {
        // Heuristic: a 200x150 PNG is typically > ~8KB base64
        return data.size() > strlen("data:image/png;base64,") + 8000;
    };

    // If minimized, avoid mutating OS iconic thumbnails; use off-screen DWM thumbnail capture instead
    if (IsIconic(hwnd)) {
        // Prefer DwmRegisterThumbnail-based off-screen capture first
        {
            std::string dwmTn = CaptureWithDwmThumbnail(hwnd, maxWidth, maxHeight);
            if (isGoodPngSize(dwmTn)) return dwmTn;
        }
        // If that failed, continue to non-iconic fallbacks below
    }

    // Preferred path for non-minimized windows (optional): Windows Graphics Capture
#ifdef ENABLE_WGC
    if (!IsIconic(hwnd)) {
        std::string wgc = CaptureWindowScreenshotWGC(hwnd, maxWidth, maxHeight);
        if (wgc.size() > strlen("data:image/png;base64,")) {
            return wgc;
        }
    }
#endif

    // Fenstergrößen ermitteln (für minimierte Fenster: normale Größe verwenden)
    RECT windowRect{};
    if (IsIconic(hwnd)) {
        // Already attempted DWM thumbnail above; if we get here, fall back to window placement size and PrintWindow/BitBlt
        WINDOWPLACEMENT wp{}; wp.length = sizeof(WINDOWPLACEMENT);
        if (GetWindowPlacement(hwnd, &wp)) {
            windowRect = wp.rcNormalPosition;
        } else if (!GetWindowRect(hwnd, &windowRect)) {
            return "data:image/png;base64,";
        }
    } else {
        if (!GetWindowRect(hwnd, &windowRect)) {
            return "data:image/png;base64,";
        }
    }

    int windowWidth = windowRect.right - windowRect.left;
    int windowHeight = windowRect.bottom - windowRect.top;

    if (windowWidth <= 0 || windowHeight <= 0) {
        return "data:image/png;base64,";
    }

    // Skalierungsfaktor berechnen
    double scaleX = (double)maxWidth / windowWidth;
    double scaleY = (double)maxHeight / windowHeight;
    double scale = (std::min)(scaleX, scaleY);
    
    int thumbnailWidth = (int)(windowWidth * scale);
    int thumbnailHeight = (int)(windowHeight * scale);

    // Device Contexts erstellen
    HDC hdcWindow = GetDC(hwnd);
    if (!hdcWindow) return "data:image/png;base64,";
    
    HDC hdcMemDC = CreateCompatibleDC(hdcWindow);
    HDC hdcThumbnail = CreateCompatibleDC(hdcWindow);

    if (!hdcMemDC || !hdcThumbnail) {
        if (hdcMemDC) DeleteDC(hdcMemDC);
        if (hdcThumbnail) DeleteDC(hdcThumbnail);
    ReleaseDC(hwnd, hdcWindow);
    return "data:image/png;base64,";
    }

    // Bitmaps erstellen
    HBITMAP hbmScreen = CreateCompatibleBitmap(hdcWindow, windowWidth, windowHeight);
    HBITMAP hbmThumbnail = CreateCompatibleBitmap(hdcWindow, thumbnailWidth, thumbnailHeight);

    if (!hbmScreen || !hbmThumbnail) {
        if (hbmScreen) DeleteObject(hbmScreen);
        if (hbmThumbnail) DeleteObject(hbmThumbnail);
        DeleteDC(hdcMemDC);
        DeleteDC(hdcThumbnail);
    ReleaseDC(hwnd, hdcWindow);
    return "data:image/png;base64,";
    }

    // Original Bitmap auswählen
    HBITMAP hOldBitmap1 = (HBITMAP)SelectObject(hdcMemDC, hbmScreen);
    HBITMAP hOldBitmap2 = (HBITMAP)SelectObject(hdcThumbnail, hbmThumbnail);

    // Fenster-Screenshot aufnehmen - versuche verschiedene Methoden
    BOOL result = FALSE;
    
    // Methode 1: PrintWindow mit PW_RENDERFULLCONTENT
    result = PrintWindow(hwnd, hdcMemDC, PW_RENDERFULLCONTENT);
    
    if (!result) {
        // Methode 2: PrintWindow mit PW_CLIENTONLY
        result = PrintWindow(hwnd, hdcMemDC, PW_CLIENTONLY);
    }
    
    if (!result) {
        // Methode 3: PrintWindow ohne Flags
        result = PrintWindow(hwnd, hdcMemDC, 0);
    }
    
    if (!result) {
        // Methode 4: Fallback - BitBlt vom Desktop
        HDC hdcDesktop = GetDC(NULL);
        if (hdcDesktop) {
            BitBlt(hdcMemDC, 0, 0, windowWidth, windowHeight, 
                   hdcDesktop, windowRect.left, windowRect.top, SRCCOPY);
            ReleaseDC(NULL, hdcDesktop);
        }
    }

    // Thumbnail erstellen (skalieren)
    SetStretchBltMode(hdcThumbnail, HALFTONE);
    StretchBlt(hdcThumbnail, 0, 0, thumbnailWidth, thumbnailHeight,
               hdcMemDC, 0, 0, windowWidth, windowHeight, SRCCOPY);

    // Bitmap -> PNG (base64)
    std::string base64Result = BitmapToPngBase64(hbmThumbnail, thumbnailWidth, thumbnailHeight);

    // Cleanup
    SelectObject(hdcMemDC, hOldBitmap1);
    SelectObject(hdcThumbnail, hOldBitmap2);
    DeleteObject(hbmScreen);
    DeleteObject(hbmThumbnail);
    DeleteDC(hdcMemDC);
    DeleteDC(hdcThumbnail);
    ReleaseDC(hwnd, hdcWindow);

    return base64Result;
}

// Thumbnail aus Cache oder neu erzeugen
std::string GetOrCaptureWindowThumbnail(HWND hwnd, int maxWidth = 200, int maxHeight = 150) {
    RECT rect;
    if (!GetWindowRect(hwnd, &rect)) {
    return "data:image/png;base64,";
    }
    ULONGLONG now = GetTickCount64();
    auto isGoodPng = [](const std::string& data) {
        // Heuristic: a 200x150 PNG is typically > 8KB (base64). Avoid caching tiny title-only captures.
        return data.size() > strlen("data:image/png;base64,") + 8000;
    };
    {
        std::lock_guard<std::mutex> lock(g_cacheMutex);
        auto it = g_thumbCache.find(hwnd);
        if (it != g_thumbCache.end()) {
            const ThumbCacheEntry& e = it->second;
            if (e.w == maxWidth && e.h == maxHeight &&
                e.rect.left == rect.left && e.rect.top == rect.top &&
                e.rect.right == rect.right && e.rect.bottom == rect.bottom &&
                (now - e.ts) < THUMB_TTL_MS) {
                return e.base64;
            }
        }
    }
    // If minimized, prefer returning last good cached image if available to mimic Alt+Tab behavior
    if (IsIconic(hwnd)) {
        std::lock_guard<std::mutex> lock(g_cacheMutex);
        auto it = g_thumbCache.find(hwnd);
        if (it != g_thumbCache.end() && isGoodPng(it->second.base64)) {
            return it->second.base64;
        }
    }

    std::string fresh = CaptureWindowScreenshot(hwnd, maxWidth, maxHeight);
    if (IsIconic(hwnd) && !isGoodPng(fresh)) {
        // Do not overwrite a good cache with a tiny minimized capture
        std::lock_guard<std::mutex> lock(g_cacheMutex);
        auto it = g_thumbCache.find(hwnd);
        if (it != g_thumbCache.end() && isGoodPng(it->second.base64)) {
            return it->second.base64;
        }
        // No good cache exists; create an icon placeholder instead of tiny minimized capture
        // Need the exe path to fetch an icon; recompute cheaply
        std::string exePath = GetExecutablePath(hwnd);
        std::string placeholder = CreateIconPlaceholderThumbnail(hwnd, exePath, maxWidth, maxHeight);
        return placeholder.size() > strlen("data:image/png;base64,") ? placeholder : fresh;
    }
    {
        std::lock_guard<std::mutex> lock(g_cacheMutex);
        g_thumbCache[hwnd] = ThumbCacheEntry{ fresh, rect, now, maxWidth, maxHeight };
    }
    return fresh;
}

// Callback für EnumWindows
// Hilfsfunktionen für Alt-Tab/TaskView-Filtern
static bool IsWindowCloaked(HWND hwnd) {
    BOOL cloaked = FALSE;
    HRESULT hr = DwmGetWindowAttribute(hwnd, DWMWA_CLOAKED, &cloaked, sizeof(cloaked));
    return SUCCEEDED(hr) && cloaked;
}

static bool IsAltTabEligible(HWND hwnd, bool includeAllDesktops) {
    if (!IsWindow(hwnd)) return false;
    // PowerToys Command Palette nie anzeigen
    if (IsPowerToysCommandPalette(hwnd)) return false;
    // Immer nur sichtbare Fenster berücksichtigen
    if (!IsWindowVisible(hwnd)) return false;
    // Minimized windows should still appear in Task View (beide Fälle)

    // Toolwindows nicht anzeigen
    LONG exStyle = GetWindowLongPtr(hwnd, GWL_EXSTYLE);
    if (exStyle & WS_EX_TOOLWINDOW) return false;
    if (exStyle & WS_EX_NOACTIVATE) return false;
    if (exStyle & WS_EX_LAYERED) {
        COLORREF cr = 0; BYTE alpha = 255; DWORD flags = 0;
        if (GetLayeredWindowAttributes(hwnd, &cr, &alpha, &flags)) {
            if ((flags & LWA_ALPHA) && alpha == 0) return false; // komplett transparent
        }
    }

    // Fenster ohne Titel normalerweise nicht in Alt-Tab, außer für UWP-Host/Explorer/WhatsApp
    if (GetWindowTextLengthW(hwnd) <= 0) {
        std::string cls = GetWindowClassName(hwnd);
        if (_stricmp(cls.c_str(), "ApplicationFrameWindow") != 0 &&
            !IsExplorerWindow(hwnd) && !IsWhatsAppWindow(hwnd)) {
            return false;
        }
    }

    // Hinweis: TITLEBARINFO-Unsichtbarkeit nicht als hartes Kriterium verwenden (moderne Apps ohne klassische Titelleiste)

    // Root/Child/Popup-Logik: Alt-Tab zeigt i.d.R. keine Child/Popup-Fenster (außer mit WS_EX_APPWINDOW)
    LONG style = GetWindowLongPtr(hwnd, GWL_STYLE);
    if (style & WS_CHILD) return false;
    BOOL hasOwner = GetWindow(hwnd, GW_OWNER) != NULL;
    bool appWindow = (exStyle & WS_EX_APPWINDOW) != 0;
    if ((style & WS_POPUP) && !appWindow) {
        if (!IsWhatsAppWindow(hwnd)) return false; // erlauben für WhatsApp
    }
    if (!appWindow && hasOwner) return false;

    // Kein harter Check auf Systemmenü/Min/Max: manche modernen Apps haben diese Stile nicht

    // Cloaked ApplicationFrameWindow-Hosts ausschließen (verhindert doppelte/unsichtbare UWP-Hosts)
    std::string cls = GetWindowClassName(hwnd);
    if (_stricmp(cls.c_str(), "ApplicationFrameWindow") == 0) {
        if (IsWindowCloaked(hwnd)) return false;
    }

    // Größe muss sinnvoll sein. Für minimierte Fenster nutze die normale (restored) Größe.
    RECT rect{};
    bool haveRect = false;
    if (IsIconic(hwnd)) {
        WINDOWPLACEMENT wp{};
        wp.length = sizeof(WINDOWPLACEMENT);
        if (GetWindowPlacement(hwnd, &wp)) {
            rect = wp.rcNormalPosition;
            haveRect = true;
        }
    }
    if (!haveRect) {
        if (!GetWindowRect(hwnd, &rect)) return false;
        haveRect = true;
    }
    int rectW = rect.right - rect.left;
    int rectH = rect.bottom - rect.top;
    if (rectW < 50 || rectH < 50) return false;

    return true;
}

struct EnumContext {
    std::vector<WindowInfo>* windows;
    IUnknown* vdmUnknown; // IVirtualDesktopManager, aber als IUnknown um Headerabhängigkeit zu minimieren
    bool includeAllDesktops;
};

static bool IsOnCurrentVirtualDesktop(HWND hwnd, IUnknown* vdmUnknown, bool includeAllDesktops) {
    if (includeAllDesktops) return true;
    if (!vdmUnknown) return true; // Falls nicht verfügbar, nicht filtern
    // Bestimme ein für Alt-Tab relevantes Fenster (RootOwner -> LastActivePopup sichtbar)
    HWND testHwnd = hwnd;
    HWND rootOwner = GetAncestor(hwnd, GA_ROOTOWNER);
    if (rootOwner) {
        HWND walk = NULL;
        HWND tryHwnd = rootOwner;
        while (tryHwnd != walk) {
            walk = tryHwnd;
            tryHwnd = GetLastActivePopup(walk);
            if (IsWindowVisible(tryHwnd)) break;
        }
        if (walk) testHwnd = walk;
    }
    IVirtualDesktopManager* vdm = nullptr;
    HRESULT hrQI = vdmUnknown->QueryInterface(IID_IVirtualDesktopManager, (void**)&vdm);
    if (FAILED(hrQI) || !vdm) return true;
    BOOL onCurrent = TRUE;
    HRESULT hr = vdm->IsWindowOnCurrentVirtualDesktop(testHwnd, &onCurrent);
    vdm->Release();
    if (FAILED(hr)) return true; // Bei Fehler nicht filtern
    if (onCurrent == TRUE) return true;
    // Fallback: Nur wenn das Top-Level-Fenster sichtbar ist.
    if (IsWindowVisible(hwnd) || IsWindowVisible(testHwnd)) return true;
    // Special-Case: WhatsApp soll auf aktuellem Desktop erscheinen, wenn sichtbar
    if (IsWhatsAppWindow(hwnd) && (IsWindowVisible(hwnd) || HasVisibleChildWindow(hwnd))) return true;
    return false;
}

BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam) {
    EnumContext* ctx = reinterpret_cast<EnumContext*>(lParam);

    if (!IsAltTabEligible(hwnd, ctx->includeAllDesktops)) return TRUE;
    if (!IsOnCurrentVirtualDesktop(hwnd, ctx->vdmUnknown, ctx->includeAllDesktops)) return TRUE;

    std::string title = GetWindowTitle(hwnd);
    if (title.empty()) {
        // Fallback für UWP-Hosts: versuche, einen Kindfenster-Titel zu verwenden
        std::string cls = GetWindowClassName(hwnd);
        if (_stricmp(cls.c_str(), "ApplicationFrameWindow") == 0) {
            title = GetFirstChildTitle(hwnd);
        }
        // Fallback für Explorer: versuche Kindtitel; sonst Standardnamen
        if (title.empty() && IsExplorerWindow(hwnd)) {
            title = GetFirstChildTitle(hwnd);
            if (title.empty()) title = "Datei-Explorer";
        }
        // Fallback für WhatsApp: setze Standardtitel
        if (title.empty() && IsWhatsAppWindow(hwnd)) {
            title = "WhatsApp";
        }
        if (title.empty()) return TRUE;
    }

    std::string execPath = GetExecutablePath(hwnd);

    WindowInfo info;
    info.hwnd = hwnd;
    info.title = title;
    info.executablePath = execPath;
    // Sichtbarkeit separat bestimmen (kann auf anderem Desktop false sein)
    // Consider minimized windows as visible for Task View-like behavior
    info.isVisible = IsWindowVisible(hwnd);

    ctx->windows->push_back(info);
    return TRUE;
}

// Hauptfunktion: Alle Fenster mit Thumbnails ermitteln
Value GetWindows(const CallbackInfo& info) {
    Env env = info.Env();
    
    // Option includeAllDesktops ermitteln (bool oder { includeAllDesktops: boolean })
    bool includeAllDesktops = false;
    if (info.Length() >= 1) {
        if (info[0].IsBoolean()) {
            includeAllDesktops = info[0].As<Boolean>().Value();
        } else if (info[0].IsObject()) {
            Object opts = info[0].As<Object>();
            if (opts.Has("includeAllDesktops") && opts.Get("includeAllDesktops").IsBoolean()) {
                includeAllDesktops = opts.Get("includeAllDesktops").As<Boolean>().Value();
            }
        }
    }

    std::vector<WindowInfo> windows;

    // COM initialisieren und VirtualDesktopManager erstellen (optional)
    IUnknown* vdmUnknown = nullptr;
    bool comInitialized = SUCCEEDED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED));
    if (comInitialized) {
        IVirtualDesktopManager* vdm = nullptr;
        HRESULT hr = CoCreateInstance(CLSID_VirtualDesktopManager, nullptr, CLSCTX_ALL,
                                      IID_IVirtualDesktopManager, (void**)&vdm);
        if (SUCCEEDED(hr) && vdm) {
            vdmUnknown = vdm; // behalten als IUnknown; später Release
        }
    }

    EnumContext ctx{ &windows, vdmUnknown, includeAllDesktops };
    EnumWindows(EnumWindowsProc, reinterpret_cast<LPARAM>(&ctx));

    if (vdmUnknown) {
        vdmUnknown->Release();
    }
    if (comInitialized) {
        CoUninitialize();
    }
    
    Array result = Array::New(env);
    
    for (size_t i = 0; i < windows.size(); i++) {
        const WindowInfo& window = windows[i];
        // Use HWND as the public id
        uint64_t hwndId = (uint64_t)(uintptr_t)window.hwnd;

        Object windowObj = Object::New(env);
        windowObj.Set("id", Number::New(env, hwndId));
        windowObj.Set("title", String::New(env, window.title));
        windowObj.Set("executablePath", String::New(env, window.executablePath));
        windowObj.Set("isVisible", Boolean::New(env, window.isVisible));
        windowObj.Set("hwnd", Number::New(env, (uintptr_t)window.hwnd));
        
    // Icon und Screenshot (mit Cache)
    std::string iconBase64 = GetWindowIconBase64(window.hwnd, window.executablePath);
    std::string thumbnailBase64 = GetOrCaptureWindowThumbnail(window.hwnd);
        windowObj.Set("thumbnail", String::New(env, thumbnailBase64));
    windowObj.Set("icon", String::New(env, iconBase64));
        
        result.Set(i, windowObj);
    }
    
    return result;
}

// Thumbnail aktualisieren
Value UpdateThumbnail(const CallbackInfo& info) {
    Env env = info.Env();
    
    if (info.Length() < 1 || !info[0].IsNumber()) {
        TypeError::New(env, "Expected window ID").ThrowAsJavaScriptException();
        return env.Null();
    }
    
    uint64_t windowId = info[0].As<Number>().Int64Value();
    HWND hwnd = (HWND)(uintptr_t)windowId;
    if (!hwnd || !IsWindow(hwnd)) {
        Error::New(env, "Window ID not found or invalid").ThrowAsJavaScriptException();
        return env.Null();
    }

    // Neuen Screenshot erstellen
    std::string newThumbnail = CaptureWindowScreenshot(hwnd);
    auto isGoodPng = [](const std::string& data) {
        return data.size() > strlen("data:image/png;base64,") + 8000;
    };
    // Cache aktualisieren
    RECT rect;
    if (GetWindowRect(hwnd, &rect)) {
        std::lock_guard<std::mutex> lock(g_cacheMutex);
        // Avoid overwriting good cache with tiny minimized captures
        if (!(IsIconic(hwnd) && !isGoodPng(newThumbnail))) {
            g_thumbCache[hwnd] = ThumbCacheEntry{ newThumbnail, rect, GetTickCount64(), 200, 150 };
        }
    }
    
    return String::New(env, newThumbnail);
}

// Fenster wieder öffnen/fokussieren
Value OpenWindow(const CallbackInfo& info) {
    Env env = info.Env();
    
    if (info.Length() < 1 || !info[0].IsNumber()) {
        TypeError::New(env, "Expected window ID").ThrowAsJavaScriptException();
        return env.Null();
    }
    
    uint64_t windowId = info[0].As<Number>().Int64Value();
    HWND hwnd = (HWND)(uintptr_t)windowId;
    if (!hwnd) {
        Error::New(env, "Window ID not found or invalid").ThrowAsJavaScriptException();
        return env.Null();
    }
    
    if (!IsWindow(hwnd)) {
        Error::New(env, "Window no longer exists").ThrowAsJavaScriptException();
        return env.Null();
    }
    
    // Fenster in den Vordergrund bringen
    if (IsIconic(hwnd)) {
        ShowWindow(hwnd, SW_RESTORE);
    }
    
    // Fenster aktivieren
    SetForegroundWindow(hwnd);
    BringWindowToTop(hwnd);
    SetActiveWindow(hwnd);
    
    return Boolean::New(env, true);
}

// ---------------------- Async Promise-based APIs ----------------------
class PromiseWorker : public AsyncWorker {
public:
    explicit PromiseWorker(Napi::Env env) : AsyncWorker(env), deferred(Napi::Promise::Deferred::New(env)) {}
    Napi::Promise GetPromise() { return deferred.Promise(); }
protected:
    Napi::Promise::Deferred deferred;
    void OnError(const Napi::Error& e) override { deferred.Reject(e.Value()); }
};

struct WindowResultEntry {
    HWND hwnd{};
    std::string title;
    std::string executablePath;
    bool isVisible{};
    std::string thumbnail;
    std::string icon;
};

class GetWindowsAsyncWorker : public PromiseWorker {
public:
    GetWindowsAsyncWorker(Napi::Env env, bool includeAll)
        : PromiseWorker(env), includeAllDesktops(includeAll) {}

    void Execute() override {
        // Similar to GetWindows but performed on worker thread
        // COM init for this thread
        bool comInitialized = SUCCEEDED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED));
        IUnknown* vdmUnknown = nullptr;
        if (comInitialized) {
            IVirtualDesktopManager* vdm = nullptr;
            HRESULT hr = CoCreateInstance(CLSID_VirtualDesktopManager, nullptr, CLSCTX_ALL,
                                          IID_IVirtualDesktopManager, (void**)&vdm);
            if (SUCCEEDED(hr) && vdm) {
                vdmUnknown = vdm;
            }
        }

        std::vector<WindowInfo> windows;
        EnumContext ctx{ &windows, vdmUnknown, includeAllDesktops };
        EnumWindows(EnumWindowsProc, reinterpret_cast<LPARAM>(&ctx));

        if (vdmUnknown) vdmUnknown->Release();
        if (comInitialized) CoUninitialize();

        // Build results and capture icon/thumbnail (may be expensive)
        for (const auto& w : windows) {
            WindowResultEntry e;
            e.hwnd = w.hwnd;
            e.title = w.title;
            e.executablePath = w.executablePath;
            e.isVisible = w.isVisible;
            e.icon = GetWindowIconBase64(w.hwnd, w.executablePath);
            e.thumbnail = GetOrCaptureWindowThumbnail(w.hwnd);
            results.push_back(std::move(e));
        }
    }

    void OnOK() override {
        Napi::Env env = this->Env();
        Array arr = Array::New(env, results.size());
        for (size_t i = 0; i < results.size(); ++i) {
            // Use HWND value as id
            uint64_t id = (uint64_t)(uintptr_t)results[i].hwnd;
            Object o = Object::New(env);
            o.Set("id", Number::New(env, id));
            o.Set("title", String::New(env, results[i].title));
            o.Set("executablePath", String::New(env, results[i].executablePath));
            o.Set("isVisible", Boolean::New(env, results[i].isVisible));
            o.Set("hwnd", Number::New(env, (uintptr_t)results[i].hwnd));
            o.Set("thumbnail", String::New(env, results[i].thumbnail));
            o.Set("icon", String::New(env, results[i].icon));
            arr.Set(i, o);
        }
        deferred.Resolve(arr);
    }

private:
    bool includeAllDesktops;
    std::vector<WindowResultEntry> results;
};

class UpdateThumbnailAsyncWorker : public PromiseWorker {
public:
    UpdateThumbnailAsyncWorker(Napi::Env env, uint64_t id)
        : PromiseWorker(env), windowId(id) {}

    void Execute() override {
    // Interpret windowId as HWND directly
    HWND hwndLocal = (HWND)(uintptr_t)windowId;
        if (!hwndLocal || !IsWindow(hwndLocal)) {
            SetError("Window ID not found or invalid");
            return;
        }
        // Capture thumbnail
        thumbnail = CaptureWindowScreenshot(hwndLocal);
        // Update cache meta
        RECT rect;
        if (GetWindowRect(hwndLocal, &rect)) {
            std::lock_guard<std::mutex> lock(g_cacheMutex);
            g_thumbCache[hwndLocal] = ThumbCacheEntry{ thumbnail, rect, GetTickCount64(), 200, 150 };
        }
    }

    void OnOK() override {
        deferred.Resolve(String::New(this->Env(), thumbnail));
    }

private:
    uint64_t windowId;
    std::string thumbnail;
};

class OpenWindowAsyncWorker : public PromiseWorker {
public:
    OpenWindowAsyncWorker(Napi::Env env, uint64_t id)
        : PromiseWorker(env), windowId(id) {}

    void Execute() override {
    HWND hwndLocal = (HWND)(uintptr_t)windowId;
        if (!hwndLocal || !IsWindow(hwndLocal)) {
            SetError("Window ID not found or invalid");
            return;
        }
        if (IsIconic(hwndLocal)) {
            ShowWindow(hwndLocal, SW_RESTORE);
        }
        SetForegroundWindow(hwndLocal);
        BringWindowToTop(hwndLocal);
        SetActiveWindow(hwndLocal);
        success = true;
    }

    void OnOK() override { deferred.Resolve(Boolean::New(this->Env(), success)); }

private:
    uint64_t windowId;
    bool success{false};
};

Value GetWindowsAsync(const CallbackInfo& info) {
    Env env = info.Env();
    bool includeAllDesktops = false;
    if (info.Length() >= 1) {
        if (info[0].IsBoolean()) includeAllDesktops = info[0].As<Boolean>().Value();
        else if (info[0].IsObject()) {
            Object opts = info[0].As<Object>();
            if (opts.Has("includeAllDesktops") && opts.Get("includeAllDesktops").IsBoolean()) {
                includeAllDesktops = opts.Get("includeAllDesktops").As<Boolean>().Value();
            }
        }
    }
    auto* worker = new GetWindowsAsyncWorker(env, includeAllDesktops);
    Promise promise = worker->GetPromise();
    worker->Queue();
    return promise;
}

Value UpdateThumbnailAsync(const CallbackInfo& info) {
    Env env = info.Env();
    if (info.Length() < 1 || !info[0].IsNumber()) {
        TypeError::New(env, "Expected window ID").ThrowAsJavaScriptException();
        return env.Null();
    }
    uint64_t id = info[0].As<Number>().Int64Value();
    auto* worker = new UpdateThumbnailAsyncWorker(env, id);
    Promise promise = worker->GetPromise();
    worker->Queue();
    return promise;
}

Value OpenWindowAsync(const CallbackInfo& info) {
    Env env = info.Env();
    if (info.Length() < 1 || !info[0].IsNumber()) {
        TypeError::New(env, "Expected window ID").ThrowAsJavaScriptException();
        return env.Null();
    }
    uint64_t id = info[0].As<Number>().Int64Value();
    auto* worker = new OpenWindowAsyncWorker(env, id);
    Promise promise = worker->GetPromise();
    worker->Queue();
    return promise;
}

Object Init(Env env, Object exports) {
    if (!g_gdiplusToken) {
        Gdiplus::GdiplusStartupInput si;
        if (Gdiplus::GdiplusStartup(&g_gdiplusToken, &si, nullptr) == Gdiplus::Ok) {
            napi_env ne = env;
            napi_add_env_cleanup_hook(ne, GdiplusCleanup, nullptr);
        }
    }
#ifdef ENABLE_WGC
    if (!g_winrtInited) {
        WinrtInitOnce();
        if (g_winrtInited) {
            napi_env ne = env;
            napi_add_env_cleanup_hook(ne, WinrtCleanup, nullptr);
        }
    }
#endif
    exports.Set("getWindows", Function::New(env, GetWindows));
    exports.Set("updateThumbnail", Function::New(env, UpdateThumbnail));
    exports.Set("openWindow", Function::New(env, OpenWindow));
    // Async variants (Promise-based)
    exports.Set("getWindowsAsync", Function::New(env, GetWindowsAsync));
    exports.Set("updateThumbnailAsync", Function::New(env, UpdateThumbnailAsync));
    exports.Set("openWindowAsync", Function::New(env, OpenWindowAsync));

    // Event registration: onWindowCreated/onWindowClosed/onWindowFocused
    exports.Set("onWindowCreated", Function::New(env, [](const CallbackInfo& info){
        Env e = info.Env();
        if (info.Length() < 1 || !info[0].IsFunction()) {
            TypeError::New(e, "Expected callback function").ThrowAsJavaScriptException();
            return;
        }
        if (!g_tsfnCreated) {
            g_tsfnCreated = ThreadSafeFunction::New(e, info[0].As<Function>(), "win-created", 0, 1);
        } else {
            // Replace callback: release old and set new
            g_tsfnCreated.Release();
            g_tsfnCreated = ThreadSafeFunction::New(e, info[0].As<Function>(), "win-created", 0, 1);
        }
        EnsureHooksInstalled();
    }));

    exports.Set("onWindowClosed", Function::New(env, [](const CallbackInfo& info){
        Env e = info.Env();
        if (info.Length() < 1 || !info[0].IsFunction()) {
            TypeError::New(e, "Expected callback function").ThrowAsJavaScriptException();
            return;
        }
        if (!g_tsfnClosed) {
            g_tsfnClosed = ThreadSafeFunction::New(e, info[0].As<Function>(), "win-closed", 0, 1);
        } else {
            g_tsfnClosed.Release();
            g_tsfnClosed = ThreadSafeFunction::New(e, info[0].As<Function>(), "win-closed", 0, 1);
        }
        EnsureHooksInstalled();
    }));

    exports.Set("onWindowFocused", Function::New(env, [](const CallbackInfo& info){
        Env e = info.Env();
        if (info.Length() < 1 || !info[0].IsFunction()) {
            TypeError::New(e, "Expected callback function").ThrowAsJavaScriptException();
            return;
        }
        if (!g_tsfnFocused) {
            g_tsfnFocused = ThreadSafeFunction::New(e, info[0].As<Function>(), "win-focused", 0, 1);
        } else {
            g_tsfnFocused.Release();
            g_tsfnFocused = ThreadSafeFunction::New(e, info[0].As<Function>(), "win-focused", 0, 1);
        }
        EnsureHooksInstalled();
    }));

    // Minimized/Restored events
    exports.Set("onWindowMinimized", Function::New(env, [](const CallbackInfo& info){
        Env e = info.Env();
        if (info.Length() < 1 || !info[0].IsFunction()) {
            TypeError::New(e, "Expected callback function").ThrowAsJavaScriptException();
            return;
        }
        if (!g_tsfnMinimized) {
            g_tsfnMinimized = ThreadSafeFunction::New(e, info[0].As<Function>(), "win-minimized", 0, 1);
        } else {
            g_tsfnMinimized.Release();
            g_tsfnMinimized = ThreadSafeFunction::New(e, info[0].As<Function>(), "win-minimized", 0, 1);
        }
        EnsureHooksInstalled();
    }));

    exports.Set("onWindowRestored", Function::New(env, [](const CallbackInfo& info){
        Env e = info.Env();
        if (info.Length() < 1 || !info[0].IsFunction()) {
            TypeError::New(e, "Expected callback function").ThrowAsJavaScriptException();
            return;
        }
        if (!g_tsfnRestored) {
            g_tsfnRestored = ThreadSafeFunction::New(e, info[0].As<Function>(), "win-restored", 0, 1);
        } else {
            g_tsfnRestored.Release();
            g_tsfnRestored = ThreadSafeFunction::New(e, info[0].As<Function>(), "win-restored", 0, 1);
        }
        EnsureHooksInstalled();
    }));

    // Unified: onWindowChange
    exports.Set("onWindowChange", Function::New(env, [](const CallbackInfo& info){
        Env e = info.Env();
        if (info.Length() < 1 || !info[0].IsFunction()) {
            TypeError::New(e, "Expected callback function").ThrowAsJavaScriptException();
            return;
        }
        if (!g_tsfnChange) {
            g_tsfnChange = ThreadSafeFunction::New(e, info[0].As<Function>(), "win-change", 0, 1);
        } else {
            g_tsfnChange.Release();
            g_tsfnChange = ThreadSafeFunction::New(e, info[0].As<Function>(), "win-change", 0, 1);
        }
        EnsureHooksInstalled();
    }));

    exports.Set("stopWindowEvents", Function::New(env, [](const CallbackInfo& info){
        (void)info;
        UninstallHooks();
        if (g_tsfnCreated) { g_tsfnCreated.Release(); g_tsfnCreated = ThreadSafeFunction(); }
        if (g_tsfnClosed) { g_tsfnClosed.Release(); g_tsfnClosed = ThreadSafeFunction(); }
        if (g_tsfnFocused) { g_tsfnFocused.Release(); g_tsfnFocused = ThreadSafeFunction(); }
    if (g_tsfnMinimized) { g_tsfnMinimized.Release(); g_tsfnMinimized = ThreadSafeFunction(); }
    if (g_tsfnRestored) { g_tsfnRestored.Release(); g_tsfnRestored = ThreadSafeFunction(); }
    if (g_tsfnChange) { g_tsfnChange.Release(); g_tsfnChange = ThreadSafeFunction(); }
    }));
    exports.Set("isUsingFallbackEvents", Function::New(env, [](const CallbackInfo& info){
        (void)info; return Boolean::New(info.Env(), g_usingFallbackEvents.load());
    }));
    return exports;
}

NODE_API_MODULE(dwm_windows, Init)