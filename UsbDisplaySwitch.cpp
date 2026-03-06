#define UNICODE
#define _UNICODE

#include <windows.h>
#include <dbt.h>
#include <strsafe.h>
#include <atomic>
#include <string>
#include <algorithm>
#include <Wtsapi32.h>
#include <vector>
#include <cwctype>      // towupper
#include <cstdio>
#include <cstdlib>      // wcstoul

#pragma comment(lib, "Advapi32.lib")
#pragma comment(lib, "User32.lib")
#pragma comment(lib, "Wtsapi32.lib")

// ============================================================================
//                              CONFIGURATION
// ============================================================================

// Windows service name (must match service registration name)
static const wchar_t *kServiceName = L"MonitorInputSwitch";

// Target HID VID/PID needle (case-insensitive substring match against device interface path)
static const wchar_t *kTargetVidPid = L"VID_24AE&PID_4056";

// Debounce / cooldown time after a trigger (milliseconds)
static constexpr DWORD kCooldownMs = 1500;

// ---------- DDC/CI target monitor + VCP setting ----------
// Which display device to control (MONITORINFOEX::szDevice), e.g. "\\.\DISPLAY1"
static const wchar_t *kTargetDisplayDevice = L"\\\\.\\DISPLAY1";

// VCP code to set.
// Example: 0x60 is "Input Source" on many monitors (vendor-specific).
static constexpr BYTE kVcpCode = 0x60;

// Value used when the USB HID device is plugged in (arrival).
static constexpr DWORD kVcpValueOnArrival = 18;

// Value used when the USB HID device is unplugged (remove).
// IMPORTANT: This is conditionally compiled. If you want removal switching,
// define USB_REMOVE_VCP_VALUE to a number (e.g. 17). If you want to disable
// removal switching, comment it out or remove the definition.
//
// Example (enable):
//   #define USB_REMOVE_VCP_VALUE 17
//
// Example (disable):
//   // #define USB_REMOVE_VCP_VALUE 17
//
#define USB_REMOVE_VCP_VALUE 17

// Self-spawn flag: the service spawns this same EXE in the active user session
// with this argument to perform DDC/CI and then exit.
static const wchar_t *kSelfSetterFlag = L"--set-vcp";

// ============================================================================


// HID interface class GUID (GUID_DEVINTERFACE_HID)
static constexpr GUID GUID_DEVINTERFACE_HID_ =
        {0x4D1E55B2, 0xF16F, 0x11CF, {0x88, 0xCB, 0x00, 0x11, 0x11, 0x00, 0x00, 0x30}};

// Global service state
static SERVICE_STATUS_HANDLE g_svcStatusHandle = nullptr;
static HANDLE g_stopEvent = nullptr;
static HANDLE g_workerThread = nullptr;

// Debounce timestamp: next allowed trigger time in GetTickCount64() milliseconds
static std::atomic<ULONGLONG> g_nextAllowedMs{0};

// ============================================================================
//                               EVENT LOGGING
// ============================================================================
//
// Minimal Windows Event Log helper. If it fails, we silently ignore errors.
// This keeps the service robust even when Event Log registration is not available.
//
static void LogEvent(WORD type, const wchar_t *msg) {
    HANDLE h = RegisterEventSourceW(nullptr, kServiceName);
    if (!h) return;
    const wchar_t *strs[1] = {msg};
    ReportEventW(h, type, 0, 0x1000, nullptr, 1, 0, strs, nullptr);
    DeregisterEventSource(h);
}

// Helper: log formatted event messages.
static void LogEventFmt(WORD type, const wchar_t *fmt, DWORD a) {
    wchar_t buf[256];
    StringCchPrintfW(buf, _countof(buf), fmt, a);
    LogEvent(type, buf);
}

// ============================================================================
//                               STRING HELPERS
// ============================================================================

// Uppercase a wide string using towupper (Unicode-safe).
static std::wstring ToUpper(std::wstring s) {
    std::transform(s.begin(), s.end(), s.begin(),
                   [](wchar_t c) { return static_cast<wchar_t>(std::towupper(c)); });
    return s;
}

// Case-insensitive substring match.
// We uppercase the haystack once and search for an already-uppercased needle.
static bool ContainsCaseInsensitive(const std::wstring &haystack, const std::wstring &needleUpper) {
    std::wstring h = ToUpper(haystack);
    return h.find(needleUpper) != std::wstring::npos;
}

// Get full path of current executable.
static std::wstring GetSelfExePath() {
    wchar_t path[MAX_PATH]{};
    GetModuleFileNameW(nullptr, path, MAX_PATH);
    return std::wstring(path);
}

// ============================================================================
//                 DDC/CI VIA RUNTIME LOADING OF Dxva2.dll
// ============================================================================
//
// Why runtime loading?
// - Avoid compile-time dependency on Windows SDK headers that may hide declarations
//   based on _WIN32_WINNT / WINVER.
// - Avoid link-time dependency on Dxva2.lib.
// - Works as long as Dxva2.dll is present at runtime (Vista+).
//
// Exported functions used from Dxva2.dll:
//   - GetNumberOfPhysicalMonitorsFromHMONITOR
//   - GetPhysicalMonitorsFromHMONITOR
//   - DestroyPhysicalMonitors
//   - SetVCPFeature
//
// Notes:
// - Many monitors require enabling "DDC/CI" in OSD.
// - VCP code/value mapping is vendor-specific.
//
// ----------------------------------------------------------------------------

#ifndef PHYSICAL_MONITOR_DESCRIPTION_SIZE
#define PHYSICAL_MONITOR_DESCRIPTION_SIZE 128
#endif

// Minimal compatible version of PHYSICAL_MONITOR (layout must match Dxva2)
typedef struct _PHYSICAL_MONITOR_MIN {
    HANDLE hPhysicalMonitor;
    WCHAR szPhysicalMonitorDescription[PHYSICAL_MONITOR_DESCRIPTION_SIZE];
} PHYSICAL_MONITOR_MIN, *LPPHYSICAL_MONITOR_MIN;

// Function pointer typedefs matching Dxva2 exports
typedef BOOL (WINAPI *PFN_GetNumberOfPhysicalMonitorsFromHMONITOR)(
    HMONITOR hMonitor, LPDWORD pdwNumberOfPhysicalMonitors);

typedef BOOL (WINAPI *PFN_GetPhysicalMonitorsFromHMONITOR)(
    HMONITOR hMonitor, DWORD dwPhysicalMonitorArraySize, LPPHYSICAL_MONITOR_MIN pPhysicalMonitorArray);

typedef BOOL (WINAPI *PFN_DestroyPhysicalMonitors)(
    DWORD dwPhysicalMonitorArraySize, LPPHYSICAL_MONITOR_MIN pPhysicalMonitorArray);

typedef BOOL (WINAPI *PFN_SetVCPFeature)(
    HANDLE hMonitor, BYTE bVCPCode, DWORD dwNewValue);

// Find HMONITOR by MONITORINFOEX::szDevice (e.g. "\\.\DISPLAY1")
struct FindMonitorCtx {
    std::wstring targetDeviceUpper;
    HMONITOR found = nullptr;
};

static BOOL CALLBACK EnumMonitorsProc(HMONITOR hMon, HDC, LPRECT, LPARAM lParam) {
    auto *ctx = reinterpret_cast<FindMonitorCtx *>(lParam);
    if (!ctx) return TRUE;

    MONITORINFOEXW mi{};
    mi.cbSize = sizeof(mi);

    if (!GetMonitorInfoW(hMon, &mi))
        return TRUE;

    if (ToUpper(mi.szDevice) == ctx->targetDeviceUpper) {
        ctx->found = hMon;
        return FALSE; // stop enumeration
    }

    return TRUE;
}

// Set VCP on a specific HMONITOR using runtime-loaded Dxva2 APIs.
static bool SetVcpOnHmonitor_RuntimeDxva2(HMONITOR hMon, DWORD vcpValue) {
    if (!hMon) return false;

    HMODULE dxva2 = LoadLibraryW(L"Dxva2.dll");
    if (!dxva2) {
        LogEvent(EVENTLOG_ERROR_TYPE, L"DDC/CI: LoadLibraryW(Dxva2.dll) failed.");
        return false;
    }

    auto pGetNumber = reinterpret_cast<PFN_GetNumberOfPhysicalMonitorsFromHMONITOR>(GetProcAddress(
        dxva2, "GetNumberOfPhysicalMonitorsFromHMONITOR"));
    auto pGetPhysical = reinterpret_cast<PFN_GetPhysicalMonitorsFromHMONITOR>(GetProcAddress(
        dxva2, "GetPhysicalMonitorsFromHMONITOR"));
    auto pDestroy = reinterpret_cast<PFN_DestroyPhysicalMonitors>(GetProcAddress(dxva2, "DestroyPhysicalMonitors"));
    auto pSetVcp = reinterpret_cast<PFN_SetVCPFeature>(GetProcAddress(dxva2, "SetVCPFeature"));

    if (!pGetNumber || !pGetPhysical || !pDestroy || !pSetVcp) {
        LogEvent(EVENTLOG_ERROR_TYPE, L"DDC/CI: Missing required exports in Dxva2.dll.");
        FreeLibrary(dxva2);
        return false;
    }

    DWORD count = 0;
    if (!pGetNumber(hMon, &count) || count == 0) {
        LogEvent(EVENTLOG_ERROR_TYPE, L"DDC/CI: No physical monitors found for HMONITOR.");
        FreeLibrary(dxva2);
        return false;
    }

    std::vector<PHYSICAL_MONITOR_MIN> phys(count);
    if (!pGetPhysical(hMon, count, phys.data())) {
        LogEvent(EVENTLOG_ERROR_TYPE, L"DDC/CI: GetPhysicalMonitorsFromHMONITOR failed.");
        FreeLibrary(dxva2);
        return false;
    }

    bool ok = false;

    // Apply to the first physical monitor.
    // If your HMONITOR maps to multiple physical monitors, you can loop over all.
    if (pSetVcp(phys[0].hPhysicalMonitor, kVcpCode, vcpValue)) {
        ok = true;
    } else {
        DWORD err = GetLastError();
        wchar_t buf[256];
        StringCchPrintfW(buf, _countof(buf),
                         L"DDC/CI: SetVCPFeature failed. GetLastError=%lu", err);
        LogEvent(EVENTLOG_ERROR_TYPE, buf);
    }

    // Always destroy the physical monitor handles to avoid handle leaks.
    pDestroy(count, phys.data());

    FreeLibrary(dxva2);
    return ok;
}

// Public helper: set VCP on a given display device name, e.g. "\\.\DISPLAY1"
static bool SetVcpOnDisplayDevice_RuntimeDxva2(const wchar_t *displayDevice, DWORD vcpValue) {
    if (!displayDevice || !displayDevice[0]) return false;

    FindMonitorCtx ctx;
    ctx.targetDeviceUpper = ToUpper(displayDevice);

    EnumDisplayMonitors(nullptr, nullptr, EnumMonitorsProc, reinterpret_cast<LPARAM>(&ctx));

    if (!ctx.found) {
        LogEvent(EVENTLOG_ERROR_TYPE, L"DDC/CI: Failed to locate target display device (HMONITOR).");
        return false;
    }

    return SetVcpOnHmonitor_RuntimeDxva2(ctx.found, vcpValue);
}

// ============================================================================
//           SPAWN THIS EXE IN ACTIVE USER SESSION (SERVICE -> USER)
// ============================================================================
//
// Windows services run in Session 0. DDC/CI and monitor enumeration can be more
// reliable when executed in the interactive user's session.
// Therefore, on HID arrival/removal we spawn this same executable with
//   "--set-vcp <value>"
// in the active console session via CreateProcessAsUserW.
//
// The child process performs DDC/CI and exits immediately.
//
static bool RunSelfSetterInActiveSession(DWORD vcpValue) {
    DWORD sessionId = WTSGetActiveConsoleSessionId();
    if (sessionId == 0xFFFFFFFF)
        return false;

    HANDLE userToken = nullptr;
    if (!WTSQueryUserToken(sessionId, &userToken))
        return false;

    HANDLE primaryToken = nullptr;
    if (!DuplicateTokenEx(
        userToken,
        MAXIMUM_ALLOWED,
        nullptr,
        SecurityIdentification,
        TokenPrimary,
        &primaryToken)) {
        CloseHandle(userToken);
        return false;
    }
    CloseHandle(userToken);

    std::wstring selfExe = GetSelfExePath();

    // Build: "<this.exe>" --set-vcp <decimalValue>
    wchar_t valBuf[32];
    StringCchPrintfW(valBuf, _countof(valBuf), L"%lu", vcpValue);

    std::wstring cmd = L"\"" + selfExe + L"\" " + std::wstring(kSelfSetterFlag) + L" " + valBuf;

    // CreateProcessAsUserW requires a writable command line buffer.
    std::vector<wchar_t> cmdBuf(cmd.begin(), cmd.end());
    cmdBuf.push_back(L'\0');

    STARTUPINFOW si{};
    si.cb = sizeof(si);
    si.lpDesktop = const_cast<LPWSTR>(L"winsta0\\default");

    PROCESS_INFORMATION pi{};

    BOOL ok = CreateProcessAsUserW(
        primaryToken,
        nullptr, // applicationName optional when included in cmd line
        cmdBuf.data(), // writable command line
        nullptr,
        nullptr,
        FALSE,
        CREATE_NO_WINDOW,
        nullptr,
        nullptr,
        &si,
        &pi);

    DWORD err = ok ? 0 : GetLastError();
    CloseHandle(primaryToken);

    if (!ok) {
        LogEventFmt(EVENTLOG_ERROR_TYPE,
                    L"CreateProcessAsUserW (self setter) failed. GetLastError=%lu", err);
        return false;
    }

    // Fire-and-forget: child process does the work and exits.
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return true;
}

// ============================================================================
//                   HID ARRIVAL/REMOVAL LISTENER WINDOW
// ============================================================================
//
// We create a hidden window and register for HID interface notifications.
// When WM_DEVICECHANGE indicates arrival/removal, we check the device interface
// path for VID/PID match and then trigger the self setter (debounced).
//
struct WindowState {
    HDEVNOTIFY notify = nullptr;
    std::wstring targetUpper;
};

static bool g_consoleMode = false;

// Shared helper to handle arrival/removal using the same debounce gate.
static void HandleMatchedTrigger(const wchar_t *reason, DWORD vcpValue) {
    ULONGLONG now = GetTickCount64();
    ULONGLONG allowed = g_nextAllowedMs.load(std::memory_order_relaxed);

    if (now < allowed)
        return;

    g_nextAllowedMs.store(now + kCooldownMs, std::memory_order_relaxed);

    wchar_t msg[256];
    StringCchPrintfW(msg, _countof(msg), L"%s. value=%lu", reason, vcpValue);
    LogEvent(EVENTLOG_INFORMATION_TYPE, msg);

    // In console mode we are already running in the interactive session,
    // so spawning via WTSQueryUserToken/CreateProcessAsUserW is unnecessary
    // and often fails due to missing privileges. Just do the DDC/CI call directly.
    if (g_consoleMode) {
        bool ok = SetVcpOnDisplayDevice_RuntimeDxva2(kTargetDisplayDevice, vcpValue);

        // Print to console to make debugging obvious.
        std::wprintf(L"[Console] %s -> Set VCP 0x%02X = %lu : %s\n",
                     reason, static_cast<unsigned>(kVcpCode), vcpValue, ok ? L"OK" : L"FAILED");

        if (!ok) {
            LogEvent(EVENTLOG_ERROR_TYPE, L"Console mode: DDC/CI set failed.");
        }
        return;
    }

    // Service mode path (Session 0): spawn setter into active user session.
    if (!RunSelfSetterInActiveSession(vcpValue)) {
        LogEvent(EVENTLOG_ERROR_TYPE, L"Failed to start self DDC/CI setter process.");
    }
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == WM_DEVICECHANGE) {
        // We care about:
        // - DBT_DEVICEARRIVAL: device plugged in
        // - DBT_DEVICEREMOVECOMPLETE: device unplugged
        if (wParam == DBT_DEVICEARRIVAL
#if defined(USB_REMOVE_VCP_VALUE)
            || wParam == DBT_DEVICEREMOVECOMPLETE
#endif
        ) {
            auto *hdr = reinterpret_cast<DEV_BROADCAST_HDR *>(lParam);
            if (hdr && hdr->dbch_devicetype == DBT_DEVTYP_DEVICEINTERFACE) {
                auto *di = reinterpret_cast<DEV_BROADCAST_DEVICEINTERFACE_W *>(lParam);
                const wchar_t *name = di->dbcc_name; // variable-length array start

                if (name && name[0] != L'\0') {
                    auto *st = reinterpret_cast<WindowState *>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
                    if (st && ContainsCaseInsensitive(name, st->targetUpper)) {
                        if (wParam == DBT_DEVICEARRIVAL) {
                            HandleMatchedTrigger(L"MATCH HID ARRIVAL", kVcpValueOnArrival);
                        }
#if defined(USB_REMOVE_VCP_VALUE)
                        else if (wParam == DBT_DEVICEREMOVECOMPLETE) {
                            // Unplug action: switch to configured VCP value (default 17).
                            HandleMatchedTrigger(L"MATCH HID REMOVAL", (DWORD) USB_REMOVE_VCP_VALUE);
                        }
#endif
                    }
                }
            }
        }
        return 0;
    }

    // Ignore close attempts (this is a hidden service window).
    if (msg == WM_CLOSE) return 0;

    if (msg == WM_DESTROY) {
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static DWORD WINAPI WorkerThreadProc(LPVOID) {
    const wchar_t *kClassName = L"UsbHidSwitchSvcHiddenWnd";

    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = GetModuleHandleW(nullptr);
    wc.lpszClassName = kClassName;

    if (!RegisterClassExW(&wc)) {
        LogEvent(EVENTLOG_ERROR_TYPE, L"RegisterClassEx failed.");
        return 1;
    }

    WindowState st;
    st.targetUpper = ToUpper(kTargetVidPid);

    HWND hwnd = CreateWindowExW(
        0, kClassName, L"",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
        nullptr, nullptr, GetModuleHandleW(nullptr), nullptr);

    if (!hwnd) {
        LogEvent(EVENTLOG_ERROR_TYPE, L"CreateWindowEx failed.");
        UnregisterClassW(kClassName, GetModuleHandleW(nullptr));
        return 2;
    }

    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(&st));

    DEV_BROADCAST_DEVICEINTERFACE_W dbi{};
    dbi.dbcc_size = sizeof(dbi);
    dbi.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE;
    dbi.dbcc_classguid = GUID_DEVINTERFACE_HID_;

    st.notify = RegisterDeviceNotificationW(hwnd, &dbi, DEVICE_NOTIFY_WINDOW_HANDLE);
    if (!st.notify) {
        LogEvent(EVENTLOG_ERROR_TYPE, L"RegisterDeviceNotification failed.");
        DestroyWindow(hwnd);
        UnregisterClassW(kClassName, GetModuleHandleW(nullptr));
        return 3;
    }

    LogEvent(EVENTLOG_INFORMATION_TYPE, L"Worker started: listening HID arrivals/removals.");

#if !defined(USB_REMOVE_VCP_VALUE)
    LogEvent(EVENTLOG_INFORMATION_TYPE,
             L"Removal switching is disabled at compile time (USB_REMOVE_VCP_VALUE not defined).");
#endif

    for (;;) {
        DWORD r = MsgWaitForMultipleObjects(1, &g_stopEvent, FALSE, INFINITE, QS_ALLINPUT);

        if (r == WAIT_OBJECT_0)
            break; // stop requested

        if (r == WAIT_OBJECT_0 + 1) {
            MSG m;
            while (PeekMessageW(&m, nullptr, 0, 0, PM_REMOVE)) {
                if (m.message == WM_QUIT)
                    goto exit_loop;

                TranslateMessage(&m);
                DispatchMessageW(&m);
            }
        }
    }

exit_loop:
    if (st.notify) UnregisterDeviceNotification(st.notify);
    DestroyWindow(hwnd);
    UnregisterClassW(kClassName, GetModuleHandleW(nullptr));
    LogEvent(EVENTLOG_INFORMATION_TYPE, L"Worker stopped.");
    return 0;
}

// ============================================================================
//                              SERVICE BOILERPLATE
// ============================================================================

static void SetSvcStatus(DWORD state, DWORD win32ExitCode = NO_ERROR, DWORD waitHintMs = 0) {
    SERVICE_STATUS s{};
    s.dwServiceType = SERVICE_WIN32_OWN_PROCESS;
    s.dwCurrentState = state;
    s.dwControlsAccepted = (state == SERVICE_START_PENDING) ? 0 : (SERVICE_ACCEPT_STOP | SERVICE_ACCEPT_SHUTDOWN);
    s.dwWin32ExitCode = win32ExitCode;
    s.dwWaitHint = waitHintMs;

    static DWORD checkpoint = 1;
    s.dwCheckPoint = (state == SERVICE_RUNNING || state == SERVICE_STOPPED) ? 0 : checkpoint++;

    SetServiceStatus(g_svcStatusHandle, &s);
}

static DWORD WINAPI HandlerEx(DWORD control, DWORD, LPVOID, LPVOID) {
    switch (control) {
        case SERVICE_CONTROL_STOP:
        case SERVICE_CONTROL_SHUTDOWN:
            SetSvcStatus(SERVICE_STOP_PENDING, NO_ERROR, 5000);
            if (g_stopEvent) SetEvent(g_stopEvent);
            return NO_ERROR;
        default:
            return NO_ERROR;
    }
}

static void WINAPI ServiceMain(DWORD, LPWSTR *) {
    g_svcStatusHandle = RegisterServiceCtrlHandlerExW(kServiceName, HandlerEx, nullptr);
    if (!g_svcStatusHandle) return;

    SetSvcStatus(SERVICE_START_PENDING, NO_ERROR, 5000);

    g_stopEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    if (!g_stopEvent) {
        SetSvcStatus(SERVICE_STOPPED, GetLastError());
        return;
    }

    g_workerThread = CreateThread(nullptr, 0, WorkerThreadProc, nullptr, 0, nullptr);
    if (!g_workerThread) {
        SetSvcStatus(SERVICE_STOPPED, GetLastError());
        CloseHandle(g_stopEvent);
        g_stopEvent = nullptr;
        return;
    }

    SetSvcStatus(SERVICE_RUNNING);

    WaitForSingleObject(g_workerThread, INFINITE);

    CloseHandle(g_workerThread);
    g_workerThread = nullptr;

    CloseHandle(g_stopEvent);
    g_stopEvent = nullptr;

    SetSvcStatus(SERVICE_STOPPED);
}

// ============================================================================
//                                  ENTRYPOINT
// ============================================================================
//
// Modes:
// 1) Setter mode: exe --set-vcp <value>
//    - Called by the service in active user session.
//    - Performs DDC/CI set and exits.
//
// 2) Service mode: started by SCM
//
// 3) Console debug mode: started manually (Ctrl+C to exit)
//
static bool TryParseDword(const wchar_t *s, DWORD *out) {
    if (!s || !s[0] || !out) return false;

    // wcstoul supports "0x.." if base==0.
    wchar_t *end = nullptr;
    unsigned long v = wcstoul(s, &end, 0);
    if (end == s || (end && *end != L'\0'))
        return false;

    *out = (DWORD) v;
    return true;
}

int wmain(int argc, wchar_t **argv) {
    // ---------------- Mode 1: setter mode ----------------
    if (argc >= 2 && _wcsicmp(argv[1], kSelfSetterFlag) == 0) {
        // Expect: --set-vcp <value>
        DWORD v = 0;
        if (argc < 3 || !TryParseDword(argv[2], &v)) {
            LogEvent(EVENTLOG_ERROR_TYPE, L"Setter mode usage: --set-vcp <value>. Missing/invalid value.");
            return 3;
        }

        // Do the DDC/CI action and exit.
        if (SetVcpOnDisplayDevice_RuntimeDxva2(kTargetDisplayDevice, v)) {
            LogEvent(EVENTLOG_INFORMATION_TYPE, L"DDC/CI: VCP set succeeded.");
            return 0;
        } else {
            LogEvent(EVENTLOG_ERROR_TYPE, L"DDC/CI: VCP set failed.");
            return 2;
        }
    }

    // ---------------- Mode 2: service mode ----------------
    SERVICE_TABLE_ENTRYW table[] = {
        {const_cast<LPWSTR>(kServiceName), ServiceMain},
        {nullptr, nullptr}
    };

    if (StartServiceCtrlDispatcherW(table))
        return 0;

    // ---------------- Mode 3: console debug mode ----------------
    DWORD err = GetLastError();
    if (err != ERROR_FAILED_SERVICE_CONTROLLER_CONNECT) {
        LogEvent(EVENTLOG_ERROR_TYPE, L"StartServiceCtrlDispatcher failed (unexpected).");
    }

    std::wprintf(L"[Debug] Running in console mode (not started by SCM).\n");
    g_consoleMode = true;

    g_stopEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    if (!g_stopEvent) return 1;

    SetConsoleCtrlHandler([](DWORD type)-> BOOL {
        if (type == CTRL_C_EVENT || type == CTRL_CLOSE_EVENT || type == CTRL_BREAK_EVENT) {
            if (g_stopEvent) SetEvent(g_stopEvent);
            return TRUE;
        }
        return FALSE;
    }, TRUE);

    WorkerThreadProc(nullptr);

    CloseHandle(g_stopEvent);
    g_stopEvent = nullptr;
    return 0;
}
