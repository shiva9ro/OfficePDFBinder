#include "delivery.h"
#include <algorithm>
#include <cstdint>

namespace {
struct Handle {
    HANDLE value;
    ~Handle() { if (value && value != INVALID_HANDLE_VALUE) CloseHandle(value); }
};
HRESULT LastErrorResult() { return HRESULT_FROM_WIN32(GetLastError()); }

// Use overlapped I/O so an unresponsive application cannot hold the COM server
// indefinitely. No automatic resend after writing: the outcome may be unknown.
HRESULT Transfer(HANDLE pipe, void* data, DWORD size, bool writing) {
    auto* bytes = static_cast<char*>(data);
    const ULONGLONG deadline = GetTickCount64() + 10000;
    while (size) {
        Handle event{CreateEventW(nullptr, TRUE, FALSE, nullptr)};
        if (!event.value) return LastErrorResult();
        OVERLAPPED ov{};
        ov.hEvent = event.value;
        DWORD done = 0;
        BOOL ok = writing ? WriteFile(pipe, bytes, size, &done, &ov)
                          : ReadFile(pipe, bytes, size, &done, &ov);
        if (!ok) {
            DWORD error = GetLastError();
            if (error != ERROR_IO_PENDING) return HRESULT_FROM_WIN32(error);
            ULONGLONG now = GetTickCount64();
            DWORD wait = now < deadline ? static_cast<DWORD>(deadline - now) : 0;
            if (WaitForSingleObject(event.value, wait) != WAIT_OBJECT_0) {
                CancelIoEx(pipe, &ov);
                GetOverlappedResult(pipe, &ov, &done, TRUE);
                return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
            }
            if (!GetOverlappedResult(pipe, &ov, &done, FALSE)) return LastErrorResult();
        }
        if (!done) return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
        bytes += done;
        size -= done;
    }
    return S_OK;
}
}

namespace binder {
std::wstring SessionPipeName() {
    DWORD session = 0;
    if (!ProcessIdToSessionId(GetCurrentProcessId(), &session)) return {};
    return L"OfficePDFBinder_Installed_SingleInstance_Session" + std::to_wstring(session);
}

HRESULT CollectPaths(IShellItemArray* selection, std::vector<std::wstring>& paths) {
    paths.clear();
    if (!selection) return E_INVALIDARG;
    DWORD count = 0;
    HRESULT hr = selection->GetCount(&count);
    if (FAILED(hr)) return hr;
    for (DWORD i = 0; i < count; ++i) {
        IShellItem* item = nullptr;
        hr = selection->GetItemAt(i, &item);
        if (FAILED(hr)) return hr;
        PWSTR path = nullptr;
        hr = item->GetDisplayName(SIGDN_FILESYSPATH, &path);
        item->Release();
        if (FAILED(hr)) return hr;  // Do not silently lose virtual items.
        paths.emplace_back(path);
        CoTaskMemFree(path);
    }
    return paths.empty() ? E_INVALIDARG : S_OK;
}

std::wstring ApplicationPath() {
    wchar_t path[32768];
    DWORD length = GetModuleFileNameW(nullptr, path, ARRAYSIZE(path));
    if (!length || length >= ARRAYSIZE(path)) return {};
    std::wstring result(path, length);
    return result.substr(0, result.find_last_of(L"\\/") + 1) + L"OfficePDFBinder_Main.exe";
}

HRESULT DeliverPaths(const std::vector<std::wstring>& paths,
                     const std::wstring& application, const std::wstring& pipeName,
                     DWORD startupTimeout) {
    if (paths.empty()) return E_INVALIDARG;
    std::wstring joined;
    for (const auto& path : paths) {
        if (path.empty() || path.find_first_of(L"\r\n") != std::wstring::npos) return E_INVALIDARG;
        if (!joined.empty()) joined += L'\n';
        joined += path;
    }
    if (joined.size() > 16 * 1024 * 1024) return E_INVALIDARG;
    int length = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, joined.data(),
                                    static_cast<int>(joined.size()), nullptr, 0, nullptr, nullptr);
    if (!length) return LastErrorResult();
    if (length > 16 * 1024 * 1024) return E_INVALIDARG;
    // OPB2 + little-endian uint32 byte count + UTF-8 newline-separated paths.
    std::string packet("OPB2", 4);
    for (int shift = 0; shift < 32; shift += 8) packet += static_cast<char>((length >> shift) & 255);
    packet.resize(8 + length);
    WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, joined.data(),
                        static_cast<int>(joined.size()), packet.data() + 8, length, nullptr, nullptr);

    std::wstring pipePath = L"\\\\.\\pipe\\" + pipeName;
    bool launched = false;
    const ULONGLONG deadline = GetTickCount64() + startupTimeout;
    HANDLE connected = INVALID_HANDLE_VALUE;
    do {
        connected = CreateFileW(pipePath.c_str(), GENERIC_READ | GENERIC_WRITE, 0,
                                nullptr, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, nullptr);
        if (connected != INVALID_HANDLE_VALUE) break;
        DWORD error = GetLastError();
        if (error != ERROR_FILE_NOT_FOUND && error != ERROR_PIPE_BUSY) return HRESULT_FROM_WIN32(error);
        if (!launched && error == ERROR_FILE_NOT_FOUND && !application.empty()) {
            std::wstring command = L"\"" + application + L"\"";
            STARTUPINFOW startup{sizeof(startup)};
            PROCESS_INFORMATION process{};
            if (!CreateProcessW(application.c_str(), command.data(), nullptr, nullptr, FALSE,
                                CREATE_NO_WINDOW, nullptr, nullptr, &startup, &process)) return LastErrorResult();
            CloseHandle(process.hThread);
            CloseHandle(process.hProcess);
            launched = true;
        }
        Sleep(50);
    } while (GetTickCount64() < deadline);
    if (connected == INVALID_HANDLE_VALUE) return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
    Handle pipe{connected};
    // Permission is handed from the shell through COM to the receiving GUI.
    ULONG serverPid = 0;
    if (GetNamedPipeServerProcessId(pipe.value, &serverPid)) AllowSetForegroundWindow(serverPid);
    HRESULT hr = Transfer(pipe.value, packet.data(), static_cast<DWORD>(packet.size()), true);
    if (FAILED(hr)) return hr;
    char ack[3]{};
    hr = Transfer(pipe.value, ack, sizeof(ack), false);
    if (FAILED(hr)) return hr;
    return std::string(ack, 3) == "OK\n" ? S_OK : E_FAIL;
}
}
