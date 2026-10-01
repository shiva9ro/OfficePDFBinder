#pragma once
#include <windows.h>
#include <shobjidl.h>
#include <string>
#include <vector>

namespace binder {
// Shared delivery for the packaged IExplorerCommand adapter and COM server.
HRESULT CollectPaths(IShellItemArray* selection, std::vector<std::wstring>& paths);
HRESULT DeliverPaths(const std::vector<std::wstring>& paths,
                     const std::wstring& application,
                     const std::wstring& pipeName,
                     DWORD startupTimeout = 30000);
std::wstring ApplicationPath();
std::wstring SessionPipeName();
}
