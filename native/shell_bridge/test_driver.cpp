// Integration-test entry point; not shipped in the installer.
#include "delivery.h"
#include <shlobj.h>
#include <iostream>

int wmain(int argc, wchar_t** argv) {
    if (argc == 2 && std::wstring(argv[1]) == L"--pipe-name") {
        std::wcout << binder::SessionPipeName() << std::endl;
        return 0;
    }
    if (argc < 4) return 2;
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) return 3;
    std::vector<PIDLIST_ABSOLUTE> ids;
    for (int i = 3; i < argc; ++i) {
        PIDLIST_ABSOLUTE id = nullptr;
        hr = SHParseDisplayName(argv[i], nullptr, &id, 0, nullptr);
        if (FAILED(hr)) break;
        ids.push_back(id);
    }
    IShellItemArray* selection = nullptr;
    if (SUCCEEDED(hr)) hr = SHCreateShellItemArrayFromIDLists(
        static_cast<UINT>(ids.size()), const_cast<PCIDLIST_ABSOLUTE*>(ids.data()), &selection);
    std::vector<std::wstring> paths;
    if (SUCCEEDED(hr)) hr = binder::CollectPaths(selection, paths);
    if (selection) selection->Release();
    for (auto id : ids) CoTaskMemFree(id);
    if (SUCCEEDED(hr) && std::wstring(argv[1]) == L"--com") {
        CLSID testId;
        hr = CLSIDFromString(argv[2], &testId);
        IExecuteCommand* command = nullptr;
        if (SUCCEEDED(hr)) hr = CoCreateInstance(testId, nullptr, CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&command));
        std::cout << "activation=" << std::hex << static_cast<unsigned long>(hr) << std::endl;
        IObjectWithSelection* target = nullptr;
        if (SUCCEEDED(hr)) hr = command->QueryInterface(IID_PPV_ARGS(&target));
        // Recreate the Shell selection for the cross-process COM call.
        std::vector<PIDLIST_ABSOLUTE> selectedIds;
        for (const auto& path : paths) {
            PIDLIST_ABSOLUTE id = nullptr;
            if (SUCCEEDED(hr)) hr = SHParseDisplayName(path.c_str(), nullptr, &id, 0, nullptr);
            if (id) selectedIds.push_back(id);
        }
        IShellItemArray* selected = nullptr;
        if (SUCCEEDED(hr)) hr = SHCreateShellItemArrayFromIDLists(static_cast<UINT>(selectedIds.size()),
            const_cast<PCIDLIST_ABSOLUTE*>(selectedIds.data()), &selected);
        if (SUCCEEDED(hr)) hr = target->SetSelection(selected);
        if (SUCCEEDED(hr)) hr = command->Execute();
        if (selected) selected->Release();
        for (auto id : selectedIds) CoTaskMemFree(id);
        if (target) target->Release();
        if (command) command->Release();
    } else if (SUCCEEDED(hr)) hr = binder::DeliverPaths(paths, argv[2], argv[1], 2000);
    std::cout << std::hex << static_cast<unsigned long>(hr) << std::endl;
    CoUninitialize();
    return FAILED(hr) ? 1 : 0;
}
