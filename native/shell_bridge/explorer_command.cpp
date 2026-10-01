#include <windows.h>
#include <shobjidl.h>
#include <shlwapi.h>
#include <new>
#include <string>

#pragma comment(linker, "/EXPORT:DllGetClassObject")
#pragma comment(linker, "/EXPORT:DllCanUnloadNow")

namespace {
HMODULE module;
LONG references = 0;
const CLSID CommandId = {0xf9869918,0x6f6d,0x4fc9,{0xa8,0xdf,0x70,0xf4,0x34,0x98,0xa8,0x04}};
const CLSID DeliveryId = {0xf9869918,0x6f6d,0x4fc9,{0xa8,0xdf,0x70,0xf4,0x34,0x98,0xa8,0x02}};

class Command final : public IExplorerCommand {
    LONG refs_ = 1;
public:
    Command() { InterlockedIncrement(&references); }
    ~Command() { InterlockedDecrement(&references); }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID id, void** out) override {
        if (!out) return E_POINTER;
        *out = nullptr;
        if (id != IID_IUnknown && id != IID_IExplorerCommand) return E_NOINTERFACE;
        *out = static_cast<IExplorerCommand*>(this); AddRef(); return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return InterlockedIncrement(&refs_); }
    ULONG STDMETHODCALLTYPE Release() override {
        ULONG count = InterlockedDecrement(&refs_); if (!count) delete this; return count;
    }
    HRESULT STDMETHODCALLTYPE GetTitle(IShellItemArray*, PWSTR* out) override {
        return out ? SHStrDupW(PRIMARYLANGID(GetUserDefaultUILanguage()) == LANG_JAPANESE
            ? L"Office PDF Binder で開く" : L"Open with Office PDF Binder", out) : E_POINTER;
    }
    HRESULT STDMETHODCALLTYPE GetIcon(IShellItemArray*, PWSTR* out) override {
        if (!out) return E_POINTER;
        *out = nullptr;
        try {
            wchar_t path[32768];
            DWORD size = GetModuleFileNameW(module, path, ARRAYSIZE(path));
            if (!size || size >= ARRAYSIZE(path)) return E_FAIL;
            std::wstring icon(path, size);
            icon.resize(icon.find_last_of(L"\\/") + 1);
            icon += L"app.ico";
            return SHStrDupW(icon.c_str(), out);
        } catch (...) { return E_OUTOFMEMORY; }
    }
    HRESULT STDMETHODCALLTYPE GetToolTip(IShellItemArray*, PWSTR* out) override {
        if (!out) return E_POINTER;
        *out = nullptr; return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE GetCanonicalName(GUID* out) override {
        if (!out) return E_POINTER;
        *out = CommandId; return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetState(IShellItemArray* items, BOOL, EXPCMDSTATE* out) override {
        if (!out) return E_POINTER;
        *out = ECS_HIDDEN;
        DWORD count = 0;
        if (!items || FAILED(items->GetCount(&count)) || !count) return S_OK;
        for (DWORD index = 0; index < count; ++index) {
            IShellItem* item = nullptr;
            if (FAILED(items->GetItemAt(index, &item))) return S_OK;
            SFGAOF attributes = 0;
            HRESULT attrResult = item->GetAttributes(SFGAO_FOLDER, &attributes);
            if (FAILED(attrResult) || (attributes & SFGAO_FOLDER)) { item->Release(); return S_OK; }
            PWSTR path = nullptr;
            HRESULT hr = item->GetDisplayName(SIGDN_FILESYSPATH, &path);
            item->Release();
            if (FAILED(hr)) return S_OK;
            const wchar_t* extension = PathFindExtensionW(path);
            bool supported = false;
            for (auto candidate : {L".pdf", L".doc", L".docx", L".docm", L".xls", L".xlsx", L".xlsm",
                                   L".ppt", L".pptx", L".pptm", L".png", L".jpg", L".jpeg", L".bmp",
                                   L".webp", L".tif", L".tiff", L".heic", L".heif", L".hif", L".svg"}) {
                if (_wcsicmp(extension, candidate) == 0) { supported = true; break; }
            }
            CoTaskMemFree(path);
            if (!supported) return S_OK;
        }
        *out = ECS_ENABLED;
        return S_OK; // No disk access, IPC or process launch while constructing the menu.
    }
    HRESULT STDMETHODCALLTYPE Invoke(IShellItemArray* items, IBindCtx*) override {
        if (!items) return E_INVALIDARG;
        // Forward the complete selection to the existing out-of-process adapter.
        // Both menus therefore use the same framing, ACK, timeout and launch logic.
        IExecuteCommand* command = nullptr;
        HRESULT hr = CoCreateInstance(DeliveryId, nullptr, CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&command));
        if (FAILED(hr)) return hr;
        CoAllowSetForegroundWindow(command, nullptr);
        IObjectWithSelection* target = nullptr;
        hr = command->QueryInterface(IID_PPV_ARGS(&target));
        if (SUCCEEDED(hr)) {
            hr = target->SetSelection(items);
            if (SUCCEEDED(hr)) hr = command->Execute();
            target->Release();
        }
        command->Release(); return hr;
    }
    HRESULT STDMETHODCALLTYPE GetFlags(EXPCMDFLAGS* out) override {
        if (!out) return E_POINTER;
        *out = ECF_DEFAULT; return S_OK;
    }
    HRESULT STDMETHODCALLTYPE EnumSubCommands(IEnumExplorerCommand** out) override {
        if (!out) return E_POINTER;
        *out = nullptr; return E_NOTIMPL;
    }
};

class Factory final : public IClassFactory {
    LONG refs_ = 1;
public:
    Factory() { InterlockedIncrement(&references); }
    ~Factory() { InterlockedDecrement(&references); }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID id, void** out) override {
        if (!out) return E_POINTER;
        *out = nullptr;
        if (id != IID_IUnknown && id != IID_IClassFactory) return E_NOINTERFACE;
        *out = static_cast<IClassFactory*>(this); AddRef(); return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return InterlockedIncrement(&refs_); }
    ULONG STDMETHODCALLTYPE Release() override {
        ULONG count = InterlockedDecrement(&refs_); if (!count) delete this; return count;
    }
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* outer, REFIID id, void** out) override {
        if (!out) return E_POINTER;
        *out = nullptr;
        if (outer) return CLASS_E_NOAGGREGATION;
        auto* command = new (std::nothrow) Command;
        if (!command) return E_OUTOFMEMORY;
        HRESULT hr = command->QueryInterface(id, out); command->Release(); return hr;
    }
    HRESULT STDMETHODCALLTYPE LockServer(BOOL lock) override {
        if (lock) InterlockedIncrement(&references); else InterlockedDecrement(&references);
        return S_OK;
    }
};
}

STDAPI DllGetClassObject(REFCLSID id, REFIID iid, void** out) {
    if (!out) return E_POINTER;
    *out = nullptr;
    if (id != CommandId) return CLASS_E_CLASSNOTAVAILABLE;
    auto* factory = new (std::nothrow) Factory;
    if (!factory) return E_OUTOFMEMORY;
    HRESULT hr = factory->QueryInterface(iid, out); factory->Release(); return hr;
}
STDAPI DllCanUnloadNow() { return InterlockedCompareExchange(&references, 0, 0) ? S_FALSE : S_OK; }
BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, void*) {
    if (reason == DLL_PROCESS_ATTACH) module = instance;
    return TRUE;
}
