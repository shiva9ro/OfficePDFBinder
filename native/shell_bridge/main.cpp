#include "delivery.h"
#include <new>

// Keep in sync with packaging/setup_office_binder.iss.
static const CLSID CLSID_BinderVerb =
#ifdef BINDER_INTEGRATION_TEST
    {0xf9869918,0x6f6d,0x4fc9,{0xa8,0xdf,0x70,0xf4,0x34,0x98,0xa8,0x03}};
static constexpr auto DeliveryPipe = L"OfficePDFBinder_ShellBridge_IntegrationTest";
#else
    {0xf9869918,0x6f6d,0x4fc9,{0xa8,0xdf,0x70,0xf4,0x34,0x98,0xa8,0x02}};
#endif
static LONG objects = 0;
static LONG locks = 0;

class Command final : public IExecuteCommand, public IObjectWithSelection {
    LONG refs_ = 1;
    IShellItemArray* selection_ = nullptr;
public:
    Command() { ++objects; }
    ~Command() { if (selection_) selection_->Release(); --objects; }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** out) override {
        if (!out) return E_POINTER;
        *out = nullptr;
        if (iid == IID_IUnknown || iid == IID_IExecuteCommand) *out = static_cast<IExecuteCommand*>(this);
        else if (iid == IID_IObjectWithSelection) *out = static_cast<IObjectWithSelection*>(this);
        else return E_NOINTERFACE;
        AddRef(); return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return InterlockedIncrement(&refs_); }
    ULONG STDMETHODCALLTYPE Release() override {
        ULONG n = InterlockedDecrement(&refs_); if (!n) delete this; return n;
    }
    HRESULT STDMETHODCALLTYPE SetSelection(IShellItemArray* value) override {
        if (value) value->AddRef();
        if (selection_) selection_->Release();
        selection_ = value; return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetSelection(REFIID iid, void** out) override {
        if (!out) return E_POINTER;
        *out = nullptr;
        return selection_ ? selection_->QueryInterface(iid, out) : E_FAIL;
    }
    HRESULT STDMETHODCALLTYPE SetKeyState(DWORD) override { return S_OK; }
    HRESULT STDMETHODCALLTYPE SetParameters(LPCWSTR) override { return S_OK; }
    HRESULT STDMETHODCALLTYPE SetPosition(POINT) override { return S_OK; }
    HRESULT STDMETHODCALLTYPE SetShowWindow(int) override { return S_OK; }
    HRESULT STDMETHODCALLTYPE SetNoShowUI(BOOL) override { return S_OK; }
    HRESULT STDMETHODCALLTYPE SetDirectory(LPCWSTR) override { return S_OK; }
    HRESULT STDMETHODCALLTYPE Execute() override {
        try {
            std::vector<std::wstring> paths;
            HRESULT hr = binder::CollectPaths(selection_, paths);
#ifndef BINDER_INTEGRATION_TEST
            const auto DeliveryPipe = binder::SessionPipeName();
            if (DeliveryPipe.empty()) return E_FAIL;
#endif
            if (SUCCEEDED(hr)) hr = binder::DeliverPaths(paths, binder::ApplicationPath(), DeliveryPipe);
            // Propagate failure to Explorer; never claim that a lost request succeeded.
            return hr;
        } catch (const std::bad_alloc&) { return E_OUTOFMEMORY; }
        catch (...) { return E_FAIL; }
    }
};

class Factory final : public IClassFactory {
    LONG refs_ = 1;
public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** out) override {
        if (!out) return E_POINTER;
        *out = nullptr;
        if (iid != IID_IUnknown && iid != IID_IClassFactory) return E_NOINTERFACE;
        *out = static_cast<IClassFactory*>(this); AddRef(); return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return InterlockedIncrement(&refs_); }
    ULONG STDMETHODCALLTYPE Release() override {
        ULONG n = InterlockedDecrement(&refs_); if (!n) delete this; return n;
    }
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* outer, REFIID iid, void** out) override {
        if (!out) return E_POINTER;
        *out = nullptr;
        if (outer) return CLASS_E_NOAGGREGATION;
        auto* command = new (std::nothrow) Command;
        if (!command) return E_OUTOFMEMORY;
        HRESULT hr = command->QueryInterface(iid, out);
        command->Release(); return hr;
    }
    HRESULT STDMETHODCALLTYPE LockServer(BOOL lock) override {
        if (lock) ++locks; else --locks;
        return S_OK;
    }
};

int WINAPI wWinMain(HINSTANCE, HINSTANCE, PWSTR arguments, int) {
    CLSID registeredClass = CLSID_BinderVerb;
#ifdef BINDER_INTEGRATION_TEST
    if (arguments && arguments[0] == L'{') {
        std::wstring id(arguments, 38);
        if (FAILED(CLSIDFromString(id.c_str(), &registeredClass))) return 1;
    }
#else
    (void)arguments;
#endif
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) return 1;
    auto* factory = new (std::nothrow) Factory;
    if (!factory) { CoUninitialize(); return 1; }
    DWORD cookie = 0;
    hr = CoRegisterClassObject(registeredClass, factory, CLSCTX_LOCAL_SERVER,
                               REGCLS_MULTIPLEUSE | REGCLS_SUSPENDED, &cookie);
    factory->Release();
    if (FAILED(hr)) { CoUninitialize(); return 1; }
    CoResumeClassObjects();
    // A short-lived local COM server, never loaded into Explorer's process.
    UINT_PTR timer = SetTimer(nullptr, 0, 1000, nullptr);
    if (!timer) { CoRevokeClassObject(cookie); CoUninitialize(); return 1; }
    unsigned idle = 0;
    MSG message;
    while (GetMessageW(&message, nullptr, 0, 0) > 0) {
        if (message.message == WM_TIMER && message.wParam == timer) {
            idle = objects || locks ? 0 : idle + 1;
            if (idle >= 30) {
                CoSuspendClassObjects();
                if (!objects && !locks) break;
                CoResumeClassObjects();
                idle = 0;
            }
        }
        TranslateMessage(&message); DispatchMessageW(&message);
    }
    KillTimer(nullptr, timer);
    CoRevokeClassObject(cookie);
    CoUninitialize();
    return 0;
}
