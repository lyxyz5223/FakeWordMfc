#pragma once
#include <atlbase.h>
#include <atlcom.h>
#include <msctf.h>
#include <string>

#include <msctf.h>
#include <windows.h>


class TSFInterceptor : public ITfTextEditSink {
public:
    STDMETHODIMP OnEndEdit(
        ITfContext* pic,
        TfEditCookie ecReadOnly,
        ITfEditRecord* pEditRecord) override
    {
        if (!pEditRecord)
            return S_OK;
        IEnumTfRanges* pEnum = nullptr;
        if (SUCCEEDED(pEditRecord->GetTextAndPropertyUpdates(TF_GTP_INCL_TEXT, nullptr, 0, &pEnum)))
        {
            ITfRange* pRange = nullptr;
            ULONG cFetch = 0;
            while (pEnum->Next(1, &pRange, &cFetch) == S_OK)
            {
                WCHAR buf[512] = {};
                ULONG cch = 0;
                if (SUCCEEDED(pRange->GetText(ecReadOnly, 0, buf, 511, &cch)) && cch)
                {
                    OutputDebugStringW(buf);
                    return E_FAIL;
                }
                pRange->Release();
            }
            pEnum->Release();
        }
        return S_OK;

    }

    ULONG STDMETHODCALLTYPE AddRef() override {
        return 1;
    }
    ULONG STDMETHODCALLTYPE Release() override {
        return 1;
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override {
        if (riid == IID_IUnknown || riid == IID_ITfTextEditSink)
        {
            *ppv = this;
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
};


class TSFKeyEventInterceptor : public ITfKeyEventSink {
public:

    STDMETHODIMP OnTestKeyDown(ITfContext* pic, WPARAM wParam, LPARAM lParam, BOOL* pfEaten) override {
        *pfEaten = ShouldInterceptKey(wParam, lParam); // *pfEaten��ʾ�Ƿ��̵�����
        return S_OK;
    }

    STDMETHODIMP OnKeyDown(ITfContext* pic, WPARAM wParam, LPARAM lParam, BOOL* pfEaten) override {
        if (ShouldInterceptKey(wParam, lParam)) {
            *pfEaten = TRUE; // ���ذ���
            // ִ�����ز���
            HandleInterceptedKey(wParam, lParam, TRUE);
            return S_OK;
        }
        *pfEaten = FALSE;
        return S_OK;
    }

    STDMETHODIMP OnSetFocus(BOOL fForeground) override {
        return S_OK;
    }

    STDMETHODIMP OnTestKeyUp(ITfContext* pic, WPARAM wParam, LPARAM lParam, BOOL* pfEaten) override {
        return S_OK;
    }

    STDMETHODIMP OnKeyUp(ITfContext* pic, WPARAM wParam, LPARAM lParam, BOOL* pfEaten) override {
        return S_OK;
    }

    STDMETHODIMP OnPreservedKey(ITfContext* pic, REFGUID rguid, BOOL* pfEaten) override {
        return S_OK;
    }

    STDMETHODIMP_(ULONG) AddRef() override {
        return InterlockedIncrement(&refCount);
    }

    STDMETHODIMP_(ULONG) Release() override {
        ULONG refCount = InterlockedDecrement(&this->refCount);
        if (refCount == 0) {
            delete this;
        }
        return refCount;
    }

    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (riid == IID_IUnknown || riid == IID_ITfKeyEventSink)
        {
            *ppv = this;
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }

private:
    BOOL ShouldInterceptKey(WPARAM wParam, LPARAM lParam) {
        return TRUE;
        // ��������Ctrl+��ĸ
        if (GetAsyncKeyState(VK_CONTROL) & 0x8000) {
            if (wParam >= 'A' && wParam <= 'Z') return TRUE;
        }

        // ���ع��ܼ�
        if (wParam >= VK_F1 && wParam <= VK_F12) return TRUE;

        // �����ض�����
        switch (wParam) {
        case VK_ESCAPE:
        case VK_TAB:
        case VK_SNAPSHOT: // Print Screen
            return TRUE;
        }

        return FALSE;
    }
    void HandleInterceptedKey(WPARAM wParam, LPARAM lParam, BOOL isKeyDown) {
        OutputDebugString(L"���������ص�key: ");
        OutputDebugString(std::to_wstring(wParam).c_str());
        OutputDebugString(L"\n");
    }
    ULONG refCount;
};

