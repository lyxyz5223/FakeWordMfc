#pragma once
#include "pch.h"

// uia_interception.cpp
#include <windows.h>
#include <uiautomation.h>
#include <iostream>
#include <string>
#include <vector>
#include "StringProcess.h"

class EventHandler : public IUIAutomationEventHandler {
private:
    LONG m_refCount;
    IUIAutomation* m_pAutomation;

public:
    EventHandler(IUIAutomation* pAutomation) : m_refCount(1), m_pAutomation(pAutomation) {
        m_pAutomation->AddRef();
    }

    ~EventHandler() {
        if (m_pAutomation) {
            m_pAutomation->Release();
        }
    }

    // IUnknown methods
    ULONG STDMETHODCALLTYPE AddRef() override {
        return InterlockedIncrement(&m_refCount);
    }

    ULONG STDMETHODCALLTYPE Release() override {
        ULONG refCount = InterlockedDecrement(&m_refCount);
        if (refCount == 0) {
            delete this;
        }
        return refCount;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppInterface) override {
        if (riid == __uuidof(IUnknown)) {
            *ppInterface = static_cast<IUIAutomationEventHandler*>(this);
        }
        else if (riid == __uuidof(IUIAutomationEventHandler)) {
            *ppInterface = static_cast<IUIAutomationEventHandler*>(this);
        }
        else {
            *ppInterface = NULL;
            return E_NOINTERFACE;
        }
        AddRef();
        return S_OK;
    }

    // IUIAutomationEventHandler method
    HRESULT STDMETHODCALLTYPE HandleAutomationEvent(IUIAutomationElement* sender, EVENTID eventId) override {
        if (eventId == UIA_Text_TextChangedEventId) {
            // 文本改变事件
            HRESULT hr;
            VARIANT value;
            VariantInit(&value);

            hr = sender->GetCurrentPropertyValue(UIA_ValueValuePropertyId, &value);
            if (SUCCEEDED(hr) && value.vt == VT_BSTR) {
                std::wstring wstr = value.bstrVal;
                std::cout << "文本改变: " << wstr2str_2ANSI(wstr) << std::endl;
                VariantClear(&value);
            }
        }
        else if (eventId == UIA_AutomationFocusChangedEventId) {
            // 焦点改变事件
            std::cout << "焦点改变\n";

            // 获取控件信息
            CONTROLTYPEID controlType;
            if (SUCCEEDED(sender->get_CurrentControlType(&controlType))) {
                if (controlType == UIA_EditControlTypeId ||
                    controlType == UIA_DocumentControlTypeId ||
                    controlType == UIA_TextControlTypeId) {
                    printf("输入框获得焦点\n");
                }
            }
        }

        return S_OK;
    }
};

class InputInterceptor {
private:
    IUIAutomation* m_pAutomation;
    EventHandler* m_pEventHandler;
    std::vector<IUIAutomationElement*> m_elements;

public:
    InputInterceptor() : m_pAutomation(nullptr), m_pEventHandler(nullptr) {}

    bool Initialize() {
        HRESULT hr = CoInitialize(NULL);
        if (FAILED(hr)) {
            printf("CoInitialize失败: 0x%08X\n", hr);

        }

        hr = CoCreateInstance(__uuidof(CUIAutomation), NULL,
            CLSCTX_INPROC_SERVER, __uuidof(IUIAutomation),
            (void**)&m_pAutomation);
        if (FAILED(hr)) {
            printf("创建UIAutomation实例失败: 0x%08X\n", hr);
            CoUninitialize();
            return false;
        }

        m_pEventHandler = new EventHandler(m_pAutomation);
        return true;
    }

    // 注册全局焦点改变事件
    bool RegisterFocusEvent() {
        if (!m_pAutomation || !m_pEventHandler) return false;

        HRESULT hr = m_pAutomation->AddAutomationEventHandler(
            UIA_AutomationFocusChangedEventId,
            0,
            TreeScope_Element,
            NULL,  // cacheRequest
            m_pEventHandler);

        if (SUCCEEDED(hr)) {
            printf("焦点改变事件注册成功\n");
            return true;
        }
        else {
            printf("焦点改变事件注册失败: 0x%08X\n", hr);
            return false;
        }
    }

    // 为特定元素注册文本改变事件
    bool RegisterTextChangeEventForWindow(HWND hwnd) {
        if (!m_pAutomation || !m_pEventHandler) return false;

        IUIAutomationElement* element = nullptr;
        HRESULT hr = m_pAutomation->ElementFromHandle(hwnd, &element);
        if (FAILED(hr) || !element) {
            _com_error err(hr);
            std::cout << "获取UIA元素失败: 0x" << std::hex << hr << ", " << wstr2str_2ANSI(err.ErrorMessage()) << std::dec << std::endl;
            return false;
        }

        // 查找所有编辑控件
        IUIAutomationCondition* condition = nullptr;
        VARIANT varProp;
        varProp.vt = VT_I4;

        // 查找编辑控件
        varProp.lVal = UIA_EditControlTypeId;
        hr = m_pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, varProp, &condition);

        if (SUCCEEDED(hr)) {
            IUIAutomationElementArray* elementArray = nullptr;
            hr = element->FindAll(TreeScope_Descendants, condition, &elementArray);

            if (SUCCEEDED(hr) && elementArray) {
                int length = 0;
                elementArray->get_Length(&length);

                for (int i = 0; i < length; i++) {
                    IUIAutomationElement* editElement = nullptr;
                    elementArray->GetElement(i, &editElement);

                    if (editElement) {
                        hr = m_pAutomation->AddAutomationEventHandler(
                            UIA_Text_TextChangedEventId,
                            editElement,
                            TreeScope_Element,
                            NULL,
                            m_pEventHandler);

                        if (SUCCEEDED(hr)) {
                            printf("文本改变事件注册成功 for element %d\n", i);
                            m_elements.push_back(editElement);
                        }
                        else {
                            printf("文本改变事件注册失败: 0x%08X\n", hr);
                            editElement->Release();
                        }
                    }
                }
                elementArray->Release();
            }
            condition->Release();
        }

        element->Release();
        return true;
    }

    void StopInterception() {
        if (m_pAutomation && m_pEventHandler) {
            // 移除焦点改变事件
            m_pAutomation->RemoveAutomationEventHandler(UIA_AutomationFocusChangedEventId, 0, m_pEventHandler);

            // 移除所有文本改变事件
            for (auto element : m_elements) {
                m_pAutomation->RemoveAutomationEventHandler(
                    UIA_Text_TextChangedEventId, element, m_pEventHandler);
                element->Release();
            }
            m_elements.clear();
        }
    }

    ~InputInterceptor() {
        StopInterception();
        if (m_pEventHandler) m_pEventHandler->Release();
        if (m_pAutomation) m_pAutomation->Release();
        CoUninitialize();
    }
};

/*
* 使用方法：
    InputInterceptor* interceptor = new InputInterceptor();

    if (interceptor->Initialize()) {
        wprintf(L"输入拦截器初始化成功\n");

        // 注册全局焦点改变事件
        if (interceptor->RegisterTextChangeEventForWindow(wordDocumentHwnd)) {
            wprintf(L"开始监听输入事件...\n");
            wprintf(L"按任意键退出\n");
        }
        if (interceptor->RegisterTextChangeEventForWindow(wordAppHwnd)) {
            wprintf(L"开始监听输入事件...\n");
            wprintf(L"按任意键退出\n");
        }
    }
    else {
        wprintf(L"输入拦截器初始化失败\n");
    }
*/

//class WordTextMonitor {
//private:
//    IUIAutomation* m_pAutomation;
//    IUIAutomationElement* m_pDocumentElement;
//    IUIAutomationTextPattern* m_pTextPattern;
//    std::wstring m_lastText;
//
//public:
//    WordTextMonitor() : m_pAutomation(NULL), m_pDocumentElement(NULL), m_pTextPattern(NULL) {}
//
//    bool Initialize(HWND hwndDoc) {
//        HRESULT hr = CoCreateInstance(__uuidof(CUIAutomation), NULL,
//            CLSCTX_INPROC_SERVER, __uuidof(IUIAutomation),
//            (void**)&m_pAutomation);
//        if (FAILED(hr)) return false;
//
//        hr = m_pAutomation->ElementFromHandle(hwndDoc, &m_pDocumentElement);
//        if (FAILED(hr)) return false;
//
//        // 获取TextPattern
//        IUIAutomationTextPattern* pTextPattern = NULL;
//        hr = m_pDocumentElement->GetCurrentPattern(UIA_TextPatternId, (IUnknown**)&pTextPattern);
//        if (SUCCEEDED(hr)) {
//            m_pTextPattern = pTextPattern;
//            return true;
//        }
//        return false;
//    }
//
//    void CheckForChanges() {
//        if (!m_pTextPattern) return;
//
//        BSTR currentText;
//        IUIAutomationTextRange* textRange;
//        if (SUCCEEDED(m_pTextPattern->get_DocumentRange(&textRange))) {
//            if (SUCCEEDED(textRange->GetText(-1, &currentText))) {
//                std::wstring newText = currentText;
//                SysFreeString(currentText);
//
//                if (newText != m_lastText) {
//                    std::cout << "文本发生变化: " << wstr2str_2ANSI(newText) << std::endl;
//                    m_lastText = newText;
//                }
//            }
//        }
//    }
//};