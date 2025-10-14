#pragma once

#pragma push_macro("FindText")
#pragma push_macro("ExitWindows")
#pragma push_macro("RGB")
#undef FindText
#undef ExitWindows
#undef RGB
#import "C:\\Program Files\\Microsoft Office\\Root\\VFS\\ProgramFilesCommonX64\\Microsoft Shared\\OFFICE16\\MSO.DLL"
#import "C:\\Program Files\\Microsoft Office\\Root\\VFS\\ProgramFilesCommonX86\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB"
#import "C:\\Program Files\\Microsoft Office\\root\\Office16\\MSWORD.OLB"/* rename("FindText", "FindText") rename("ExitWindows", "ExitWindows")*/
#pragma pop_macro("RGB")
#pragma pop_macro("ExitWindows")
#pragma pop_macro("FindText")

#include <atlbase.h>
#include <atlcom.h>
#include <string>
#include "StringProcess.h"
//#include <any>
#include <iostream>
#include <functional>
#include <thread>
#include <future>
#include <any>

#include "ComResultAndCatch.h"

std::wstring HResultToWString(HRESULT hr);


// 定义事件接收器类
class WordEventSink : public IDispatch
{
public:
    // IUnknown 方法
    STDMETHOD(QueryInterface)(REFIID riid, void** ppvObject) override
    {
        if (riid == IID_IUnknown || riid == IID_IDispatch || riid == eventUuid)
        {
            *ppvObject = static_cast<IDispatch*>(this);
            AddRef();
            return S_OK;
        }
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    STDMETHOD_(ULONG, AddRef)() override
    {
        return InterlockedIncrement(&_refCount);
    }

    STDMETHOD_(ULONG, Release)() override
    {
        ULONG count = InterlockedDecrement(&_refCount);
        if (count == 0)
        {
            delete this;
        }
        return count;
    }

    // IDispatch 方法
    STDMETHOD(GetTypeInfoCount)(UINT* pctinfo) override
    {
        *pctinfo = 0;
        return S_OK;
    }

    STDMETHOD(GetTypeInfo)(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override
    {
        *ppTInfo = nullptr;
        return E_NOTIMPL;
    }

    STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) override
    {
        return E_NOTIMPL;
    }

    STDMETHOD(Invoke)(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr) override
    {
        std::cout << "dispIdMember: " << dispIdMember << std::endl;
        // 进行事件处理函数分发
        if (callbackInfoMap.count(dispIdMember))
        {
            auto [begin, end] = callbackInfoMap.equal_range(dispIdMember);
            for (auto &it = begin; it != end; it++)
            {
                auto& info = it->second;
                info.callback(pDispParams, info.userData);
            }
            return S_OK;
        }
        return E_NOTIMPL;
    }

public:
    //enum class CallbackExecType {
    //    Sync,
    //    Async
    //};
    typedef std::any UserDataType;
    typedef std::function<Result<bool>(DISPPARAMS* pDispParams, UserDataType userData)> DispatchInvokeCallbackFunction;

private:
    struct DispatchInvokeCallbackInfo {
        DISPID dispIdMember;
        //CallbackExecType execType;
        DispatchInvokeCallbackFunction callback;
        UserDataType userData;
    };
    typedef std::unordered_multimap<DISPID, DispatchInvokeCallbackInfo> CallbackInfoMapType;
    ULONG _refCount = 1;
    IUnknown* ptr = nullptr;
    DWORD dwCookie = 0;
    IID eventUuid;
    CallbackInfoMapType callbackInfoMap;
public:
    /**
     * @param p 需要绑定的接口
     * @param eventUuid 需要绑定的事件的uuid
     */
    WordEventSink(IUnknown* p, const IID eventUuid) {
        this->ptr = p;
        this->eventUuid = eventUuid;
    }

    // 一键绑定
    HRESULT attach()
    {
        try {
            if (!ptr)
                return E_POINTER;
            if (dwCookie)
                return E_PENDING;

            CComQIPtr<IConnectionPointContainer> pCPC(ptr);
            if (!pCPC)
                return E_NOINTERFACE;

            CComPtr<IConnectionPoint> pCP;
            HRESULT hr = pCPC->FindConnectionPoint(eventUuid, &pCP);
            if (FAILED(hr))
                return hr;

            return pCP->Advise(this, &dwCookie); // 挂接
        }
        catch (...) {
            return E_POINTER;
        }
    }

    void detach()
    {
        try {
            if (dwCookie != 0)
            {
                CComQIPtr<IConnectionPointContainer> pCPC(ptr);
                if (pCPC)
                {
                    CComPtr<IConnectionPoint> pCP;
                    if (SUCCEEDED(pCPC->FindConnectionPoint(eventUuid, &pCP)))
                        pCP->Unadvise(dwCookie);
                }
                dwCookie = 0;
            }
        }
        catch (...) {
        }
    }

    /**
     * 注册事件
     * @param eventDispIdMember 注册的事件号码
     * @param callback 回调函数
     */
    void registerEvent(DISPID dispIdMember, DispatchInvokeCallbackFunction callback, UserDataType userData) {
        DispatchInvokeCallbackInfo info{
            dispIdMember,
            callback,
            userData
        };
        this->callbackInfoMap.insert(CallbackInfoMapType::value_type(dispIdMember, info));
    }
};

#define ERRORRESULT_IDNOTFOUNT(type) ErrorResult<type>(E_FAIL, L"Id not found")

class WordController
{
    typedef unsigned long long IdType;

public:
	~WordController();
	WordController();
    
    static Result<bool> comInitialize() {
        HRESULT hr = CoInitializeEx(0, COINIT_MULTITHREADED);
        //HRESULT hr = CoInitialize(0);
        return Result<bool>::CreateFromHRESULT(hr, 0);
    }
    static void comUninitialize() {
        CoUninitialize();
    }

    Word::_ApplicationPtr getApp() const {
        return pWordApp;
    }

    Result<Word::_DocumentPtr> getDocument(IdType id) const {
        if (pWordDocumentMap.count(id))
            return SuccessResult(0, pWordDocumentMap.at(id));
        return ERRORRESULT_IDNOTFOUNT(Word::_DocumentPtr);
    }


    Result<Word::WindowPtr> getWindow(IdType id) const {
        if (pWordWindowMap.count(id))
            return SuccessResult(0, pWordWindowMap.at(id));
        return ERRORRESULT_IDNOTFOUNT(Word::WindowPtr);
    }



    Result<Word::_DocumentPtr> getActiveDocument() const;


    Result<Word::_ApplicationPtr> createApp();
    /**
     * 打开 Word 文档
     * @param file 要打开的文档路径
     * @param ConfirmConversions 是否显示转换对话框（默认 false）
     * @param ReadOnly 是否以只读方式打开（默认 false）
     * @param AddToRecentFiles 是否添加到最近文件列表（默认 true）
     * @param PasswordDocument 打开文档所需密码（默认空）
     * @param PasswordTemplate 打开模板所需密码（默认空）
     * @param Revert 如果文档已打开，是否放弃未保存更改并重新打开（默认 false）
     * @param WritePasswordDocument 保存文档所需密码（默认空）
     * @param WritePasswordTemplate 保存模板所需密码（默认空）
     * @param Format 文档格式（默认 wdOpenFormatAuto）
     * @param Encoding 文档编码（默认 msoEncodingAutoDetect）
     * @param Visible 打开后是否可见（默认 true）
     * @param OpenConflictDocument 是否打开冲突文档（默认 false）
     * @param OpenAndRepair 是否尝试修复损坏文档（默认 false）
     * @param DocumentDirection 文档方向（默认从左到右）
     * @param NoEncodingDialog 是否隐藏编码对话框（默认 false）
     */
    Result<Word::_DocumentPtr> openDocument(
        IdType id,
        std::wstring file,
        bool ConfirmConversions = false,
        bool ReadOnly = false,
        bool AddToRecentFiles = true,
        std::wstring PasswordDocument = L"",
        std::wstring PasswordTemplate = L"",
        bool Revert = false,
        std::wstring WritePasswordDocument = L"",
        std::wstring WritePasswordTemplate = L"",
        Word::WdOpenFormat Format = Word::WdOpenFormat::wdOpenFormatAuto,
        Office::MsoEncoding Encoding = Office::MsoEncoding::msoEncodingAutoDetect,
        bool Visible = true,
        bool OpenConflictDocument = false,
        bool OpenAndRepair = false,
        Word::WdDocumentDirection DocumentDirection = Word::WdDocumentDirection::wdLeftToRight,
        bool NoEncodingDialog = false
    );

    Result<Word::_DocumentPtr> addDocument(IdType id, bool visible = false, std::wstring templateName = L"", bool asNewTemplate = false, Word::WdNewDocumentType documentType = Word::wdNewBlankDocument);
    Result<Word::WindowPtr> newWindow(IdType id);

    Result<bool> setAppVisible(bool visible) const;

    /**
     * 注册事件分发槽
     * @param parent 事件隶属的接口
     * @param eventUuid 事件
     */
    Result<bool> registerEventSink(IUnknown* parent, const IID& eventUuid) {
        // 实例化并存储事件槽对象
        WordEventSink* pWordAppSink = new WordEventSink(parent, eventUuid);
        //// 获取Word事件连接点容器
        //CComPtr<IConnectionPointContainer> icpc;
        //HRESULT hr = pWordApp.QueryInterface(IID_IConnectionPointContainer, &icpc);
        //Result icpcResult = Result<int>::CreateFromHRESULT(hr, 0);
        //if (!icpcResult)
        //{
        //    MessageBox(0, icpcResult.getMessage().c_str(), L"error", MB_ICONERROR);
        //    return;
        //}
        //// 获取Word事件连接点
        //CComPtr<IConnectionPoint> icp;
        //hr = icpc->FindConnectionPoint(__uuidof(Word::ApplicationEvents4), &icp);
        //Result icpResult = Result<int>::CreateFromHRESULT(hr, 0);
        //if (!icpResult)
        //{
        //    MessageBox(0, icpResult.getMessage().c_str(), L"error", MB_ICONERROR);
        //    return;
        //}
        //// 将事件接收器绑定到连接点
        //DWORD dwCookie;
        //hr = icp->Advise(pWordAppSink, &dwCookie);
        //Result adviseResult = Result<int>::CreateFromHRESULT(hr, 0);
        //if (!adviseResult)
        //{
        //    MessageBox(0, adviseResult.getMessage().c_str(), L"error", MB_ICONERROR);
        //    return;
        //}

        // 在EventSink中实现了Advise，因此不需要重新写一遍上述流程，改为下面步骤
        HRESULT hr = dynamic_cast<WordEventSink*>(pWordAppSink)->attach();
        Result pSinkAttachResult = Result<bool>::CreateFromHRESULT(hr, 0);
        if (!pSinkAttachResult)
        {
            delete pWordAppSink;
            std::wstring msg = L"无法注册Word事件槽！\n";
            msg += pSinkAttachResult.getMessage();
            MessageBox(0, msg.c_str(), L"Error", MB_ICONERROR);
        }
        else
            eventSinkMap[eventUuid] = pWordAppSink;
        return pSinkAttachResult;
    }

    /**
     * 注册事件
     * @param eventUuid 之前注册事件槽使用的eventUuid
     * @param 
     */
    Result<bool> registerEvent(const IID& eventUuid, DISPID dispIdMember, WordEventSink::DispatchInvokeCallbackFunction callback, WordEventSink::UserDataType userData) {
        if (!eventSinkMap.count(eventUuid))
        {
            LPOLESTR guidStr = NULL;
            HRESULT hr = StringFromCLSID(eventUuid, &guidStr);
            std::wstring errMsg = L"该事件槽未注册："; errMsg += guidStr ? guidStr : L"";
            std::wstring msg = L"无法注册Word事件！\n"; msg += errMsg;
            if (SUCCEEDED(hr))
            {
                CoTaskMemFree(guidStr);
            }
            MessageBox(0, msg.c_str(), L"Error", MB_ICONERROR);
            return ErrorResult<bool>(0, errMsg);
        }
        eventSinkMap[eventUuid]->registerEvent(dispIdMember, callback, userData);
        return SuccessResult(true);
    }

private:
    // 哈希函数
    struct IIDHash {
        std::size_t operator()(const IID& iid) const noexcept {
            // 把 16 字节当成两个 64 位整数来 hash
            const std::uint64_t* p = reinterpret_cast<const std::uint64_t*>(&iid);
            std::hash<std::uint64_t> h;
            return h(p[0]) ^ (h(p[1]) + 0x9e3779b97f4a7c15ULL);
        }
    };

    // 相等比较（IID 是 POD，直接按位比）
    struct IIDEqual {
        bool operator()(const IID& a, const IID& b) const noexcept {
            return InlineIsEqualGUID(a, b) != FALSE;   // 用 MS 提供的内联函数
            // 或者 return memcmp(&a, &b, sizeof(GUID)) == 0;
        }
    };
    Word::_ApplicationPtr pWordApp;
    std::unordered_map<IdType, Word::_DocumentPtr> pWordDocumentMap;
    std::unordered_map<IdType, Word::WindowPtr> pWordWindowMap;
    std::unordered_map<const IID, WordEventSink*, IIDHash, IIDEqual> eventSinkMap;
};


// 模板参数第一个: nID，代表当前EventSink类要绑定多个不同事件的唯一标识符，可以自定义用于Sink Map的事件（可能同名）分发区分
//class ATL_NO_VTABLE EventSink
//	: /*public IDispatch,*/
//    public CComObjectRootEx<CComSingleThreadModel>,
//    public IDispEventSimpleImpl</*nID =*/ 1, EventSink, &__uuidof(Word::ApplicationEvents4)>
//	//, public IDispEventSimpleImpl</*nID =*/ 2, EventSink, &__uuidof(Word::ApplicationEvents)> // 如果这里 nID 是 1，则会报错：基类相同
//{
//private:
//    // 保存 Word App
//    Word::_ApplicationPtr app;
//    // 保存Cookie
//    DWORD dwCookie = 0;
//    // Note declaration of structure for type information
//    static _ATL_FUNC_INFO OnDocChangeInfo;
//    static _ATL_FUNC_INFO OnQuitInfo;
//    
//public:
//	~EventSink() {}
//	EventSink() {}
//
//    BEGIN_COM_MAP(EventSink)
//
//    END_COM_MAP()
//    // Note the mapping from Word events to our event handler functions.
//    BEGIN_SINK_MAP(EventSink)
//        SINK_ENTRY_INFO(/*nID =*/ 1, __uuidof(Word::ApplicationEvents4), /*dispid =*/ 3, OnDocumentChange, &OnDocChangeInfo)
//        SINK_ENTRY_INFO(/*nID =*/ 1, __uuidof(Word::ApplicationEvents4), /*dispid =*/ 2, OnQuit, &OnQuitInfo)
//    END_SINK_MAP()
//
//    // 一键绑定
//    HRESULT Attach(Word::_ApplicationPtr pApp)
//    {
//        if (!pApp) return E_POINTER;
//        app = pApp;                       // 保存到成员变量
//
//        CComQIPtr<IConnectionPointContainer> pCPC(pApp);
//        if (!pCPC) return E_NOINTERFACE;
//
//        CComPtr<IConnectionPoint> pCP;
//        HRESULT hr = pCPC->FindConnectionPoint(__uuidof(Word::ApplicationEvents4), &pCP);
//        if (FAILED(hr)) return hr;
//
//        return pCP->Advise(this->GetUnknown(), &dwCookie); // 挂接
//    }
//
//    void Detach()
//    {
//        if (dwCookie != 0)
//        {
//            CComQIPtr<IConnectionPointContainer> pCPC(app);
//            if (pCPC)
//            {
//                CComPtr<IConnectionPoint> pCP;
//                if (SUCCEEDED(pCPC->FindConnectionPoint(__uuidof(Word::ApplicationEvents4), &pCP)))
//                    pCP->Unadvise(dwCookie);
//            }
//            dwCookie = 0;
//        }
//        app.Release();   // 释放智能指针
//    }
//
//private:
//    // Event handlers
//    // Note the __stdcall calling convention and 
//    // dispinterface-style signature
//    void __stdcall OnQuit()
//    {
//        MessageBox(0, L"Word application quit.", L"Quit", 0);
//    }
//
//    void __stdcall OnDocumentChange()
//    {
//        // Get a pointer to the _Document interface on the active document
//        CComPtr<Word::_Document> pDoc;
//
//        // Get the name from the active document
//        CComBSTR bstrName;
//        if (pDoc)
//            pDoc->get_Name(&bstrName);
//
//        // Create a display string
//        CComBSTR bstrDisplay(_T("New document title:\n"));
//        bstrDisplay += bstrName;
//
//        // Display the name to the user
//        USES_CONVERSION;
//        MessageBox(NULL, W2CT(bstrDisplay), _T("IDispEventSimpleImpl : Active Document Changed"), MB_OK);
//    }
//
//
//};