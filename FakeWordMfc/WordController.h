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


// �����¼���������
class WordEventSink : public IDispatch
{
public:
    // IUnknown ����
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

    // IDispatch ����
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
        // �����¼��������ַ�
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
     * @param p ��Ҫ�󶨵Ľӿ�
     * @param eventUuid ��Ҫ�󶨵��¼���uuid
     */
    WordEventSink(IUnknown* p, const IID eventUuid) {
        this->ptr = p;
        this->eventUuid = eventUuid;
    }

    // һ����
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

            return pCP->Advise(this, &dwCookie); // �ҽ�
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
     * ע���¼�
     * @param eventDispIdMember ע����¼�����
     * @param callback �ص�����
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
     * �� Word �ĵ�
     * @param file Ҫ�򿪵��ĵ�·��
     * @param ConfirmConversions �Ƿ���ʾת���Ի���Ĭ�� false��
     * @param ReadOnly �Ƿ���ֻ����ʽ�򿪣�Ĭ�� false��
     * @param AddToRecentFiles �Ƿ���ӵ�����ļ��б�Ĭ�� true��
     * @param PasswordDocument ���ĵ��������루Ĭ�Ͽգ�
     * @param PasswordTemplate ��ģ���������루Ĭ�Ͽգ�
     * @param Revert ����ĵ��Ѵ򿪣��Ƿ����δ������Ĳ����´򿪣�Ĭ�� false��
     * @param WritePasswordDocument �����ĵ��������루Ĭ�Ͽգ�
     * @param WritePasswordTemplate ����ģ���������루Ĭ�Ͽգ�
     * @param Format �ĵ���ʽ��Ĭ�� wdOpenFormatAuto��
     * @param Encoding �ĵ����루Ĭ�� msoEncodingAutoDetect��
     * @param Visible �򿪺��Ƿ�ɼ���Ĭ�� true��
     * @param OpenConflictDocument �Ƿ�򿪳�ͻ�ĵ���Ĭ�� false��
     * @param OpenAndRepair �Ƿ����޸����ĵ���Ĭ�� false��
     * @param DocumentDirection �ĵ�����Ĭ�ϴ����ң�
     * @param NoEncodingDialog �Ƿ����ر���Ի���Ĭ�� false��
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
     * ע���¼��ַ���
     * @param parent �¼������Ľӿ�
     * @param eventUuid �¼�
     */
    Result<bool> registerEventSink(IUnknown* parent, const IID& eventUuid) {
        // ʵ�������洢�¼��۶���
        WordEventSink* pWordAppSink = new WordEventSink(parent, eventUuid);
        //// ��ȡWord�¼����ӵ�����
        //CComPtr<IConnectionPointContainer> icpc;
        //HRESULT hr = pWordApp.QueryInterface(IID_IConnectionPointContainer, &icpc);
        //Result icpcResult = Result<int>::CreateFromHRESULT(hr, 0);
        //if (!icpcResult)
        //{
        //    MessageBox(0, icpcResult.getMessage().c_str(), L"error", MB_ICONERROR);
        //    return;
        //}
        //// ��ȡWord�¼����ӵ�
        //CComPtr<IConnectionPoint> icp;
        //hr = icpc->FindConnectionPoint(__uuidof(Word::ApplicationEvents4), &icp);
        //Result icpResult = Result<int>::CreateFromHRESULT(hr, 0);
        //if (!icpResult)
        //{
        //    MessageBox(0, icpResult.getMessage().c_str(), L"error", MB_ICONERROR);
        //    return;
        //}
        //// ���¼��������󶨵����ӵ�
        //DWORD dwCookie;
        //hr = icp->Advise(pWordAppSink, &dwCookie);
        //Result adviseResult = Result<int>::CreateFromHRESULT(hr, 0);
        //if (!adviseResult)
        //{
        //    MessageBox(0, adviseResult.getMessage().c_str(), L"error", MB_ICONERROR);
        //    return;
        //}

        // ��EventSink��ʵ����Advise����˲���Ҫ����дһ���������̣���Ϊ���沽��
        HRESULT hr = dynamic_cast<WordEventSink*>(pWordAppSink)->attach();
        Result pSinkAttachResult = Result<bool>::CreateFromHRESULT(hr, 0);
        if (!pSinkAttachResult)
        {
            delete pWordAppSink;
            std::wstring msg = L"�޷�ע��Word�¼��ۣ�\n";
            msg += pSinkAttachResult.getMessage();
            MessageBox(0, msg.c_str(), L"Error", MB_ICONERROR);
        }
        else
            eventSinkMap[eventUuid] = pWordAppSink;
        return pSinkAttachResult;
    }

    /**
     * ע���¼�
     * @param eventUuid ֮ǰע���¼���ʹ�õ�eventUuid
     * @param 
     */
    Result<bool> registerEvent(const IID& eventUuid, DISPID dispIdMember, WordEventSink::DispatchInvokeCallbackFunction callback, WordEventSink::UserDataType userData) {
        if (!eventSinkMap.count(eventUuid))
        {
            LPOLESTR guidStr = NULL;
            HRESULT hr = StringFromCLSID(eventUuid, &guidStr);
            std::wstring errMsg = L"���¼���δע�᣺"; errMsg += guidStr ? guidStr : L"";
            std::wstring msg = L"�޷�ע��Word�¼���\n"; msg += errMsg;
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
    // ��ϣ����
    struct IIDHash {
        std::size_t operator()(const IID& iid) const noexcept {
            // �� 16 �ֽڵ������� 64 λ������ hash
            const std::uint64_t* p = reinterpret_cast<const std::uint64_t*>(&iid);
            std::hash<std::uint64_t> h;
            return h(p[0]) ^ (h(p[1]) + 0x9e3779b97f4a7c15ULL);
        }
    };

    // ��ȱȽϣ�IID �� POD��ֱ�Ӱ�λ�ȣ�
    struct IIDEqual {
        bool operator()(const IID& a, const IID& b) const noexcept {
            return InlineIsEqualGUID(a, b) != FALSE;   // �� MS �ṩ����������
            // ���� return memcmp(&a, &b, sizeof(GUID)) == 0;
        }
    };
    Word::_ApplicationPtr pWordApp;
    std::unordered_map<IdType, Word::_DocumentPtr> pWordDocumentMap;
    std::unordered_map<IdType, Word::WindowPtr> pWordWindowMap;
    std::unordered_map<const IID, WordEventSink*, IIDHash, IIDEqual> eventSinkMap;
};


// ģ�������һ��: nID������ǰEventSink��Ҫ�󶨶����ͬ�¼���Ψһ��ʶ���������Զ�������Sink Map���¼�������ͬ�����ַ�����
//class ATL_NO_VTABLE EventSink
//	: /*public IDispatch,*/
//    public CComObjectRootEx<CComSingleThreadModel>,
//    public IDispEventSimpleImpl</*nID =*/ 1, EventSink, &__uuidof(Word::ApplicationEvents4)>
//	//, public IDispEventSimpleImpl</*nID =*/ 2, EventSink, &__uuidof(Word::ApplicationEvents)> // ������� nID �� 1����ᱨ��������ͬ
//{
//private:
//    // ���� Word App
//    Word::_ApplicationPtr app;
//    // ����Cookie
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
//    // һ����
//    HRESULT Attach(Word::_ApplicationPtr pApp)
//    {
//        if (!pApp) return E_POINTER;
//        app = pApp;                       // ���浽��Ա����
//
//        CComQIPtr<IConnectionPointContainer> pCPC(pApp);
//        if (!pCPC) return E_NOINTERFACE;
//
//        CComPtr<IConnectionPoint> pCP;
//        HRESULT hr = pCPC->FindConnectionPoint(__uuidof(Word::ApplicationEvents4), &pCP);
//        if (FAILED(hr)) return hr;
//
//        return pCP->Advise(this->GetUnknown(), &dwCookie); // �ҽ�
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
//        app.Release();   // �ͷ�����ָ��
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