#include <windows.h>
#include <msctf.h>
#include <string>
#include <vector>
#include <iostream>
#include "StringProcess.h"
#ifndef GET_X_LPARAM
#define GET_X_LPARAM(lp)                        ((int)(short)LOWORD(lp))
#endif
#ifndef GET_Y_LPARAM
#define GET_Y_LPARAM(lp)                        ((int)(short)HIWORD(lp))
#endif
#include "FileLogger.h"
#include <comdef.h>


class CTextStore;

/**
 * @brief CTSFManager
 *
 * 管理 Text Services Framework (TSF) 的生命周期和与编辑控件的交互。
 * 负责初始化/反初始化 TSF、处理来自 TSF 的文本输入，并维护必要的 TSF 对象句柄。
 */
class CTSFManager
{
private:
    ITfThreadMgr* m_pThreadMgr;
    ITfDocumentMgr* m_pDocumentMgr;
    ITfContext* m_pContext;
    DWORD m_ContextCookie = 0;
    CTextStore* m_pTextStore;
    TfClientId m_clientId;
    HWND m_hwndEdit;
    FileLogger& logger;

public:
    /**
     * @brief 构造函数
     *
     * 初始化内部指针为 nullptr，并将 clientId 和窗口句柄设为默认值。
     */
    CTSFManager(FileLogger& logger) :
        logger(logger),
        m_pThreadMgr(nullptr),
        m_pDocumentMgr(nullptr),
        m_pContext(nullptr),
        m_pTextStore(nullptr),
        m_clientId(TF_CLIENTID_NULL),
        m_hwndEdit(nullptr)
    {
    }

    /**
     * @brief 析构函数
     *
     * 在对象销毁时确保释放并反初始化 TSF 相关资源。
     */
    ~CTSFManager()
    {
        Uninitialize();
    }

    /**
     * @brief 初始化 TSF 管理器并关联到指定的编辑窗口
     *
     * @param hwndEdit 要与 TSF 交互的编辑控件窗口句柄
     * @return TRUE 初始化成功，FALSE 初始化失败
     */
    BOOL Initialize(HWND hwndEdit);

    /**
     * @brief 反初始化 TSF 管理器
     *
     * 释放所有 TSF 接口和内部资源。
     */
    void Uninitialize();

private:
};

/**
 * @brief CTextStore
 *
 * 实现 ITextStoreACP 接口，负责向 TSF 报告编辑缓冲区文本与选区信息，
 * 并接收来自 TSF 的输入/锁定请求。将大部分操作委托给 CTSFManager 以完成 UI 层的更新。
 */
class CTextStore : public ITextStoreACP
    , public ITfContextOwnerCompositionSink
    , public ITfCompositionSink
{
private:
    FileLogger& logger;

    LONG m_cRef;
    HWND m_hwndEdit;
    ITfContext* m_pContext;
    TfClientId m_clientId;
    DWORD m_dwLockType;
    ITextStoreACPSink* m_pSink = nullptr; // 存储槽

    // 文本内容
    std::wstring m_text;
    TS_SELECTION_ACP m_selection;
    TsViewCookie m_viewCookie;

    // Composition
    DWORD m_CompSinkCookie = 0;
    std::wstring m_compText;
    BOOL isCompFinished = FALSE;
public:
    /**
     * @brief 构造函数
     *
     * @param pManager 指向上层 CTSFManager 的指针
     * @param hwndEdit 关联的编辑窗口句柄
     */
    CTextStore(FileLogger& logger, CTSFManager* pManager, HWND hwndEdit) :
        logger(logger),
        m_cRef(1),
        m_hwndEdit(hwndEdit),
        m_pContext(nullptr),
        m_clientId(TF_CLIENTID_NULL),
        m_dwLockType(0),
        m_pSink(nullptr),
        m_viewCookie(0),
        m_compSel({ 0, 0, {TS_AE_NONE} })
    {
    logger << "[CTextStore::CTextStore] ctor hwndEdit=" << (void*)hwndEdit << "\n";
        m_selection.acpStart = 0;
        m_selection.acpEnd = 0;
        m_selection.style.ase = TS_AE_NONE;
        m_selection.style.fInterimChar = FALSE;
    }

    /**
     * @brief 析构函数
     *
     * 释放对 ITextStoreACPSink 和 ITfContext 的引用
     */
    virtual ~CTextStore()
    {
    logger << "[CTextStore::~CTextStore] dtor\n";
        if (m_pSink)
        {
            m_pSink->Release();
            m_pSink = nullptr;
        }
        if (m_pContext)
        {
            m_pContext->Release();
        }
    }

    /**
     * @brief 初始化文本存储
     *
     * 保存并 AddRef 提供的 ITfContext，同时记录 clientId。
     * @param pContext TSF 上下文
     * @param clientId 注册的客户端 ID
     */
    void Initialize(ITfContext* pContext, TfClientId clientId)
    {
    logger << "[CTextStore::Initialize] pContext=" << (void*)pContext << " clientId=" << clientId << "\n";
        m_pContext = pContext;
        m_pContext->AddRef();
        m_clientId = clientId;

        ITfSource* pSource = nullptr;
        HRESULT hr = m_pContext->QueryInterface(&pSource);
        if (SUCCEEDED(hr))
        {
            //if (m_CompSinkCookie != TF_INVALID_COOKIE && m_CompSinkCookie != 0)
            //{
            //    logger << L"[CTextStore::Initialize] Warning: Already registered, unregistering first\n";
            //    hr = pSource->UnadviseSink(m_CompSinkCookie);
            //    m_CompSinkCookie = TF_INVALID_COOKIE;
            //    if (FAILED(hr))
            //    {
            //        _com_error err(hr);
            //        logger << L"[CTextStore::Initialize] pSource->UnadviseSink error: " << err.ErrorMessage() << L", code: " << hr;
            //    }
            //}
            hr = pSource->AdviseSink(IID_ITfContextOwnerCompositionSink,
                static_cast<ITfContextOwnerCompositionSink*>(this),
                &m_CompSinkCookie);   // DWORD 成员，析构时 Unadvise
            if (FAILED(hr))
            {
                _com_error err(hr);
                logger << L"[CTextStore::Initialize] pSource->AdviseSink IID_ITfContextOwnerCompositionSink error: " << err.ErrorMessage() << L", code: " << hr;
            }
            //hr = pSource->AdviseSink(IID_ITfCompositionSink,
            //    static_cast<ITfCompositionSink*>(this),
            //    &m_CompSinkCookie);   // DWORD 成员，析构时 Unadvise
            //if (FAILED(hr))
            //{
            //    _com_error err(hr);
            //    logger << L"[CTextStore::Initialize] pSource->AdviseSink IID_ITfCompositionSink error: " << err.ErrorMessage() << L", code: " << hr;
            //}
            logger << L"[CTextStore::Initialize] Initialize successful.";
            pSource->Release();
        }
        else {
            _com_error err(hr);
            logger << L"[CTextStore::Initialize] m_pContext->QueryInterface pSource error: " << err.ErrorMessage() << L", code: " << hr;
        }
    }

    /**
     * @brief IUnknown::QueryInterface
     *
     * 查询并返回支持的接口指针（IUnknown 或 ITextStoreACP）。
     * @param riid 请求的接口 ID
     * @param ppvObj 输出接口指针
     * @return S_OK 或 E_NOINTERFACE
     */
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override
    {
    logger << "[CTextStore::QueryInterface] riid==IID_IUnknown? " << (riid == IID_IUnknown) << " riid==IID_ITextStoreACP? " << (riid == IID_ITextStoreACP) << "\n";
        if (riid == IID_IUnknown || riid == IID_ITextStoreACP)
        {
            *ppvObj = static_cast<ITextStoreACP*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == IID_ITfContextOwnerCompositionSink)
        {
            *ppvObj = static_cast<ITfContextOwnerCompositionSink*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == IID_ITfCompositionSink)
        {
            *ppvObj = static_cast<ITfCompositionSink*>(this);
            AddRef();
            return S_OK;
        }
        *ppvObj = nullptr;
        return E_NOINTERFACE;
    }

    /**
     * @brief IUnknown::AddRef
     *
     * 增加引用计数并返回新的计数值。
     * @return 新的引用计数
     */
    STDMETHODIMP_(ULONG) AddRef() override
    {
    ULONG ret = InterlockedIncrement(&m_cRef);
    logger << "[CTextStore::AddRef] newRef=" << ret << "\n";
    return ret;
    }

    /**
     * @brief IUnknown::Release
     *
     * 减少引用计数，计数为 0 时销毁对象，返回剩余计数。
     * @return 剩余引用计数
     */
    STDMETHODIMP_(ULONG) Release() override
    {
        LONG cRef = InterlockedDecrement(&m_cRef);
        logger << "[CTextStore::Release] newRef=" << cRef << "\n";
        if (cRef == 0)
        {
            logger << "[CTextStore::Release] deleting self\n";
            delete this;
        }
        return cRef;
    }

    /**
     * @brief ITextStoreACP::AdviseSink
     *
     * 注册或替换 ITextStoreACPSink 回调。TSF 会通过该回调通知锁定等事件。
     * @param riid 必须为 IID_ITextStoreACPSink
     * @param punk 提供的 sink 对象
     * @param dwMask 面具（未使用）
     * @return S_OK 或错误码
     */
    STDMETHODIMP AdviseSink(REFIID riid, IUnknown* punk, DWORD dwMask) override
    {
    logger << "[CTextStore::AdviseSink] riid==IID_ITextStoreACPSink? " << (riid == IID_ITextStoreACPSink) << " punk=" << punk << " dwMask=" << dwMask << "\n";
        if (riid != IID_ITextStoreACPSink)
            return E_INVALIDARG;

        if (m_pSink)
        {
            m_pSink->Release();
            m_pSink = nullptr;
        }

        HRESULT hr = punk->QueryInterface(IID_ITextStoreACPSink, (void**)&m_pSink);
        if (FAILED(hr))
        {
            _com_error err(hr);
            logger << L"[CTextStore::AdviseSink] punk->QueryInterface IID_ITextStoreACPSink error: " << err.ErrorMessage() << L", HRESULT code: " << hr;
        }
        return hr;
    }

    /**
     * @brief ITextStoreACP::UnadviseSink
     *
     * 取消注册先前注册的 ITextStoreACPSink。
     * @param punk 可选，通常忽略，直接释放保存的 m_pSink 即可。
     * @return S_OK
     */
    STDMETHODIMP UnadviseSink(IUnknown* punk) override
    {
    logger << "[CTextStore::UnadviseSink] punk=" << punk << "\n";
        if (m_pSink)
        {
            m_pSink->Release();
            m_pSink = nullptr;
        }
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::GetStatus
     *
     * 返回文本存储的状态信息（如是否是瞬态文本、是否有隐藏文本等）。
     * @param pdcs 输出的 TS_STATUS 结构
     * @return S_OK 或 E_INVALIDARG
     */
    STDMETHODIMP GetStatus(TS_STATUS* pdcs) override
    {
    logger << "[CTextStore::GetStatus] pdcs=" << pdcs << "\n";
        if (!pdcs)
            return E_INVALIDARG;

        pdcs->dwDynamicFlags = 0;
        pdcs->dwStaticFlags = TS_SS_TRANSITORY | TS_SS_NOHIDDENTEXT;
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::QueryInsert
     *
     * 用于询问在给定位置插入指定长度文本是否可行，返回实际将要插入的位置范围。
     * @param acpTestStart/acpTestEnd 测试的插入范围
     * @param cch 字符长度
     * @param pacpResultStart/pacpResultEnd 输出实际位置
     * @return S_OK 或 E_INVALIDARG
     */
    STDMETHODIMP QueryInsert(LONG acpTestStart, LONG acpTestEnd, ULONG cch,
        LONG* pacpResultStart, LONG* pacpResultEnd) override
    {
    logger << "[CTextStore::QueryInsert] acpTestStart=" << acpTestStart << " acpTestEnd=" << acpTestEnd << " cch=" << cch << "\n";
        if (!pacpResultStart || !pacpResultEnd)
            return E_INVALIDARG;

        *pacpResultStart = acpTestStart;
        *pacpResultEnd = acpTestEnd;
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::GetSelection
     *
     * 返回当前的选区信息（以字符偏移 ACP 表示）。
     * @param ulIndex/ulCount 索引和请求数量
     * @param pSelection 输出缓冲
     * @param pcFetched 实际返回数量
     * @return S_OK 或 E_INVALIDARG / TS_E_NOSELECTION
     */
    STDMETHODIMP GetSelection(ULONG ulIndex, ULONG ulCount,
        TS_SELECTION_ACP* pSelection, ULONG* pcFetched) override
    {
    logger << "[CTextStore::GetSelection] ulIndex=" << ulIndex << " ulCount=" << ulCount << "\n";
        // 参数验证
        if (!pSelection || !pcFetched)
            return E_INVALIDARG;

        *pcFetched = 0;

        // 检查索引：支持默认选择(0xFFFFFFFF)或第一个选择(0)
        if (ulIndex != 0 && ulIndex != TS_DEFAULT_SELECTION)
            return TS_E_NOSELECTION;

        if (ulCount == 0)
            return E_INVALIDARG;

        // 获取当前编辑控件的选择状态
        DWORD dwStart = 0, dwEnd = 0;

        if (IsWindow(m_hwndEdit))
        {
            // 从编辑控件获取实际选择
            SendMessage(m_hwndEdit, EM_GETSEL, (WPARAM)&dwStart, (LPARAM)&dwEnd);

            // 更新内部选择状态
            m_selection.acpStart = static_cast<LONG>(dwStart);
            m_selection.acpEnd = static_cast<LONG>(dwEnd);
        }

        // 验证选择范围
        //LONG textLength = static_cast<LONG>(m_text.length());
        //if (m_selection.acpStart < 0)
        //    m_selection.acpStart = 0;
        //if (m_selection.acpStart > textLength)
        //    m_selection.acpStart = textLength;
        //if (m_selection.acpEnd < m_selection.acpStart)
        //    m_selection.acpEnd = m_selection.acpStart;
        //if (m_selection.acpEnd > textLength)
        //    m_selection.acpEnd = textLength;

        // 对于默认选择请求，确保返回一个有效的插入点
        if (ulIndex == TS_DEFAULT_SELECTION)
        {
            // 通常默认选择就是当前光标位置（插入点）
            // 确保选择是一个插入点（开始=结束）
            //m_selection.acpEnd = m_selection.acpStart;
        }

        // 输出结果
        if (m_compSel.acpStart < 0 || m_compSel.acpEnd < 0)
            m_compSel.acpStart = m_compSel.acpEnd = 0;
        pSelection[0] = m_compSel;
        *pcFetched = 1;
        logger << "[GetSelection] Current Selection: " << m_selection.acpStart << " -> " << m_selection.acpEnd << "\n";
        return S_OK;
    }
    /**
     * @brief ITextStoreACP::SetSelection
     *
     * 设置当前选区信息（通常由 TSF 控制）。
     * @param ulCount 选区数量（这里只支持 1）
     * @param pSelection 输入选区
     * @return S_OK 或 E_INVALIDARG / E_NOTIMPL
     */
    STDMETHODIMP SetSelection(ULONG ulCount, const TS_SELECTION_ACP* pSelection) override
    {
    logger << "[CTextStore::SetSelection] ulCount=" << ulCount << " pSelection=" << pSelection << "\n";
        if (!pSelection)
            return E_INVALIDARG;
        //if (ulCount != 1)
        //    return E_NOTIMPL;
        m_compSel = pSelection[0];
        if (m_pSink)
            m_pSink->OnSelectionChange();
        logger << "[SetSelection] Current Selection: " << m_selection.acpStart << " -> " << m_selection.acpEnd << "\n";
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::GetText
     *
     * 从内部文本缓冲区读取指定范围的文本（ACP 偏移）。
     * @param acpStart/acpEnd 读取范围（acpEnd 为 -1 表示直到文本末尾）
     * @param pchPlain 输出缓冲
     * @param cchPlainReq 缓冲大小
     * @param pcchPlainOut 实际写入字符数
     * @param prgRunInfo 文本运行信息（可选）
     * @param ulRunInfoReq 运行信息请求大小
     * @param pulRunInfoOut 运行信息实际返回大小（可选）
     * @param pacpNext 下一个可读取位置（可选）
     * @return S_OK 或 TF_E_INVALIDPOS / E_INVALIDARG
     */
    STDMETHODIMP GetText(LONG acpStart, LONG acpEnd, WCHAR* pchPlain, ULONG cchPlainReq,
        ULONG* pcchPlainOut, TS_RUNINFO* prgRunInfo, ULONG ulRunInfoReq,
        ULONG* pulRunInfoOut, LONG* pacpNext) override
    {
    logger << "[CTextStore::GetText] acpStart=" << acpStart << " acpEnd=" << acpEnd << " cchPlainReq=" << cchPlainReq << "\n";
        // 初始化参数
        if (prgRunInfo)
        {
            prgRunInfo->type = TS_RT_PLAIN;
            prgRunInfo->uCount = 0;
        }
        if (pulRunInfoOut)
            *pulRunInfoOut = 0;
        if (pacpNext)
            *pacpNext = 0;
        if (!pcchPlainOut)
            return E_INVALIDARG;
        auto tLen = GetWindowTextLength(m_hwndEdit) + 1;

        WCHAR* t = new WCHAR[tLen];
        GetWindowText(m_hwndEdit, t, tLen);
        m_text = t;
        delete[] t;
        // 计算文本范围
        LONG acpEndAdjusted = acpEnd;
        if (acpEnd == -1)
            acpEndAdjusted = static_cast<LONG>(m_compText.length());

        if (acpStart < 0 || acpStart > acpEndAdjusted || acpEndAdjusted > static_cast<LONG>(m_compText.length()))
            return TF_E_INVALIDPOS;

        ULONG cchToCopy = acpEndAdjusted - acpStart;
        if (cchToCopy > cchPlainReq)
            cchToCopy = cchPlainReq;

        // 复制文本
        if (pchPlain && cchToCopy > 0)
        {
            memcpy(pchPlain, m_compText.c_str() + acpStart, cchToCopy * sizeof(WCHAR));
        }

        *pcchPlainOut = cchToCopy;
    logger << "[CTextStore::GetText] pcchPlainOut=" << *pcchPlainOut << "\n";

        if (pulRunInfoOut)
            *pulRunInfoOut = 0;

        if (pacpNext)
            *pacpNext = acpStart + cchToCopy;

        return S_OK;
    }

    /**
     * @brief ITextStoreACP::SetText
     *
     * TSF 请求在指定范围设置文本时调用。此实现会将文本通知到上层管理器进行处理。
     * @param dwFlags 标志
     * @param acpStart/acpEnd 目标范围
     * @param pchText 新文本缓冲
     * @param cch 文本字符数
     * @param pChange 输出的变更信息（可选）
     * @return S_OK
     */
    STDMETHODIMP SetText(DWORD dwFlags, LONG acpStart, LONG acpEnd,
        const WCHAR* pchText, ULONG cch, TS_TEXTCHANGE* pChange) override
    {
        logger << "[SetText] dwFlags: " << dwFlags << "\n";
        logger << "[SetText] pchText: " << (pchText ? wstr2str_2ANSI(pchText) : "") << "\n";
        logger << "[SetText] Current Selection: " << m_selection.acpStart << " -> " << m_selection.acpEnd << "\n";
        logger << "[SetText] Current Selection Param: " << acpStart << " -> " << acpEnd << "\n";
        // 处理文本设置
        if (pchText && cch > 0)
        {
            // 更新文本
            //size_t tLen = 256;
            //wchar_t* t = new wchar_t[tLen];
            //for (int rst = 0; (rst = ::GetWindowText(m_hwndEdit, t, tLen)) != 0 && rst == tLen - 1;)
            //{
            //    delete[] t;
            //    tLen += 256;
            //    t = new WCHAR[tLen];
            //}
            //m_text = t;
            //delete[] t;
            //SendMessage(m_hwndEdit, EM_GETSEL, (WPARAM) & m_selection.acpStart, (LPARAM)&m_selection.acpEnd);
            //std::wstring newText(pchText, cch);
            //auto selStart = m_selection.acpStart;
            //auto selEnd = m_selection.acpEnd;
            //auto selLen = selEnd - selStart;
            //m_text.replace(selStart, selLen, newText);
            //SetWindowText(m_hwndEdit, m_text.c_str());
            // SendMessage(m_hwndEdit, EM_SETSEL, selStart, selEnd);
            m_compText.replace(acpStart, acpEnd - acpStart, pchText);
            if (m_isComp)
                isCompFinished = TRUE;
            else
                isCompFinished = FALSE;

        }
        if (cch == 0)
        {
            m_compText.clear();
            m_compSel.acpStart = m_compSel.acpEnd = 0;
        }
        if (pChange)
        {
            pChange->acpStart = acpStart;
            pChange->acpOldEnd = acpEnd;
            pChange->acpNewEnd = acpStart + cch;
            if (m_pSink)
                m_pSink->OnTextChange(dwFlags, pChange);
        }
        return S_OK;
    }

    /**+
     * @brief ITextStoreACP::GetFormattedText
     *
     * 可返回指定范围内的富文本表示（例如剪贴板/拖放用），当前不实现。
     */
    STDMETHODIMP GetFormattedText(LONG acpStart, LONG acpEnd, IDataObject** ppDataObject) override
    {
    logger << "[CTextStore::GetFormattedText] acpStart=" << acpStart << " acpEnd=" << acpEnd << "\n";
        return E_NOTIMPL;
    }

    /**
     * @brief ITextStoreACP::GetEmbedded
     *
     * 获取嵌入对象（如 OLE 对象），当前不实现。
     */
    STDMETHODIMP GetEmbedded(LONG acpPos, REFGUID rguidService, REFIID riid, IUnknown** ppunk) override
    {
    logger << "[CTextStore::GetEmbedded] acpPos=" << acpPos << "\n";
        return E_NOTIMPL;
    }

    /**
     * @brief ITextStoreACP::QueryInsertEmbedded
     *
     * 查询是否允许插入嵌入对象，这里一律返回不可插入。
     */
    STDMETHODIMP QueryInsertEmbedded(const GUID* pguidService, const FORMATETC* pFormatEtc, BOOL* pfInsertable) override
    {
    logger << "[CTextStore::QueryInsertEmbedded] pguidService=" << pguidService << " pFormatEtc=" << pFormatEtc << "\n";
        if (pfInsertable)
            *pfInsertable = FALSE;
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::InsertEmbedded
     *
     * 插入嵌入对象，不实现。
     */
    STDMETHODIMP InsertEmbedded(DWORD dwFlags, LONG acpStart, LONG acpEnd, IDataObject* pDataObject,
        TS_TEXTCHANGE* pChange) override
    {
    logger << "[CTextStore::InsertEmbedded] dwFlags=" << dwFlags << " acpStart=" << acpStart << " acpEnd=" << acpEnd << " pDataObject=" << pDataObject << "\n";
        return E_NOTIMPL;
    }

    /**
     * @brief ITextStoreACP::InsertTextAtSelection
     *
     * 在当前选区插入文本的简化实现。
     * @param dwFlags 标志
     * @param pchText 待插入文本缓冲
     * @param cch 文本字符数
     * @param pacpStart 返回实际插入的开始位置
     * @param pacpEnd 返回实际插入的结束位置
     * @param pChange 变更信息
     * @return S_OK 或 E_INVALIDARG
     */
    STDMETHODIMP InsertTextAtSelection(DWORD dwFlags, const WCHAR* pchText, ULONG cch,
        LONG* pacpStart, LONG* pacpEnd, TS_TEXTCHANGE* pChange) override
    {
    logger << "[CTextStore::InsertTextAtSelection] dwFlags=" << dwFlags << " pacpStart=" << (void*)pacpStart << " pacpEnd=" << (void*)pacpEnd << " cch=" << cch << "\n";
    // 参数验证
        if (!pchText || cch == 0)
        {
            logger << "[InsertTextAtSelection] dwFlags: " << dwFlags
                << ", pchText: " << ""
                << ", *pacpStart: " << *pacpStart
                << ", *pacpEnd: " << *pacpEnd
                << "\n";
            return E_INVALIDARG;
        }
        logger << "[InsertTextAtSelection] dwFlags: " << dwFlags
            << ", pchText: " << wstr2str_2ANSI(pchText)
            << ", *pacpStart: " << *pacpStart
            << ", *pacpEnd: " << *pacpEnd
            << "\n";
        LONG insertPos = pacpStart ? *pacpStart : m_selection.acpStart;
        LONG textLength = static_cast<LONG>(m_text.length());
        if (insertPos < 0 || insertPos > textLength)
            return TF_E_INVALIDPOS;

        // 如果不是只查询模式，执行实际插入
        if (!(dwFlags & TS_IAS_QUERYONLY))
        {
            //m_text.insert(insertPos, pchText, cch);
            m_compText.insert(insertPos, pchText, cch);
            m_compSel.acpStart = m_compSel.acpEnd = insertPos + cch;
        }
        if (pacpStart) *pacpStart = insertPos;
        if (pacpEnd) *pacpEnd = insertPos + cch;
        if (pChange)
        {
            pChange->acpStart = insertPos;
            pChange->acpOldEnd = insertPos;
            pChange->acpNewEnd = insertPos + cch;
        }
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::InsertEmbeddedAtSelection
     *
     * 在选区插入嵌入对象，不实现。
     */
    STDMETHODIMP InsertEmbeddedAtSelection(DWORD dwFlags, IDataObject* pDataObject,
        LONG* pacpStart, LONG* pacpEnd, TS_TEXTCHANGE* pChange) override
    {
    logger << "[CTextStore::InsertEmbeddedAtSelection] dwFlags=" << dwFlags << " pDataObject=" << pDataObject << "\n";
        return E_NOTIMPL;
    }

    /**
     * @brief ITextStoreACP::RequestSupportedAttrs
     *
     * 请求 TSF 支持的属性，这里简单返回成功（不执行实际缓存）。
     */
    STDMETHODIMP RequestSupportedAttrs(DWORD dwFlags, ULONG cFilterAttrs, const TS_ATTRID* paFilterAttrs) override
    {
    logger << "[CTextStore::RequestSupportedAttrs] dwFlags=" << dwFlags << " cFilterAttrs=" << cFilterAttrs << "\n";
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::RequestAttrsAtPosition
     *
     * 请求在某个位置的属性（如文本样式），这里不维护属性缓存，直接返回成功。
     */
    STDMETHODIMP RequestAttrsAtPosition(LONG acpPos, ULONG cFilterAttrs,
        const TS_ATTRID* paFilterAttrs, DWORD dwFlags) override
    {
    logger << "[CTextStore::RequestAttrsAtPosition] acpPos=" << acpPos << " cFilterAttrs=" << cFilterAttrs << " dwFlags=" << dwFlags << "\n";
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::RequestAttrsTransitioningAtPosition
     *
     * 请求属性从一个运行到下一个运行过渡的信息，当前未实现。
     */
    STDMETHODIMP RequestAttrsTransitioningAtPosition(
        /* [in] */ LONG acpPos,
        /* [in] */ ULONG cFilterAttrs,
        /* [unique][size_is][in] */ __RPC__in_ecount_full_opt(cFilterAttrs) const TS_ATTRID* paFilterAttrs,
        /* [in] */ DWORD dwFlags)
    {
    logger << "[CTextStore::RequestAttrsTransitioningAtPosition] acpPos=" << acpPos << " cFilterAttrs=" << cFilterAttrs << " dwFlags=" << dwFlags << "\n";
        return E_NOTIMPL;
    }

    /**
     * @brief ITextStoreACP::FindNextAttrTransition
     *
     * 查找从 acpStart 开始下一个属性过渡的位置，当前未实现。
     */
    STDMETHODIMP FindNextAttrTransition(LONG acpStart, LONG acpHalt, ULONG cFilterAttrs,
        const TS_ATTRID* paFilterAttrs, DWORD dwFlags,
        LONG* pacpNext, BOOL* pfFound, LONG* plFoundOffset) override
    {
    logger << "[CTextStore::FindNextAttrTransition] acpStart=" << acpStart << " acpHalt=" << acpHalt << "\n";
        return E_NOTIMPL;
    }

    /**
     * @brief ITextStoreACP::RetrieveRequestedAttrs
     *
     * 返回之前请求的属性值，这里不维护具体属性，直接返回 0 个结果。
     */
    STDMETHODIMP RetrieveRequestedAttrs(ULONG ulCount, TS_ATTRVAL* paAttrVals, ULONG* pcFetched) override
    {
    logger << "[CTextStore::RetrieveRequestedAttrs] ulCount=" << ulCount << "\n";
        if (pcFetched)
            *pcFetched = 0;
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::GetEndACP
     *
     * 返回文本末尾的 ACP 偏移（即文本长度）。
     * @param pacp 输出的 ACP 偏移
     * @return S_OK 或 E_INVALIDARG
     */
    STDMETHODIMP GetEndACP(LONG* pacp) override
    {
    logger << "[CTextStore::GetEndACP] pacp=" << pacp << "\n";
        if (!pacp)
            return E_INVALIDARG;

        *pacp = static_cast<LONG>(m_text.length());
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::GetActiveView
     *
     * 返回当前活动视图的 cookie（用于与视图相关的 API）。
     * @param pvcView 输出的视图 cookie
     * @return S_OK 或 E_INVALIDARG
     */
    STDMETHODIMP GetActiveView(TsViewCookie* pvcView) override
    {
    logger << "[CTextStore::GetActiveView] pvcView=" << pvcView << "\n";
        if (!pvcView)
            return E_INVALIDARG;

        *pvcView = m_viewCookie;
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::GetWnd
     *
     * 返回与视图关联的窗口句柄。
     * @param vcView 视图 cookie
     * @param phwnd 输出窗口句柄
     * @return S_OK 或 E_INVALIDARG
     */
    STDMETHODIMP GetWnd(TsViewCookie vcView, HWND* phwnd) override
    {
    logger << "[CTextStore::GetWnd] vcView=" << vcView << " phwnd=" << phwnd << "\n";
        if (!phwnd)
            return E_INVALIDARG;

        *phwnd = m_hwndEdit;
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::RequestLock
     *
     * TSF 请求对文本存储上锁（同步或异步）。该实现简单处理嵌套与回调通知。
     * @param dwLockFlags 锁标志
     * @param phrSession 返回会话结果（S_OK 或错误码）
     * @return S_OK 或 E_INVALIDARG
     */
    STDMETHODIMP RequestLock(DWORD dwLockFlags, HRESULT* phrSession) override
    {
    logger << "[CTextStore::RequestLock] dwLockFlags=" << dwLockFlags << " currentLockType=" << m_dwLockType << "\n";
        if (!phrSession)
            return E_INVALIDARG;

        if (m_dwLockType != 0)
        {
            // 已经有锁，检查是否可以嵌套
            if (dwLockFlags & TS_LF_SYNC)
            {
                *phrSession = TS_E_SYNCHRONOUS;
                return S_OK;
            }
        }

        *phrSession = S_OK;
        m_dwLockType = dwLockFlags;

        // 通知锁被授予
        if (m_pSink)
        {
            m_pSink->OnLockGranted(dwLockFlags); // 立即授予锁并处理文本请求
        }
        m_dwLockType = 0; // 重置锁的状态
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::GetACPFromPoint
     *
     * 根据屏幕/视图坐标返回对应的 ACP 偏移，当前未实现。
     */
    STDMETHODIMP GetACPFromPoint(TsViewCookie vcView, const POINT* pt, DWORD dwFlags, LONG* pacp) override
    {
    logger << "[CTextStore::GetACPFromPoint] vcView=" << vcView << " pt=" << (pt ? ("(" + std::to_string(pt->x) + "," + std::to_string(pt->y) + ")") : std::string("null")) << " dwFlags=" << dwFlags << "\n";
        if (!pt || !pacp)
            return E_INVALIDARG;

        // TSF 传入的 pt 通常为屏幕坐标，EM_CHARFROMPOS 需要客户端坐标。
        // 先将屏幕坐标转换为编辑控件的客户端坐标再发送消息。
        if (!IsWindow(m_hwndEdit))
            return E_FAIL;

        POINT ptClient = *pt;
        if (!ScreenToClient(m_hwndEdit, &ptClient))
            return E_FAIL;

        LPARAM lparam = MAKELPARAM(ptClient.x, ptClient.y);
        LRESULT lr = SendMessage(m_hwndEdit, EM_CHARFROMPOS, 0, lparam);
        // EM_CHARFROMPOS 返回的字符索引位于 LOWORD
        *pacp = static_cast<LONG>(LOWORD(lr));
        return S_OK;

        return E_FAIL;
    }


    class TestLogRectInfo {
        RECT* prc;
        FileLogger& logger;
    public:
        TestLogRectInfo(FileLogger& l, RECT* rc) : logger(l) {
            prc = rc;
        }

        ~TestLogRectInfo() {
            logger << "[ITextStoreACP::GetTextExt] Rect(x, y, width, height): "
                << prc->left << ", " << prc->top << ", "
                << prc->right - prc->left << ", " << prc->bottom - prc->top << "\n";
        }
    };
    /**
     * @brief ITextStoreACP::GetTextExt
     *
     * 返回指定文本范围在屏幕或窗口坐标系下的矩形区域，pfClipped 指示是否被裁剪。
     * 此实现简化为返回整个编辑控件的矩形。
     * @param vcView 视图 cookie
     * @param acpStart/acpEnd 文本范围
     * @param prc 输出矩形
     * @param pfClipped 输出是否被裁剪
     * @return S_OK 或 E_INVALIDARG / E_FAIL
     */
    STDMETHODIMP GetTextExt(TsViewCookie vcView, LONG acpStart, LONG acpEnd,
        RECT* prc, BOOL* pfClipped) override
    {
    logger << "[CTextStore::GetTextExt] vcView=" << vcView << " acpStart=" << acpStart << " acpEnd=" << acpEnd << " prc=" << prc << "\n";
        if (!prc || !pfClipped)
            return E_INVALIDARG;
        prc->left = 0;
        prc->right = 0;
        prc->top = 0;
        prc->bottom = 0;
        *pfClipped = FALSE;
        TestLogRectInfo tlri(logger, prc);
        if (!IsWindow(m_hwndEdit))
            return E_FAIL;

        // 初始化返回值为 FALSE
        *pfClipped = FALSE;

        // 如果请求的是整个文档或无效范围，返回整个编辑控件区域
        if (acpStart < 0 || acpEnd < acpStart)
        {
            if (GetClientRect(m_hwndEdit, prc))
            {
                MapWindowPoints(m_hwndEdit, nullptr, (LPPOINT)prc, 2);
                return S_OK;
            }
            return E_FAIL;
        }

        // 获取文本长度
        LONG textLength = static_cast<LONG>(GetWindowTextLength(m_hwndEdit));
        if (acpStart >= textLength)
        {
            // 位置超出文本范围，返回编辑控件右下角的一个小矩形
            if (GetClientRect(m_hwndEdit, prc))
            {
                prc->left = prc->right - 2;
                prc->top = prc->bottom - 16;
                MapWindowPoints(m_hwndEdit, nullptr, (LPPOINT)prc, 2);
                return S_OK;
            }
            return E_FAIL;
        }

        // 调整结束位置
        if (acpEnd > textLength)
            acpEnd = textLength;

        // 对于单行编辑控件，使用 EM_POSFROMCHAR
        DWORD style = GetWindowLong(m_hwndEdit, GWL_STYLE);
        BOOL isMultiLine = (style & ES_MULTILINE) != 0;

        if (!isMultiLine)
        {
            // 单行编辑控件 - 使用 EM_POSFROMCHAR
            LRESULT posStart = SendMessage(m_hwndEdit, EM_POSFROMCHAR, (WPARAM)acpStart, 0);
            LRESULT posEnd = SendMessage(m_hwndEdit, EM_POSFROMCHAR, (WPARAM)acpEnd, 0);

            if (posStart != -1 && posEnd != -1)
            {
                int xStart = GET_X_LPARAM(posStart);
                int yStart = GET_Y_LPARAM(posStart);
                int xEnd = GET_X_LPARAM(posEnd);
                int yEnd = GET_Y_LPARAM(posEnd);

                // 如果起始和结束在同一行
                if (yStart == yEnd)
                {
                    prc->left = xStart;
                    prc->top = yStart;
                    prc->right = xEnd;

                    // 获取字符高度
                    HDC hdc = GetDC(m_hwndEdit);
                    if (hdc)
                    {
                        HFONT hFont = (HFONT)SendMessage(m_hwndEdit, WM_GETFONT, 0, 0);
                        HFONT hOldFont = nullptr;
                        if (hFont)
                            hOldFont = (HFONT)SelectObject(hdc, hFont);

                        TEXTMETRIC tm;
                        if (GetTextMetrics(hdc, &tm))
                        {
                            prc->bottom = yStart + tm.tmHeight;
                        }
                        else
                        {
                            prc->bottom = yStart + 16; // 回退值
                        }

                        if (hOldFont)
                            SelectObject(hdc, hOldFont);
                        ReleaseDC(m_hwndEdit, hdc);
                    }
                    else
                    {
                        prc->bottom = yStart + 16;
                    }
                }
                else
                {
                    // 跨行文本，返回整个编辑控件区域
                    if (!GetClientRect(m_hwndEdit, prc))
                        return E_FAIL;
                }
            }
            else
            {
                // EM_POSFROMCHAR 失败，返回整个编辑控件区域
                if (!GetClientRect(m_hwndEdit, prc))
                    return E_FAIL;
            }
        }
        else
        {
            // 多行编辑控件 - 使用更复杂的位置计算
            // 简化处理：返回基于插入符位置的矩形
            LRESULT pos = SendMessage(m_hwndEdit, EM_POSFROMCHAR, (WPARAM)acpStart, 0);
            if (pos != -1)
            {
                int x = GET_X_LPARAM(pos);
                int y = GET_Y_LPARAM(pos);

                prc->left = x;
                prc->top = y;
                prc->right = x + 2; // 小宽度表示插入点

                // 获取行高
                HDC hdc = GetDC(m_hwndEdit);
                if (hdc)
                {
                    HFONT hFont = (HFONT)SendMessage(m_hwndEdit, WM_GETFONT, 0, 0);
                    HFONT hOldFont = nullptr;
                    if (hFont)
                        hOldFont = (HFONT)SelectObject(hdc, hFont);

                    TEXTMETRIC tm;
                    if (GetTextMetrics(hdc, &tm))
                    {
                        prc->bottom = y + tm.tmHeight;
                    }
                    else
                    {
                        prc->bottom = y + 16;
                    }

                    if (hOldFont)
                        SelectObject(hdc, hOldFont);
                    ReleaseDC(m_hwndEdit, hdc);
                }
                else
                {
                    prc->bottom = y + 16;
                }
            }
            else
            {
                // 失败时返回整个编辑控件区域
                if (!GetClientRect(m_hwndEdit, prc))
                    return E_FAIL;
            }
        }

        // 转换为屏幕坐标
        MapWindowPoints(m_hwndEdit, nullptr, (LPPOINT)prc, 2);
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::GetScreenExt
     *
     * 返回关联窗口在屏幕坐标下的外接矩形。
     * @param vcView 视图 cookie
     * @param prc 输出矩形
     * @return S_OK 或 E_INVALIDARG / E_FAIL
     */
    STDMETHODIMP GetScreenExt(TsViewCookie vcView, RECT* prc) override
    {
    logger << "[CTextStore::GetScreenExt] vcView=" << vcView << " prc=" << prc << "\n";
        if (!prc)
            return E_INVALIDARG;

        // 返回屏幕区域
        RECT rect = { 0 };
        if (GetWindowRect(m_hwndEdit, &rect))
        {
            *prc = rect;
            return S_OK;
        }

        return E_FAIL;
    }

private:
    TS_SELECTION_ACP m_compSel;
    BOOL m_compActive = FALSE;
    BOOL m_isComp = FALSE;
public:
    // --------------- ITfContextOwnerCompositionSink ---------------
    STDMETHODIMP OnStartComposition(ITfCompositionView* pView, BOOL* pfOk)
    {
        logger << "[ITfContextOwnerCompositionSink] OnStartComposition\n";
        *pfOk = TRUE;                     // 同意开始
        m_isComp = TRUE;
        return S_OK;
    }

    STDMETHODIMP OnUpdateComposition(ITfCompositionView* pView,
        ITfRange* pRangeNew)
    {
        logger << "[ITfContextOwnerCompositionSink] OnUpdateComposition\n";
        m_isComp = TRUE;
        ITfRange* range = nullptr;
        if (SUCCEEDED(pView->GetRange(&range)))
        {
            
        }
        return S_OK;
    }

    STDMETHODIMP OnEndComposition(ITfCompositionView* pView)
    {
        logger << "[ITfContextOwnerCompositionSink] OnEndComposition\n";
        
        WCHAR* t = new WCHAR[GetWindowTextLength(m_hwndEdit) + 1];
        GetWindowText(m_hwndEdit, t, GetWindowTextLength(m_hwndEdit) + 1);
        m_text = t;
        delete[] t;
        if (isCompFinished)
            m_text.replace(m_selection.acpStart, m_selection.acpEnd - m_selection.acpStart, m_compText);
        SetWindowText(m_hwndEdit, m_text.c_str());
        SendMessage(m_hwndEdit, EM_SETSEL, m_selection.acpStart + m_compText.size(), m_selection.acpStart + m_compText.size());
        SendMessage(m_hwndEdit, EM_SCROLLCARET, 0, 0);
        m_compText.clear();
        // 组合结束，把标记清掉，但字符保留
        m_compSel.acpStart = -1;
        m_compSel.acpEnd = -1;
        m_compActive = FALSE;
        m_isComp = FALSE;
        isCompFinished = FALSE;
        return S_OK;
    }

    // -----------ITfCompositionSink--------------
    STDMETHODIMP OnCompositionTerminated(TfEditCookie ecWrite, __RPC__in_opt ITfComposition* pComposition)
    {
        logger << "[ITfCompositionSink] OnCompositionTerminated\n";
        //WCHAR* t = new WCHAR[GetWindowTextLength(m_hwndEdit) + 1];
        //GetWindowText(m_hwndEdit, t, GetWindowTextLength(m_hwndEdit) + 1);
        //m_text = t;
        //delete[] t;
        //if (isCompFinished)
        //    m_text.replace(m_selection.acpStart, m_selection.acpEnd - m_selection.acpStart, m_compText);
        //SetWindowText(m_hwndEdit, m_text.c_str());
        //SendMessage(m_hwndEdit, EM_SETSEL, m_selection.acpStart + m_compText.size(), m_selection.acpStart + m_compText.size());
        //SendMessage(m_hwndEdit, EM_SCROLLCARET, 0, 0);
        //m_compText.clear();
        //m_isComp = FALSE;
        //isCompFinished = FALSE;
        return S_OK;
    }

private:


private:

};

// 暂时无用
class CCompositionSink : public ITfCompositionSink
{
private:
    ULONG m_cRef = 0;
public:
    STDMETHODIMP OnCompositionTerminated(TfEditCookie ecWrite, ITfComposition* pComposition) override
    {
        // 组合结束了
        ITfRange* pRange = NULL;
        if (SUCCEEDED(pComposition->GetRange(&pRange)))
        {
            // 处理组合结束
            pRange->SetText(ecWrite, 0, L"", 0);
            pRange->Release();
        }
        return S_OK;
    }

    /**
 * @brief IUnknown::QueryInterface
 *
 * 查询并返回支持的接口指针（IUnknown 或 ITextStoreACP）。
 * @param riid 请求的接口 ID
 * @param ppvObj 输出接口指针
 * @return S_OK 或 E_NOINTERFACE
 */
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override
    {
        if (riid == IID_IUnknown || riid == IID_ITfCompositionSink)
        {
            *ppvObj = this;
            AddRef();
            return S_OK;
        }
        *ppvObj = nullptr;
        return E_NOINTERFACE;
    }

    /**
     * @brief IUnknown::AddRef
     *
     * 增加引用计数并返回新的计数值。
     * @return 新的引用计数
     */
    STDMETHODIMP_(ULONG) AddRef() override
    {
        return InterlockedIncrement(&m_cRef);
    }

    /**
     * @brief IUnknown::Release
     *
     * 减少引用计数，计数为 0 时销毁对象，返回剩余计数。
     * @return 剩余引用计数
     */
    STDMETHODIMP_(ULONG) Release() override
    {
        LONG cRef = InterlockedDecrement(&m_cRef);
        if (cRef == 0)
        {
            delete this;
        }
        return cRef;
    }

};

// 实现上下文组合接口
class CContextComposition : public ITfContextComposition
{
private:
    ULONG m_cRef = 0;
    ITfContext* m_pContext;
public:
    CContextComposition(ITfContext* pContext) : m_pContext(pContext)
    {
        if (m_pContext)
            m_pContext->AddRef();
    }
    virtual ~CContextComposition()
    {
        if (m_pContext)
            m_pContext->Release();
    }

    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override
    {
        if (riid == IID_IUnknown || riid == IID_ITfContextComposition)
        {
            *ppvObj = static_cast<ITfContextComposition*>(this);
            AddRef();
            return S_OK;
        }
        *ppvObj = nullptr;
        return E_NOINTERFACE;
    }

    STDMETHODIMP_(ULONG) AddRef() override
    {
        return InterlockedIncrement(&m_cRef);
    }

    STDMETHODIMP_(ULONG) Release() override
    {
        LONG cRef = InterlockedDecrement(&m_cRef);
        if (cRef == 0)
        {
            delete this;
        }
        return cRef;
    }

};