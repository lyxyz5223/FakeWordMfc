#include <windows.h>
#include <msctf.h>
#include <string>
#include <vector>

class CTextStore;

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

public:
    CTSFManager() :
        m_pThreadMgr(nullptr),
        m_pDocumentMgr(nullptr),
        m_pContext(nullptr),
        m_pTextStore(nullptr),
        m_clientId(TF_CLIENTID_NULL),
        m_hwndEdit(nullptr)
    {
    }

    ~CTSFManager()
    {
        Uninitialize();
    }

    BOOL Initialize(HWND hwndEdit);

    void Uninitialize();

    // 处理文本输入
    void HandleTextInput(const std::wstring& text);

private:
    BOOL ShouldBlockInput(const std::wstring& text);

    void AddTextToEditControl(const std::wstring& text);
};

class CTextStore : public ITextStoreACP
{
private:
    LONG m_cRef;
    CTSFManager* m_pManager;
    HWND m_hwndEdit;
    ITfContext* m_pContext;
    TfClientId m_clientId;
    DWORD m_dwLockType;
    ITextStoreACPSink* m_pSink;

    // 文本内容
    std::wstring m_text;
    TS_SELECTION_ACP m_selection;
    TsViewCookie m_viewCookie;

public:
    CTextStore(CTSFManager* pManager, HWND hwndEdit) :
        m_cRef(1),
        m_pManager(pManager),
        m_hwndEdit(hwndEdit),
        m_pContext(nullptr),
        m_clientId(TF_CLIENTID_NULL),
        m_dwLockType(0),
        m_pSink(nullptr),
        m_viewCookie(0)
    {
        m_selection.acpStart = 0;
        m_selection.acpEnd = 0;
        m_selection.style.ase = TS_AE_NONE;
        m_selection.style.fInterimChar = FALSE;
    }

    virtual ~CTextStore()
    {
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

    // 初始化
    void Initialize(ITfContext* pContext, TfClientId clientId)
    {
        m_pContext = pContext;
        m_pContext->AddRef();
        m_clientId = clientId;
    }

    // IUnknown 方法
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override
    {
        if (riid == IID_IUnknown || riid == IID_ITextStoreACP)
        {
            *ppvObj = static_cast<ITextStoreACP*>(this);
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

    // ITextStoreACP 方法
    STDMETHODIMP AdviseSink(REFIID riid, IUnknown* punk, DWORD dwMask) override
    {
        if (riid != IID_ITextStoreACPSink)
            return E_INVALIDARG;

        if (m_pSink)
        {
            m_pSink->Release();
            m_pSink = nullptr;
        }

        HRESULT hr = punk->QueryInterface(IID_ITextStoreACPSink, (void**)&m_pSink);
        if (SUCCEEDED(hr) && m_pSink)
        {
            m_pSink->AddRef();
        }

        return hr;
    }

    STDMETHODIMP UnadviseSink(IUnknown* punk) override
    {
        if (m_pSink)
        {
            m_pSink->Release();
            m_pSink = nullptr;
        }
        return S_OK;
    }

    STDMETHODIMP GetStatus(TS_STATUS* pdcs) override
    {
        if (!pdcs)
            return E_INVALIDARG;

        pdcs->dwDynamicFlags = 0;
        pdcs->dwStaticFlags = TS_SS_TRANSITORY | TS_SS_NOHIDDENTEXT;
        return S_OK;
    }

    STDMETHODIMP QueryInsert(LONG acpTestStart, LONG acpTestEnd, ULONG cch,
        LONG* pacpResultStart, LONG* pacpResultEnd) override
    {
        if (!pacpResultStart || !pacpResultEnd)
            return E_INVALIDARG;

        *pacpResultStart = acpTestStart;
        *pacpResultEnd = acpTestEnd;
        return S_OK;
    }

    STDMETHODIMP GetSelection(ULONG ulIndex, ULONG ulCount,
        TS_SELECTION_ACP* pSelection, ULONG* pcFetched) override
    {
        if (!pSelection || !pcFetched)
            return E_INVALIDARG;

        if (ulIndex != 0)
            return TS_E_NOSELECTION;

        *pcFetched = 1;
        pSelection[0] = m_selection;
        return S_OK;
    }

    STDMETHODIMP SetSelection(ULONG ulCount, const TS_SELECTION_ACP* pSelection) override
    {
        if (ulCount != 1 || !pSelection)
            return E_INVALIDARG;

        m_selection = pSelection[0];
        return S_OK;
    }

    STDMETHODIMP GetText(LONG acpStart, LONG acpEnd, WCHAR* pchPlain, ULONG cchPlainReq,
        ULONG* pcchPlainOut, TS_RUNINFO* prgRunInfo, ULONG ulRunInfoReq,
        ULONG* pulRunInfoOut, LONG* pacpNext) override
    {
        if (!pcchPlainOut)
            return E_INVALIDARG;

        // 计算文本范围
        LONG acpEndAdjusted = acpEnd;
        if (acpEnd == -1)
            acpEndAdjusted = static_cast<LONG>(m_text.length());

        if (acpStart < 0 || acpStart > acpEndAdjusted || acpEndAdjusted > static_cast<LONG>(m_text.length()))
            return TF_E_INVALIDPOS;

        ULONG cchToCopy = acpEndAdjusted - acpStart;
        if (cchToCopy > cchPlainReq)
            cchToCopy = cchPlainReq;

        // 复制文本
        if (pchPlain && cchToCopy > 0)
        {
            memcpy(pchPlain, m_text.c_str() + acpStart, cchToCopy * sizeof(WCHAR));
        }

        *pcchPlainOut = cchToCopy;

        if (pulRunInfoOut)
            *pulRunInfoOut = 0;

        if (pacpNext)
            *pacpNext = acpStart + cchToCopy;

        return S_OK;
    }

    STDMETHODIMP SetText(DWORD dwFlags, LONG acpStart, LONG acpEnd,
        const WCHAR* pchText, ULONG cch, TS_TEXTCHANGE* pChange) override
    {
        // 处理文本设置
        if (pchText && cch > 0)
        {
            std::wstring newText(pchText, cch);
            if (m_pManager)
            {
                m_pManager->HandleTextInput(newText);
            }
        }

        if (pChange)
        {
            pChange->acpStart = acpStart;
            pChange->acpOldEnd = acpEnd;
            pChange->acpNewEnd = acpStart + cch;
        }

        return S_OK;
    }

    STDMETHODIMP GetFormattedText(LONG acpStart, LONG acpEnd, IDataObject** ppDataObject) override
    {
        return E_NOTIMPL;
    }

    STDMETHODIMP GetEmbedded(LONG acpPos, REFGUID rguidService, REFIID riid, IUnknown** ppunk) override
    {
        return E_NOTIMPL;
    }

    STDMETHODIMP QueryInsertEmbedded(const GUID* pguidService, const FORMATETC* pFormatEtc, BOOL* pfInsertable) override
    {
        if (pfInsertable)
            *pfInsertable = FALSE;
        return S_OK;
    }

    STDMETHODIMP InsertEmbedded(DWORD dwFlags, LONG acpStart, LONG acpEnd, IDataObject* pDataObject,
        TS_TEXTCHANGE* pChange) override
    {
        return E_NOTIMPL;
    }

    STDMETHODIMP InsertTextAtSelection(DWORD dwFlags, const WCHAR* pchText, ULONG cch,
        LONG* pacpStart, LONG* pacpEnd, TS_TEXTCHANGE* pChange) override
    {
        if (!pchText || cch == 0)
            return E_INVALIDARG;

        // 处理文本插入
        //std::wstring newText(pchText, cch);
        //if (m_pManager)
        //{
        //    m_pManager->HandleTextInput(newText);
        //}

        if (pacpStart)
            *pacpStart = 0;
        if (pacpEnd)
            *pacpEnd = static_cast<LONG>(cch);

        return S_OK;
    }

    STDMETHODIMP InsertEmbeddedAtSelection(DWORD dwFlags, IDataObject* pDataObject,
        LONG* pacpStart, LONG* pacpEnd, TS_TEXTCHANGE* pChange) override
    {
        return E_NOTIMPL;
    }

    STDMETHODIMP RequestSupportedAttrs(DWORD dwFlags, ULONG cFilterAttrs, const TS_ATTRID* paFilterAttrs) override
    {
        return S_OK;
    }

    STDMETHODIMP RequestAttrsAtPosition(LONG acpPos, ULONG cFilterAttrs,
        const TS_ATTRID* paFilterAttrs, DWORD dwFlags) override
    {
        return S_OK;
    }

    STDMETHODIMP RequestAttrsTransitioningAtPosition(
        /* [in] */ LONG acpPos,
        /* [in] */ ULONG cFilterAttrs,
        /* [unique][size_is][in] */ __RPC__in_ecount_full_opt(cFilterAttrs) const TS_ATTRID* paFilterAttrs,
        /* [in] */ DWORD dwFlags)
    {
        return E_NOTIMPL;
    }

    STDMETHODIMP FindNextAttrTransition(LONG acpStart, LONG acpHalt, ULONG cFilterAttrs,
        const TS_ATTRID* paFilterAttrs, DWORD dwFlags,
        LONG* pacpNext, BOOL* pfFound, LONG* plFoundOffset) override
    {
        return E_NOTIMPL;
    }

    STDMETHODIMP RetrieveRequestedAttrs(ULONG ulCount, TS_ATTRVAL* paAttrVals, ULONG* pcFetched) override
    {
        if (pcFetched)
            *pcFetched = 0;
        return S_OK;
    }

    STDMETHODIMP GetEndACP(LONG* pacp) override
    {
        if (!pacp)
            return E_INVALIDARG;

        *pacp = static_cast<LONG>(m_text.length());
        return S_OK;
    }

    STDMETHODIMP GetActiveView(TsViewCookie* pvcView) override
    {
        if (!pvcView)
            return E_INVALIDARG;

        *pvcView = m_viewCookie;
        return S_OK;
    }

    STDMETHODIMP GetWnd(TsViewCookie vcView, HWND* phwnd) override
    {
        if (!phwnd)
            return E_INVALIDARG;

        *phwnd = m_hwndEdit;
        return S_OK;
    }

    STDMETHODIMP RequestLock(DWORD dwLockFlags, HRESULT* phrSession) override
    {
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

        m_dwLockType = dwLockFlags;
        *phrSession = S_OK;

        // 通知锁被授予
        if (m_pSink)
        {
            m_pSink->OnLockGranted(dwLockFlags);
        }

        m_dwLockType = 0;
        return S_OK;
    }

    // 以下是被遗漏的纯虚函数
    STDMETHODIMP GetACPFromPoint(TsViewCookie vcView, const POINT* pt, DWORD dwFlags, LONG* pacp) override
    {
        return E_NOTIMPL;
    }

    STDMETHODIMP GetTextExt(TsViewCookie vcView, LONG acpStart, LONG acpEnd,
        RECT* prc, BOOL* pfClipped) override
    {
        if (!prc || !pfClipped)
            return E_INVALIDARG;

        // 简单实现：返回整个编辑框的矩形
        if (GetClientRect(m_hwndEdit, prc))
        {
            MapWindowPoints(m_hwndEdit, nullptr, (LPPOINT)prc, 2);
            *pfClipped = FALSE;
            return S_OK;
        }

        return E_FAIL;
    }

    STDMETHODIMP GetScreenExt(TsViewCookie vcView, RECT* prc) override
    {
        if (!prc)
            return E_INVALIDARG;

        // 返回屏幕区域
        if (GetWindowRect(m_hwndEdit, prc))
        {
            return S_OK;
        }

        return E_FAIL;
    }

    // 更新文本内容
    void SetText(const std::wstring& text)
    {
        m_text = text;
    }

    // 获取文本内容
    const std::wstring& GetText() const
    {
        return m_text;
    }
};
