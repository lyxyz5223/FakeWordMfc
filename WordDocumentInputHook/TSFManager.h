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
 * ���� Text Services Framework (TSF) ���������ں���༭�ؼ��Ľ�����
 * �����ʼ��/����ʼ�� TSF���������� TSF ���ı����룬��ά����Ҫ�� TSF ��������
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
     * @brief ���캯��
     *
     * ��ʼ���ڲ�ָ��Ϊ nullptr������ clientId �ʹ��ھ����ΪĬ��ֵ��
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
     * @brief ��������
     *
     * �ڶ�������ʱȷ���ͷŲ�����ʼ�� TSF �����Դ��
     */
    ~CTSFManager()
    {
        Uninitialize();
    }

    /**
     * @brief ��ʼ�� TSF ��������������ָ���ı༭����
     *
     * @param hwndEdit Ҫ�� TSF �����ı༭�ؼ����ھ��
     * @return TRUE ��ʼ���ɹ���FALSE ��ʼ��ʧ��
     */
    BOOL Initialize(HWND hwndEdit);

    /**
     * @brief ����ʼ�� TSF ������
     *
     * �ͷ����� TSF �ӿں��ڲ���Դ��
     */
    void Uninitialize();

private:
};

/**
 * @brief CTextStore
 *
 * ʵ�� ITextStoreACP �ӿڣ������� TSF ����༭�������ı���ѡ����Ϣ��
 * ���������� TSF ������/�������󡣽��󲿷ֲ���ί�и� CTSFManager ����� UI ��ĸ��¡�
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
    ITextStoreACPSink* m_pSink = nullptr; // �洢��

    // �ı�����
    std::wstring m_text;
    TS_SELECTION_ACP m_selection;
    TsViewCookie m_viewCookie;

    // Composition
    DWORD m_CompSinkCookie = 0;
    std::wstring m_compText;
    BOOL isCompFinished = FALSE;
public:
    /**
     * @brief ���캯��
     *
     * @param pManager ָ���ϲ� CTSFManager ��ָ��
     * @param hwndEdit �����ı༭���ھ��
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
     * @brief ��������
     *
     * �ͷŶ� ITextStoreACPSink �� ITfContext ������
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
     * @brief ��ʼ���ı��洢
     *
     * ���沢 AddRef �ṩ�� ITfContext��ͬʱ��¼ clientId��
     * @param pContext TSF ������
     * @param clientId ע��Ŀͻ��� ID
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
                &m_CompSinkCookie);   // DWORD ��Ա������ʱ Unadvise
            if (FAILED(hr))
            {
                _com_error err(hr);
                logger << L"[CTextStore::Initialize] pSource->AdviseSink IID_ITfContextOwnerCompositionSink error: " << err.ErrorMessage() << L", code: " << hr;
            }
            //hr = pSource->AdviseSink(IID_ITfCompositionSink,
            //    static_cast<ITfCompositionSink*>(this),
            //    &m_CompSinkCookie);   // DWORD ��Ա������ʱ Unadvise
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
     * ��ѯ������֧�ֵĽӿ�ָ�루IUnknown �� ITextStoreACP����
     * @param riid ����Ľӿ� ID
     * @param ppvObj ����ӿ�ָ��
     * @return S_OK �� E_NOINTERFACE
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
     * �������ü����������µļ���ֵ��
     * @return �µ����ü���
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
     * �������ü���������Ϊ 0 ʱ���ٶ��󣬷���ʣ�������
     * @return ʣ�����ü���
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
     * ע����滻 ITextStoreACPSink �ص���TSF ��ͨ���ûص�֪ͨ�������¼���
     * @param riid ����Ϊ IID_ITextStoreACPSink
     * @param punk �ṩ�� sink ����
     * @param dwMask ��ߣ�δʹ�ã�
     * @return S_OK �������
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
     * ȡ��ע����ǰע��� ITextStoreACPSink��
     * @param punk ��ѡ��ͨ�����ԣ�ֱ���ͷű���� m_pSink ���ɡ�
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
     * �����ı��洢��״̬��Ϣ�����Ƿ���˲̬�ı����Ƿ��������ı��ȣ���
     * @param pdcs ����� TS_STATUS �ṹ
     * @return S_OK �� E_INVALIDARG
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
     * ����ѯ���ڸ���λ�ò���ָ�������ı��Ƿ���У�����ʵ�ʽ�Ҫ�����λ�÷�Χ��
     * @param acpTestStart/acpTestEnd ���ԵĲ��뷶Χ
     * @param cch �ַ�����
     * @param pacpResultStart/pacpResultEnd ���ʵ��λ��
     * @return S_OK �� E_INVALIDARG
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
     * ���ص�ǰ��ѡ����Ϣ�����ַ�ƫ�� ACP ��ʾ����
     * @param ulIndex/ulCount ��������������
     * @param pSelection �������
     * @param pcFetched ʵ�ʷ�������
     * @return S_OK �� E_INVALIDARG / TS_E_NOSELECTION
     */
    STDMETHODIMP GetSelection(ULONG ulIndex, ULONG ulCount,
        TS_SELECTION_ACP* pSelection, ULONG* pcFetched) override
    {
    logger << "[CTextStore::GetSelection] ulIndex=" << ulIndex << " ulCount=" << ulCount << "\n";
        // ������֤
        if (!pSelection || !pcFetched)
            return E_INVALIDARG;

        *pcFetched = 0;

        // ���������֧��Ĭ��ѡ��(0xFFFFFFFF)���һ��ѡ��(0)
        if (ulIndex != 0 && ulIndex != TS_DEFAULT_SELECTION)
            return TS_E_NOSELECTION;

        if (ulCount == 0)
            return E_INVALIDARG;

        // ��ȡ��ǰ�༭�ؼ���ѡ��״̬
        DWORD dwStart = 0, dwEnd = 0;

        if (IsWindow(m_hwndEdit))
        {
            // �ӱ༭�ؼ���ȡʵ��ѡ��
            SendMessage(m_hwndEdit, EM_GETSEL, (WPARAM)&dwStart, (LPARAM)&dwEnd);

            // �����ڲ�ѡ��״̬
            m_selection.acpStart = static_cast<LONG>(dwStart);
            m_selection.acpEnd = static_cast<LONG>(dwEnd);
        }

        // ��֤ѡ��Χ
        //LONG textLength = static_cast<LONG>(m_text.length());
        //if (m_selection.acpStart < 0)
        //    m_selection.acpStart = 0;
        //if (m_selection.acpStart > textLength)
        //    m_selection.acpStart = textLength;
        //if (m_selection.acpEnd < m_selection.acpStart)
        //    m_selection.acpEnd = m_selection.acpStart;
        //if (m_selection.acpEnd > textLength)
        //    m_selection.acpEnd = textLength;

        // ����Ĭ��ѡ������ȷ������һ����Ч�Ĳ����
        if (ulIndex == TS_DEFAULT_SELECTION)
        {
            // ͨ��Ĭ��ѡ����ǵ�ǰ���λ�ã�����㣩
            // ȷ��ѡ����һ������㣨��ʼ=������
            //m_selection.acpEnd = m_selection.acpStart;
        }

        // ������
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
     * ���õ�ǰѡ����Ϣ��ͨ���� TSF ���ƣ���
     * @param ulCount ѡ������������ֻ֧�� 1��
     * @param pSelection ����ѡ��
     * @return S_OK �� E_INVALIDARG / E_NOTIMPL
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
     * ���ڲ��ı���������ȡָ����Χ���ı���ACP ƫ�ƣ���
     * @param acpStart/acpEnd ��ȡ��Χ��acpEnd Ϊ -1 ��ʾֱ���ı�ĩβ��
     * @param pchPlain �������
     * @param cchPlainReq �����С
     * @param pcchPlainOut ʵ��д���ַ���
     * @param prgRunInfo �ı�������Ϣ����ѡ��
     * @param ulRunInfoReq ������Ϣ�����С
     * @param pulRunInfoOut ������Ϣʵ�ʷ��ش�С����ѡ��
     * @param pacpNext ��һ���ɶ�ȡλ�ã���ѡ��
     * @return S_OK �� TF_E_INVALIDPOS / E_INVALIDARG
     */
    STDMETHODIMP GetText(LONG acpStart, LONG acpEnd, WCHAR* pchPlain, ULONG cchPlainReq,
        ULONG* pcchPlainOut, TS_RUNINFO* prgRunInfo, ULONG ulRunInfoReq,
        ULONG* pulRunInfoOut, LONG* pacpNext) override
    {
    logger << "[CTextStore::GetText] acpStart=" << acpStart << " acpEnd=" << acpEnd << " cchPlainReq=" << cchPlainReq << "\n";
        // ��ʼ������
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
        // �����ı���Χ
        LONG acpEndAdjusted = acpEnd;
        if (acpEnd == -1)
            acpEndAdjusted = static_cast<LONG>(m_compText.length());

        if (acpStart < 0 || acpStart > acpEndAdjusted || acpEndAdjusted > static_cast<LONG>(m_compText.length()))
            return TF_E_INVALIDPOS;

        ULONG cchToCopy = acpEndAdjusted - acpStart;
        if (cchToCopy > cchPlainReq)
            cchToCopy = cchPlainReq;

        // �����ı�
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
     * TSF ������ָ����Χ�����ı�ʱ���á���ʵ�ֻὫ�ı�֪ͨ���ϲ���������д���
     * @param dwFlags ��־
     * @param acpStart/acpEnd Ŀ�귶Χ
     * @param pchText ���ı�����
     * @param cch �ı��ַ���
     * @param pChange ����ı����Ϣ����ѡ��
     * @return S_OK
     */
    STDMETHODIMP SetText(DWORD dwFlags, LONG acpStart, LONG acpEnd,
        const WCHAR* pchText, ULONG cch, TS_TEXTCHANGE* pChange) override
    {
        logger << "[SetText] dwFlags: " << dwFlags << "\n";
        logger << "[SetText] pchText: " << (pchText ? wstr2str_2ANSI(pchText) : "") << "\n";
        logger << "[SetText] Current Selection: " << m_selection.acpStart << " -> " << m_selection.acpEnd << "\n";
        logger << "[SetText] Current Selection Param: " << acpStart << " -> " << acpEnd << "\n";
        // �����ı�����
        if (pchText && cch > 0)
        {
            // �����ı�
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
     * �ɷ���ָ����Χ�ڵĸ��ı���ʾ�����������/�Ϸ��ã�����ǰ��ʵ�֡�
     */
    STDMETHODIMP GetFormattedText(LONG acpStart, LONG acpEnd, IDataObject** ppDataObject) override
    {
    logger << "[CTextStore::GetFormattedText] acpStart=" << acpStart << " acpEnd=" << acpEnd << "\n";
        return E_NOTIMPL;
    }

    /**
     * @brief ITextStoreACP::GetEmbedded
     *
     * ��ȡǶ������� OLE ���󣩣���ǰ��ʵ�֡�
     */
    STDMETHODIMP GetEmbedded(LONG acpPos, REFGUID rguidService, REFIID riid, IUnknown** ppunk) override
    {
    logger << "[CTextStore::GetEmbedded] acpPos=" << acpPos << "\n";
        return E_NOTIMPL;
    }

    /**
     * @brief ITextStoreACP::QueryInsertEmbedded
     *
     * ��ѯ�Ƿ��������Ƕ���������һ�ɷ��ز��ɲ��롣
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
     * ����Ƕ����󣬲�ʵ�֡�
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
     * �ڵ�ǰѡ�������ı��ļ�ʵ�֡�
     * @param dwFlags ��־
     * @param pchText �������ı�����
     * @param cch �ı��ַ���
     * @param pacpStart ����ʵ�ʲ���Ŀ�ʼλ��
     * @param pacpEnd ����ʵ�ʲ���Ľ���λ��
     * @param pChange �����Ϣ
     * @return S_OK �� E_INVALIDARG
     */
    STDMETHODIMP InsertTextAtSelection(DWORD dwFlags, const WCHAR* pchText, ULONG cch,
        LONG* pacpStart, LONG* pacpEnd, TS_TEXTCHANGE* pChange) override
    {
    logger << "[CTextStore::InsertTextAtSelection] dwFlags=" << dwFlags << " pacpStart=" << (void*)pacpStart << " pacpEnd=" << (void*)pacpEnd << " cch=" << cch << "\n";
    // ������֤
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

        // �������ֻ��ѯģʽ��ִ��ʵ�ʲ���
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
     * ��ѡ������Ƕ����󣬲�ʵ�֡�
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
     * ���� TSF ֧�ֵ����ԣ�����򵥷��سɹ�����ִ��ʵ�ʻ��棩��
     */
    STDMETHODIMP RequestSupportedAttrs(DWORD dwFlags, ULONG cFilterAttrs, const TS_ATTRID* paFilterAttrs) override
    {
    logger << "[CTextStore::RequestSupportedAttrs] dwFlags=" << dwFlags << " cFilterAttrs=" << cFilterAttrs << "\n";
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::RequestAttrsAtPosition
     *
     * ������ĳ��λ�õ����ԣ����ı���ʽ�������ﲻά�����Ի��棬ֱ�ӷ��سɹ���
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
     * �������Դ�һ�����е���һ�����й��ɵ���Ϣ����ǰδʵ�֡�
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
     * ���Ҵ� acpStart ��ʼ��һ�����Թ��ɵ�λ�ã���ǰδʵ�֡�
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
     * ����֮ǰ���������ֵ�����ﲻά���������ԣ�ֱ�ӷ��� 0 �������
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
     * �����ı�ĩβ�� ACP ƫ�ƣ����ı����ȣ���
     * @param pacp ����� ACP ƫ��
     * @return S_OK �� E_INVALIDARG
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
     * ���ص�ǰ���ͼ�� cookie����������ͼ��ص� API����
     * @param pvcView �������ͼ cookie
     * @return S_OK �� E_INVALIDARG
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
     * ��������ͼ�����Ĵ��ھ����
     * @param vcView ��ͼ cookie
     * @param phwnd ������ھ��
     * @return S_OK �� E_INVALIDARG
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
     * TSF ������ı��洢������ͬ�����첽������ʵ�ּ򵥴���Ƕ����ص�֪ͨ��
     * @param dwLockFlags ����־
     * @param phrSession ���ػỰ�����S_OK ������룩
     * @return S_OK �� E_INVALIDARG
     */
    STDMETHODIMP RequestLock(DWORD dwLockFlags, HRESULT* phrSession) override
    {
    logger << "[CTextStore::RequestLock] dwLockFlags=" << dwLockFlags << " currentLockType=" << m_dwLockType << "\n";
        if (!phrSession)
            return E_INVALIDARG;

        if (m_dwLockType != 0)
        {
            // �Ѿ�����������Ƿ����Ƕ��
            if (dwLockFlags & TS_LF_SYNC)
            {
                *phrSession = TS_E_SYNCHRONOUS;
                return S_OK;
            }
        }

        *phrSession = S_OK;
        m_dwLockType = dwLockFlags;

        // ֪ͨ��������
        if (m_pSink)
        {
            m_pSink->OnLockGranted(dwLockFlags); // �����������������ı�����
        }
        m_dwLockType = 0; // ��������״̬
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::GetACPFromPoint
     *
     * ������Ļ/��ͼ���귵�ض�Ӧ�� ACP ƫ�ƣ���ǰδʵ�֡�
     */
    STDMETHODIMP GetACPFromPoint(TsViewCookie vcView, const POINT* pt, DWORD dwFlags, LONG* pacp) override
    {
    logger << "[CTextStore::GetACPFromPoint] vcView=" << vcView << " pt=" << (pt ? ("(" + std::to_string(pt->x) + "," + std::to_string(pt->y) + ")") : std::string("null")) << " dwFlags=" << dwFlags << "\n";
        if (!pt || !pacp)
            return E_INVALIDARG;

        // TSF ����� pt ͨ��Ϊ��Ļ���꣬EM_CHARFROMPOS ��Ҫ�ͻ������ꡣ
        // �Ƚ���Ļ����ת��Ϊ�༭�ؼ��Ŀͻ��������ٷ�����Ϣ��
        if (!IsWindow(m_hwndEdit))
            return E_FAIL;

        POINT ptClient = *pt;
        if (!ScreenToClient(m_hwndEdit, &ptClient))
            return E_FAIL;

        LPARAM lparam = MAKELPARAM(ptClient.x, ptClient.y);
        LRESULT lr = SendMessage(m_hwndEdit, EM_CHARFROMPOS, 0, lparam);
        // EM_CHARFROMPOS ���ص��ַ�����λ�� LOWORD
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
     * ����ָ���ı���Χ����Ļ�򴰿�����ϵ�µľ�������pfClipped ָʾ�Ƿ񱻲ü���
     * ��ʵ�ּ�Ϊ���������༭�ؼ��ľ��Ρ�
     * @param vcView ��ͼ cookie
     * @param acpStart/acpEnd �ı���Χ
     * @param prc �������
     * @param pfClipped ����Ƿ񱻲ü�
     * @return S_OK �� E_INVALIDARG / E_FAIL
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

        // ��ʼ������ֵΪ FALSE
        *pfClipped = FALSE;

        // ���������������ĵ�����Ч��Χ�����������༭�ؼ�����
        if (acpStart < 0 || acpEnd < acpStart)
        {
            if (GetClientRect(m_hwndEdit, prc))
            {
                MapWindowPoints(m_hwndEdit, nullptr, (LPPOINT)prc, 2);
                return S_OK;
            }
            return E_FAIL;
        }

        // ��ȡ�ı�����
        LONG textLength = static_cast<LONG>(GetWindowTextLength(m_hwndEdit));
        if (acpStart >= textLength)
        {
            // λ�ó����ı���Χ�����ر༭�ؼ����½ǵ�һ��С����
            if (GetClientRect(m_hwndEdit, prc))
            {
                prc->left = prc->right - 2;
                prc->top = prc->bottom - 16;
                MapWindowPoints(m_hwndEdit, nullptr, (LPPOINT)prc, 2);
                return S_OK;
            }
            return E_FAIL;
        }

        // ��������λ��
        if (acpEnd > textLength)
            acpEnd = textLength;

        // ���ڵ��б༭�ؼ���ʹ�� EM_POSFROMCHAR
        DWORD style = GetWindowLong(m_hwndEdit, GWL_STYLE);
        BOOL isMultiLine = (style & ES_MULTILINE) != 0;

        if (!isMultiLine)
        {
            // ���б༭�ؼ� - ʹ�� EM_POSFROMCHAR
            LRESULT posStart = SendMessage(m_hwndEdit, EM_POSFROMCHAR, (WPARAM)acpStart, 0);
            LRESULT posEnd = SendMessage(m_hwndEdit, EM_POSFROMCHAR, (WPARAM)acpEnd, 0);

            if (posStart != -1 && posEnd != -1)
            {
                int xStart = GET_X_LPARAM(posStart);
                int yStart = GET_Y_LPARAM(posStart);
                int xEnd = GET_X_LPARAM(posEnd);
                int yEnd = GET_Y_LPARAM(posEnd);

                // �����ʼ�ͽ�����ͬһ��
                if (yStart == yEnd)
                {
                    prc->left = xStart;
                    prc->top = yStart;
                    prc->right = xEnd;

                    // ��ȡ�ַ��߶�
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
                            prc->bottom = yStart + 16; // ����ֵ
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
                    // �����ı������������༭�ؼ�����
                    if (!GetClientRect(m_hwndEdit, prc))
                        return E_FAIL;
                }
            }
            else
            {
                // EM_POSFROMCHAR ʧ�ܣ����������༭�ؼ�����
                if (!GetClientRect(m_hwndEdit, prc))
                    return E_FAIL;
            }
        }
        else
        {
            // ���б༭�ؼ� - ʹ�ø����ӵ�λ�ü���
            // �򻯴������ػ��ڲ����λ�õľ���
            LRESULT pos = SendMessage(m_hwndEdit, EM_POSFROMCHAR, (WPARAM)acpStart, 0);
            if (pos != -1)
            {
                int x = GET_X_LPARAM(pos);
                int y = GET_Y_LPARAM(pos);

                prc->left = x;
                prc->top = y;
                prc->right = x + 2; // С��ȱ�ʾ�����

                // ��ȡ�и�
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
                // ʧ��ʱ���������༭�ؼ�����
                if (!GetClientRect(m_hwndEdit, prc))
                    return E_FAIL;
            }
        }

        // ת��Ϊ��Ļ����
        MapWindowPoints(m_hwndEdit, nullptr, (LPPOINT)prc, 2);
        return S_OK;
    }

    /**
     * @brief ITextStoreACP::GetScreenExt
     *
     * ���ع�����������Ļ�����µ���Ӿ��Ρ�
     * @param vcView ��ͼ cookie
     * @param prc �������
     * @return S_OK �� E_INVALIDARG / E_FAIL
     */
    STDMETHODIMP GetScreenExt(TsViewCookie vcView, RECT* prc) override
    {
    logger << "[CTextStore::GetScreenExt] vcView=" << vcView << " prc=" << prc << "\n";
        if (!prc)
            return E_INVALIDARG;

        // ������Ļ����
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
        *pfOk = TRUE;                     // ͬ�⿪ʼ
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
        // ��Ͻ������ѱ����������ַ�����
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

// ��ʱ����
class CCompositionSink : public ITfCompositionSink
{
private:
    ULONG m_cRef = 0;
public:
    STDMETHODIMP OnCompositionTerminated(TfEditCookie ecWrite, ITfComposition* pComposition) override
    {
        // ��Ͻ�����
        ITfRange* pRange = NULL;
        if (SUCCEEDED(pComposition->GetRange(&pRange)))
        {
            // ������Ͻ���
            pRange->SetText(ecWrite, 0, L"", 0);
            pRange->Release();
        }
        return S_OK;
    }

    /**
 * @brief IUnknown::QueryInterface
 *
 * ��ѯ������֧�ֵĽӿ�ָ�루IUnknown �� ITextStoreACP����
 * @param riid ����Ľӿ� ID
 * @param ppvObj ����ӿ�ָ��
 * @return S_OK �� E_NOINTERFACE
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
     * �������ü����������µļ���ֵ��
     * @return �µ����ü���
     */
    STDMETHODIMP_(ULONG) AddRef() override
    {
        return InterlockedIncrement(&m_cRef);
    }

    /**
     * @brief IUnknown::Release
     *
     * �������ü���������Ϊ 0 ʱ���ٶ��󣬷���ʣ�������
     * @return ʣ�����ü���
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

// ʵ����������Ͻӿ�
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