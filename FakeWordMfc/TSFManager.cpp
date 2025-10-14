#include "pch.h"
#include "TSFManager.h"

BOOL CTSFManager::Initialize(HWND hwndEdit)
{
    m_hwndEdit = hwndEdit;

    // 创建TSF线程管理器
    HRESULT hr = CoCreateInstance(CLSID_TF_ThreadMgr,
        nullptr,
        CLSCTX_INPROC_SERVER,
        IID_ITfThreadMgr,
        (void**)&m_pThreadMgr);
    if (FAILED(hr) || !m_pThreadMgr)
        return FALSE;

    // 激活线程管理器
    hr = m_pThreadMgr->Activate(&m_clientId);
    if (FAILED(hr))
        return FALSE;

    // 创建文本存储
    m_pTextStore = new CTextStore(this, hwndEdit);
    if (!m_pTextStore)
        return FALSE;

    // 创建文档管理器
    hr = m_pThreadMgr->CreateDocumentMgr(&m_pDocumentMgr);
    if (FAILED(hr) || !m_pDocumentMgr)
    {
        return FALSE;
    }

    // 使用文本存储创建上下文
    hr = m_pDocumentMgr->CreateContext(m_clientId, 0,
        static_cast<ITextStoreACP*>(m_pTextStore),
        &m_pContext, &m_ContextCookie);
    if (FAILED(hr) || !m_pContext)
        return FALSE;

    // 将文档管理器设置为活动状态
    hr = m_pDocumentMgr->Push(m_pContext);
    if (FAILED(hr))
        return FALSE;

    // 将TSF与窗口关联
    ITfDocumentMgr* docMgr = nullptr;
    hr = m_pThreadMgr->AssociateFocus(hwndEdit, m_pDocumentMgr, &docMgr);
    if (FAILED(hr))
        return FALSE;

    // 初始化文本存储
    m_pTextStore->Initialize(m_pContext, m_clientId);


    return TRUE;
}

void CTSFManager::Uninitialize()
{
    if (m_pThreadMgr && m_hwndEdit)
    {
        ITfDocumentMgr* docMgr = nullptr;
        m_pThreadMgr->AssociateFocus(m_hwndEdit, nullptr, &docMgr);
    }

    if (m_pDocumentMgr)
    {
        m_pDocumentMgr->Pop(TF_POPF_ALL);
        m_pDocumentMgr->Release();
        m_pDocumentMgr = nullptr;
    }

    if (m_pContext)
    {
        m_pContext->Release();
        m_pContext = nullptr;
    }

    if (m_pTextStore)
    {
        m_pTextStore->Release();
        m_pTextStore = nullptr;
    }

    if (m_pThreadMgr)
    {
        if (m_clientId != TF_CLIENTID_NULL)
        {
            m_pThreadMgr->Deactivate();
            m_clientId = TF_CLIENTID_NULL;
        }
        m_pThreadMgr->Release();
        m_pThreadMgr = nullptr;
    }
}

