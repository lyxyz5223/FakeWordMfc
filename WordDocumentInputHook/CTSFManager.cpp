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
        return FALSE;

    // 使用文本存储创建上下文
    hr = m_pDocumentMgr->CreateContext(m_clientId, 0,
        m_pTextStore,
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

void CTSFManager::HandleTextInput(const std::wstring& text)
{
    // 这里可以添加文本过滤逻辑
    if (!ShouldBlockInput(text))
    {
        // 将文本添加到编辑框
        AddTextToEditControl(text);

        // 更新文本存储
        if (m_pTextStore)
        {
            std::wstring currentText = m_pTextStore->GetText();
            currentText += text;
            m_pTextStore->SetText(currentText);
        }
    }
    else
    {
        MessageBeep(MB_ICONWARNING);
    }
}

BOOL CTSFManager::ShouldBlockInput(const std::wstring& text)
{
    // 实现过滤逻辑
    for (wchar_t c : text)
    {
        if (c >= L'0' && c <= L'9')
            return TRUE;
    }

    static const std::vector<std::wstring> blockedWords = {
        L"badword", L"sensitive", L"blockme", 
        L"叼你妈", L"艹"
    };

    for (const auto& word : blockedWords)
    {
        if (text.find(word) != std::wstring::npos)
            return TRUE;
    }

    return FALSE;
}

void CTSFManager::AddTextToEditControl(const std::wstring& text)
{
    if (!m_hwndEdit) return;

    // 获取当前文本
    int len = GetWindowTextLengthW(m_hwndEdit);
    std::wstring currentText(len + 1, L'\0');
    GetWindowTextW(m_hwndEdit, &currentText[0], len + 1);
    currentText.resize(len);

    // 添加新文本
    currentText += text;

    // 设置新文本
    SetWindowTextW(m_hwndEdit, currentText.c_str());
}
