#include "pch.h"
#include "WordController.h"


std::wstring HResultToWString(HRESULT hr)
{
    WCHAR* lpBuffer = new WCHAR[sizeof(WCHAR*) * 8];
    // 获取错误消息的长度
    DWORD messageSize = FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        NULL,
        hr,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        lpBuffer,
        0,
        NULL);
    delete[] lpBuffer;
    lpBuffer = new WCHAR[messageSize];
    //MessageBox(0, std::to_wstring(messageSize).c_str(), L"messageSize", 0);
    // 将错误消息复制到字符串
    FormatMessage(
        FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        NULL,
        hr,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        lpBuffer,
        messageSize,
        NULL);
    std::wstring message = lpBuffer;
    delete[] lpBuffer;
    return message;
}

WordController::~WordController()
{
	// 析构时自动调用Release，因此不需要手动调用
	//pWordApp.Release();
    for (auto it = eventSinkMap.begin(); it != eventSinkMap.end(); it++)
    {
        it->second->detach();
        it->second->Release();
        //delete it->second;
    }
    eventSinkMap.clear();
    // 智能指针不需要手动调用Release
    pWordDocumentMap.clear();
    pWordWindowMap.clear();

    if (pWordApp)
    {
        try {
            pWordApp->Application->Quit();
        }
        COM_CATCH_NORESULT(L"Word quit error.")
        COM_CATCH_ALL_NORESULT(L"Word quit error.")
    }
}

WordController::WordController()
{
    //auto initResult = comInitialize();
    //if (!initResult)
    //{
    //    std::wstring msg = L"无法初始化COM服务！\n原因：\n";
    //    msg += initResult.getMessage();
    //    MessageBeep(MB_ICONERROR);
    //    MessageBox(0, msg.c_str(), L"Error", MB_ICONERROR);
    //    //return;
    //}
    auto createResult = createApp();
    if (!createResult)
    {
        std::wstring msg = L"无法打开Word！\n原因：\n";
        msg += createResult.getMessage();
        MessageBeep(MB_ICONERROR);
        MessageBox(0, msg.c_str(), L"Error", MB_ICONERROR);
        return;
    }
    
}

Result<Word::_DocumentPtr> WordController::getActiveDocument() const
{
    Result<Word::_DocumentPtr> docRst(false, 0, L"Unknown error.");
    try {
        Word::_DocumentPtr docPtr = pWordApp->ActiveDocument;
        return SuccessResult(0, docPtr);
    }
    COM_CATCH_NOMSGBOX(L"无法获取当前活动文档！", docRst)
    COM_CATCH_ALL_NOMSGBOX(L"无法获取当前活动文档！", docRst)
    return docRst;
}

Result<Word::_ApplicationPtr> WordController::createApp()
{
    auto& app = pWordApp;
    HRESULT hr = app.CreateInstance(__uuidof(Word::Application)); // 创建Word应用实例
    return Result<Word::_ApplicationPtr>::CreateFromHRESULT(hr, app);
}

Result<Word::_DocumentPtr> WordController::openDocument(IdType docId, std::wstring file, bool ConfirmConversions, bool ReadOnly, bool AddToRecentFiles, std::wstring PasswordDocument, std::wstring PasswordTemplate, bool Revert, std::wstring WritePasswordDocument, std::wstring WritePasswordTemplate, Word::WdOpenFormat Format, Office::MsoEncoding Encoding, bool Visible, bool OpenConflictDocument, bool OpenAndRepair, Word::WdDocumentDirection DocumentDirection, bool NoEncodingDialog)
{
    // 原生类型 -> _variant_t
    _variant_t vFile(file.c_str());
    _variant_t vConfirm(ConfirmConversions);
    _variant_t vReadOnly(ReadOnly);
    _variant_t vAddRecent(AddToRecentFiles);
    _variant_t vPwdDoc(PasswordDocument.c_str());
    _variant_t vPwdTpl(PasswordTemplate.c_str());
    _variant_t vRevert(Revert);
    _variant_t vWrPwdDoc(WritePasswordDocument.c_str());
    _variant_t vWrPwdTpl(WritePasswordTemplate.c_str());
    _variant_t vFormat(Format);
    _variant_t vEnc(Encoding);
    _variant_t vVisible(Visible);
    _variant_t vConflict(OpenConflictDocument);
    _variant_t vRepair(OpenAndRepair);
    _variant_t vDir(DocumentDirection);
    _variant_t vNoEnc(NoEncodingDialog);
    Result<Word::_DocumentPtr> result(true);
    try {
        //pWordApp->Documents->Open(
        //    _com_util::ConvertStringToBSTR(wstr2str_2ANSI(file).c_str()),
        //    ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate,
        //    Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding,
        //    Visible, OpenConflictDocument, OpenAndRepair, DocumentDirection, NoEncodingDialog
        //);
        Word::DocumentsPtr dsp = pWordApp->Documents;
        Word::_DocumentPtr docPtr = pWordApp->Documents->Open(
            &vFile, &vConfirm, &vReadOnly, &vAddRecent,
            &vPwdDoc, &vPwdTpl, &vRevert,
            &vWrPwdDoc, &vWrPwdTpl,
            &vFormat, &vEnc,
            &vVisible, &vConflict, &vRepair,
            &vDir, &vNoEnc);
        pWordApp->PutVisible(Word::xlMaximum);
        pWordDocumentMap[docId] = docPtr;
        return SuccessResult(0, docPtr);
    }
    COM_CATCH(L"无法打开文档！", result)
    COM_CATCH_ALL(L"无法打开文档！", result)
    return result;
}

Result<Word::_DocumentPtr> WordController::addDocument(IdType id, bool visible, std::wstring templateName, bool asNewTemplate, Word::WdNewDocumentType documentType)
{
    _variant_t vTemplateName(templateName.empty() ? vtMissing : templateName.c_str());
    _variant_t vAsNewTemplate(asNewTemplate);
    _variant_t vDocumentType((long)documentType);
    _variant_t vVisible(visible);
    Result<Word::_DocumentPtr> result(true);
    try {
        Word::DocumentsPtr dsp = pWordApp->Documents;
        //Word::_DocumentPtr docPtr = pWordApp->Documents->Add();
        Word::_DocumentPtr docPtr = pWordApp->Documents->Add(&vTemplateName, &vAsNewTemplate, &vDocumentType, &vVisible);
        pWordDocumentMap[id] = docPtr;
        return SuccessResult(0, docPtr);
    }
    COM_CATCH(L"无法添加新文档！", result)
    COM_CATCH_ALL(L"无法添加新文档！", result)
    return result;
}

Result<Word::WindowPtr> WordController::newWindow(IdType id)
{
    Result<Word::WindowPtr> result(true);
    try {
        Word::WindowPtr windowPtr = pWordApp->NewWindow();
        pWordWindowMap[id] = windowPtr;
        return SuccessResult(0, windowPtr);
    }
    COM_CATCH(L"无法创建新窗口！", result)
    COM_CATCH_ALL(L"无法创建新窗口！", result)
    return result;
}


Result<bool> WordController::setAppVisible(bool visible) const
{
    _variant_t vVisible(visible);
    Result<bool> result(true);
    try {
        pWordApp->PutVisible(vVisible.boolVal);
        return SuccessResult(0, true);
    }
    COM_CATCH(L"无法设置窗口可见性", result)
    COM_CATCH_ALL(L"无法设置窗口可见性", result)
    return result;
}

//// Define type info structure
//_ATL_FUNC_INFO EventSink::OnDocChangeInfo = { CC_STDCALL, VT_EMPTY, 0 };
//_ATL_FUNC_INFO EventSink::OnQuitInfo = { CC_STDCALL, VT_EMPTY, 0 };
//// (don't actually need two structure since they're the same)