
// FakeWordMfcDlg.cpp: 实现文件
//

#include "pch.h"
#include "framework.h"
#include "FakeWordMfc.h"
#include "FakeWordMfcDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#include "StringProcess.h"
#include <sstream>
#include "UIAutomation.h"
#include "CStateBoxDlg.h"
#include "CSettingDialog.h"

struct EnumWindowsStruct {
	CFakeWordMfcDlg* ptr;
	void* userData;
};
// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CFakeWordMfcDlg 对话框



CFakeWordMfcDlg::CFakeWordMfcDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_FAKEWORDMFC_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
#ifdef _DEBUG
	{
		AllocConsole();
		FILE* stream = nullptr;
		freopen_s(&stream, "CONOUT$", "w", stdout);
		freopen_s(&stream, "CONOUT$", "w", stderr);
	}
#endif // !_DEBUG
}

void CFakeWordMfcDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_BTNOPENFILE, m_BtnOpenFile);
	DDX_Control(pDX, IDC_EDITSLACKINGOFFFILEPATH, m_EditFishFilePath);
	DDX_Control(pDX, IDC_EDITWORKFILEPATH, m_EditWorkFilePath);
	DDX_Control(pDX, IDC_BTNSELECTFISHFILE, m_BtnSelectFishFile);
	DDX_Control(pDX, IDC_BTNSELECTWORKFILE, m_BtnSelectWorkFile);
}

BEGIN_MESSAGE_MAP(CFakeWordMfcDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BTNOPENFILE, &CFakeWordMfcDlg::OnBnClickedBtnOpenFile)
	ON_BN_CLICKED(IDC_BTNSELECTFISHFILE, &CFakeWordMfcDlg::OnBnClickedBtnSelectFishFile)
	ON_BN_CLICKED(IDC_BTNSELECTWORKFILE, &CFakeWordMfcDlg::OnBnClickedBtnSelectWorkFile)
	ON_BN_CLICKED(IDC_BTNOPENWORD, &CFakeWordMfcDlg::OnBnClickedBtnOpenWord)
	ON_WM_GETMINMAXINFO()
	ON_WM_CLOSE()
	ON_WM_HOTKEY()
	ON_BN_CLICKED(IDC_BTNSETTING, &CFakeWordMfcDlg::OnClickedBtnSetting)
END_MESSAGE_MAP()


// CFakeWordMfcDlg 消息处理程序
//#include <msctf.h>
//#include "TSFInterceptor.h"
//#include "TSFManager.h"
BOOL CFakeWordMfcDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	// 启用现代 ToolTip
	CMFCToolTipInfo params;
	params.m_bVislManagerTheme = TRUE;

	// 创建现代风格 ToolTip
	m_pToolTip = new CMFCToolTipCtrl(&params);
	if (m_pToolTip->Create(this))
	{
		m_pToolTip->AddTool(GetDlgItem(IDC_BTNOPENWORD), _T("若多次点击，只有最后一次点击打开的Word/WPS Office生效"));
		m_pToolTip->Activate(TRUE);
	}
	//打开按钮图标
	SHSTOCKICONINFO sii = { 0 };
	sii.cbSize = sizeof(sii);
	SHGetStockIconInfo(SIID_FOLDEROPEN, SHGSI_ICON | SHGSI_SMALLICON, &sii);
	m_BtnSelectFishFile.SetIcon(sii.hIcon);// AfxGetApp()->LoadIcon(IDC_ICON);
	m_BtnSelectWorkFile.SetIcon(sii.hIcon);// AfxGetApp()->LoadIcon(IDC_ICON);

	// 读取配置文件
	wordDetectMode = WordDetectMode::fromNumber(configManager.readConfigInt(configSection, configKeys[0], wordDetectMode.toInt()));
	try {
		oiRatio = std::stold(configManager.readConfigString(configSection, configKeys[1], std::to_wstring(oiRatio)));
	}
	catch (...) {

	}
	int vkc = configManager.readConfigInt(configSection, configKeys[2], switchReplaceModeHotkey.virtualKeyCode);
	int mods = configManager.readConfigInt(configSection, configKeys[3], switchReplaceModeHotkey.modifiers);
	switchReplaceModeHotkey.virtualKeyCode = vkc;
	switchReplaceModeHotkey.modifiers = mods;

	// 默认模式为摸鱼，设置默认切换快捷键为Ctrl+Alt+Shift+S
	::ShowWindow(GetShellWindow(), SW_SHOW);
	// MOD_NOREPEAT确保按着快捷键时只会收到一次消息而不会一直收到
	if (!RegisterHotKey(GetSafeHwnd(), ID_HOTKEY_DEFAULT_SWITCH_REPLACE_MODE, switchReplaceModeHotkey.modifiers, switchReplaceModeHotkey.virtualKeyCode))
	{
		MessageBox(L"用于切换替换文本的模式的全局快捷键注册失败！", L"Error", MB_ICONERROR);
	}

	WordController::comInitialize();
	//CoInitializeEx(0, COINIT_APARTMENTTHREADED);
	// 测试输入框拦截
	// 将当前软件的一个输入框设置为焦点，用做测试
	//this->GetDlgItem(IDC_EDITSLACKINGOFFFILEPATH)->SetFocus();
	//CTSFManager* tsfMgr = new CTSFManager();
	//tsfMgr->Initialize(GetDlgItem(IDC_EDITSLACKINGOFFFILEPATH)->GetSafeHwnd());

	// CComPtr 类型
	//pWordApp.CoCreateInstance(
	//	L"Word.Application",			// ProgID
	//	nullptr,
	//	CLSCTX_ALL);


	// TODO: 窗口初始化测试代码
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CFakeWordMfcDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CFakeWordMfcDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CFakeWordMfcDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CFakeWordMfcDlg::OnBnClickedBtnSelectFishFile()
{
	// 打开小说文件
	const TCHAR extsFilter[] =
		_T("文本文档 (*.txt)|*.txt|")
		_T("legado开源阅读导出小说文件 (*.txt)|*.txt|")
		_T("legado开源阅读下载文件 (*.nb)|*.nb|")
		_T("Microsoft Word 文档 (*.docx, *.doc)|*.docx;*.doc|") // ';'分割
		_T("All Files (*.*)|*.*||");//';'号间隔多个后缀名

	//获取当前工作路径作为文件打开的默认路径
	//auto bufferLen = MAX_PATH;
	//TCHAR* defaultDir = new TCHAR[bufferLen];
	//DWORD needLen = GetCurrentDirectory(bufferLen, defaultDir);
	//if (needLen > bufferLen)
	//{
	//	delete[] defaultDir;
	//	defaultDir = new TCHAR[needLen];
	//	GetCurrentDirectory(needLen, defaultDir);
	//}
	//下文中记得要delete[] defaultdir;释放内存!

	CFileDialog fileSelectionDlg(TRUE/*TRUE为"打开"窗口,FALSE为"另存为"窗口*/,
		NULL/*默认自动追加的文件扩展名*/,
		L""/*"文件名"一栏默认填充文本*/,
		NULL/*自定义属性*/,
		extsFilter/*文件扩展名过滤*/,
		this/*父窗口*/,
		0/*文件选择窗口的版本*/,
		TRUE/*是否启用新样式文件选择窗口*/
	);
	INT_PTR result = fileSelectionDlg.DoModal();//打开窗口并等待返回
	//delete[] defaultDir;
	if (result != IDOK)
		return;
	//else
	//点击了 "打开"/"保存" 按钮
	std::wstring filePath = fileSelectionDlg.GetPathName();
	//std::wstring fileName = fileSelectionDlg.GetFileName();
	m_EditFishFilePath.SetWindowText(filePath.c_str());
	fishFileTxtReader.setFilePath(filePath);
	selectedFishFile.filePath = filePath;
	selectedFishFile.fileName = fileSelectionDlg.GetFileName();
	selectedFishFile.fileExtension = fileSelectionDlg.GetFileExt();
	if (selectedFishFile.fileExtension == L".docx" || selectedFishFile.fileExtension == L".doc")
		selectedFishFile.fileType = SupportedFileType::Word;
	else
		selectedFishFile.fileType = SupportedFileType::Txt;
}

void CFakeWordMfcDlg::OnBnClickedBtnSelectWorkFile()
{
	// 打开小说文件
	const TCHAR extsFilter[] =
		_T("Microsoft Word 文档 (*.docx, *.doc)|*.docx;*.doc|") // ';'分割
		_T("文本文档 (*.txt)|*.txt|")
		_T("legado开源阅读导出小说文件 (*.txt)|*.txt|")
		_T("legado开源阅读下载文件 (*.nb)|*.nb|")
		_T("All Files (*.*)|*.*||");//';'号间隔多个后缀名
	CFileDialog fileSelectionDlg(TRUE/*TRUE为"打开"窗口,FALSE为"另存为"窗口*/,
		NULL/*默认自动追加的文件扩展名*/,
		L""/*"文件名"一栏默认填充文本*/,
		NULL/*自定义属性*/,
		extsFilter/*文件扩展名过滤*/,
		this/*父窗口*/,
		0/*文件选择窗口的版本*/,
		TRUE/*是否启用新样式文件选择窗口*/
	);
	INT_PTR result = fileSelectionDlg.DoModal();//打开窗口并等待返回
	if (result != IDOK)
		return;
	//else
	//点击了 "打开"/"保存" 按钮
	std::wstring fExt = fileSelectionDlg.GetFileExt();
	std::wstring fPath = fileSelectionDlg.GetPathName();
	std::wstring fName = fileSelectionDlg.GetFileName();
	m_EditWorkFilePath.SetWindowText(fPath.c_str());
	workFileTxtReader.setFilePath(fPath);
	selectedFishFile.filePath = fPath;
	selectedFishFile.fileName = fileSelectionDlg.GetFileName();
	selectedFishFile.fileExtension = fileSelectionDlg.GetFileExt();
	if (selectedFishFile.fileExtension == L".docx" || selectedFishFile.fileExtension == L".doc")
		selectedFishFile.fileType = SupportedFileType::Word;
	else
		selectedFishFile.fileType = SupportedFileType::Txt;

}


void CFakeWordMfcDlg::OnBnClickedBtnOpenFile()
{
	switch (selectedFishFile.fileType)
	{
	case SupportedFileType::Word:
	{
		break;
	}
	default:
	case SupportedFileType::Txt:
	{
		if (!fishFileTxtReader.open())
		{
			MessageBeep(MB_ICONERROR);
			MessageBox(L"糟糕！小说打开失败了呜呜呜~", L"布豪", MB_ICONERROR);
		}
		else
		{
			fishFileTxtReader.readLines();
			processedFishFileTxtWordText = fishFileTxtReader.getFullTextConstReference();
			std::replace(processedFishFileTxtWordText.begin(), processedFishFileTxtWordText.end(), L'\n', L'\r');
		}

		break;
	}
	}
	switch (selectedFishFile.fileType)
	{
	case SupportedFileType::Word:
	{
		break;
	}
	default:
	case SupportedFileType::Txt:
	{
		if (!workFileTxtReader.open())
		{
			MessageBeep(MB_ICONERROR);
			MessageBox(L"糟糕！无法打开正经文件。", L"错误", MB_ICONERROR);
		}
		else
		{
			workFileTxtReader.readLines();
			processedWorkFileTxtWordText = workFileTxtReader.getFullTextConstReference();
			std::replace(processedWorkFileTxtWordText.begin(), processedWorkFileTxtWordText.end(), L'\n', L'\r');
		}

		break;
	}
	}
}



void CFakeWordMfcDlg::OnBnClickedBtnOpenWord()
{
	auto newId = getAvailableWordWindowId();
	{
		//// 下面代码用作简单调试
		//auto initHResult = CoInitialize(0);
		////CONNECT_E_NOCONNECTION;
		//Word::_ApplicationPtr pWordApp;
		//HRESULT hr = pWordApp.CreateInstance(__uuidof(Word::Application)); // 创建Word应用实例
		//// 实例化并存储事件槽对象
		//CExcelEventSink* pWordAppSink = new CExcelEventSink(GetSafeHwnd());
		//// 获取Word事件连接点容器
		//CComPtr<IConnectionPointContainer> icpc;
		//hr = pWordApp.QueryInterface(IID_IConnectionPointContainer, &icpc);
		//// 获取Word事件连接点
		//CComPtr<IConnectionPoint> icp;
		//hr = icpc->FindConnectionPoint(__uuidof(Word::ApplicationEvents4), &icp);
		//// 将事件接收器绑定到连接点
		//DWORD dwCookie;
		//hr = icp->Advise(pWordAppSink, &dwCookie);
		//Result adviseResult = Result<int>::CreateFromHRESULT(hr, 0);
		//if (hr == E_NOINTERFACE)
		//{
		//	MessageBox(L"E_NOINTERFACE: 查询接口失败，未实现该接口", L"", MB_ICONERROR);
		//}
	}
	if (wordCtrller)
	{
		destroyStateBoxDlg();
		try {
			wordCtrller->getApp()->Quit();
		}
		catch (...) {

		}
		delete wordCtrller;
	}
	wordCtrller = new WordController();
	//wordTextChangeDetecterThread = std::thread(&CFakeWordMfcDlg::wordTextChangeDetecterWorker, this);

	Result<bool> appRegEvtSnkRslt = wordCtrller->registerEventSink(wordCtrller->getApp(), __uuidof(Word::ApplicationEvents4));
	//wordCtrller->openDocument(L"C:\\Users\\lyxyz5223\\Desktop\\test.docx");
	wordCtrller->addDocument(newId, true);
	try {
		Word::WindowPtr windowPtr = wordCtrller->getApp()->ActiveWindow;
		windowPtr->PutVisible(VARIANT_TRUE);
		try {
			wordAppHwnd = reinterpret_cast<HWND>(windowPtr->GetHwnd());
		}
		COM_CATCH_NORESULT(L"无法获取Word主窗口句柄")
		COM_CATCH_ALL_NORESULT(L"无法获取Word主窗口句柄")
	}
	COM_CATCH_NORESULT(L"Word主窗口获取失败")
	COM_CATCH_ALL_NORESULT(L"Word主窗口获取失败")



	CComPtr<Word::_Document> docPtr = wordCtrller->getDocument(newId).getData();
	Result<bool> docRegEvtSnkRslt = wordCtrller->registerEventSink(docPtr.p, __uuidof(Word::DocumentEvents2));
	// 释放document接口，防止提前释放导致内存管理异常
	//docPtr.Detach();
	wordCtrller->registerEvent(__uuidof(Word::ApplicationEvents4), 6/*文档即将关闭*/, [this](DISPPARAMS* params, WordEventSink::UserDataType userData) {
		return this->onWordDocumentBeforeClose(params, userData);
		}, 0);
	wordCtrller->registerEvent(__uuidof(Word::ApplicationEvents4), 10/*窗口激活*/, [this](DISPPARAMS* params, WordEventSink::UserDataType userData) {
		return this->onWordWindowActivate(params, userData);
		}, 0);
	//wordCtrller->registerEvent(__uuidof(Word::ApplicationEvents4), 11/*窗口去激活*/, [this](DISPPARAMS* params, WordEventSink::UserDataType userData) {
	//	return this->onWordDocumentBeforeClose(params, userData);
	//	}, 0);
	findWordDocumentHWND(); // 先获取一遍文档HWND

	switch (wordDetectMode)
	{
	default:
	case WordDetectMode::AppEvent:
	{
		wordCtrller->registerEvent(
			__uuidof(Word::ApplicationEvents4),
			12/*WindowSelectionChange*/,
			[this](DISPPARAMS* params, WordEventSink::UserDataType userData) {
				return this->onWordWindowSelectionChangeEvent(params, userData);
			}, 0);
		break;
	}
	case WordDetectMode::LoopCheck:
	{
		if (!wordTextChangeDetecterThread.joinable())
			wordTextChangeDetecterThread = std::thread(&CFakeWordMfcDlg::wordTextChangeDetecterWorker, this);
		break;
	}
	case WordDetectMode::HotKey:
	{
		// 未实现
		break;
	}
	case WordDetectMode::DllInject:
	{
		// 注入dll到目标窗口进程
#ifdef _DEBUG
		HMODULE injectDllHandle = LoadLibrary(L"D:\\C++\\FakeWordMfc\\x64\\Debug\\WordDocumentInputHook.dll");
#else
		HMODULE injectDllHandle = LoadLibrary(L"WordDocumentInputHook.dll");
#endif
		if (injectDllHandle)
		{
			static long long injectId = getAvailableInjectId();
			std::string injectIdStr = std::to_string(injectId);
			typedef HRESULT(*InjectToProcessByWindowHwndFunctionType)(HWND hwnd, const char* injectId);
			InjectToProcessByWindowHwndFunctionType InjectToProcessByWindowHwnd = (InjectToProcessByWindowHwndFunctionType)GetProcAddress(injectDllHandle, "InjectToProcessByWindowHwnd");
			if (InjectToProcessByWindowHwnd)
			{
				HRESULT hr = InjectToProcessByWindowHwnd(wordDocumentHwnd, injectIdStr.c_str());
				if (SUCCEEDED(hr))
					MessageBox(L"成功：将dll注入Word文档编辑框窗口进程！", L"", MB_ICONINFORMATION);
				else
				{
					std::wstring errMsg = HResultToWString(hr);
					MessageBeep(MB_ICONERROR);
					MessageBox((L"无法将dll注入Word文档编辑框窗口进程，软件功能可能受影响\n\t错误代码：" + std::to_wstring(hr) + L"\n\t错误消息: " + errMsg).c_str(), L"", MB_ICONERROR);
				}
			}
		}
		else
		{
			auto win32Error = GetLastError();
			std::wstring errMsg = HResultToWString(HRESULT_FROM_WIN32(win32Error));
			MessageBeep(MB_ICONERROR);
			MessageBox((L"无法将dll注入Word文档编辑框窗口进程，软件功能可能受影响\n\t错误代码：" + std::to_wstring(win32Error) + L"\n\t错误消息: " + errMsg).c_str(), L"", MB_ICONERROR);
		}
		break;
	}
	}
	
	// 嵌入状态窗口
	std::unique_lock lock(mtxEmbeddedStateDlgMap);
	wordWindowIdSet.insert(newId);
	if (!embeddedStateDlg)
	{
		//CWnd* wordDocumentCwnd = CWnd::FromHandle(wordDocumentHwnd);
		newStateBoxDlg(wordAppHwnd);
		updateStateBoxDlg(embeddedStateDlg);
	}
}

//InputInterceptor* interceptor = new InputInterceptor();
//if (interceptor->Initialize()) {
//	wprintf(L"输入拦截器初始化成功\n");
//	// 注册全局焦点改变事件
//	if (interceptor->RegisterTextChangeEventForWindow(wordDocumentHwnd)) {
//		wprintf(L"开始监听输入事件...\n");
//		wprintf(L"按任意键退出\n");
//	}
//	if (interceptor->RegisterTextChangeEventForWindow(wordAppHwnd)) {
//		wprintf(L"开始监听输入事件...\n");
//		wprintf(L"按任意键退出\n");
//	}
//}
//else {
//	wprintf(L"输入拦截器初始化失败\n");
//}
//std::thread([=]() {
//		WordTextMonitor* wtm = new WordTextMonitor();
//		wtm->Initialize(wordDocumentHwnd);
//		while (true)
//		{
//			wtm->CheckForChanges();
//			Sleep(100);
//		}
//	}
//).detach();


void CFakeWordMfcDlg::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
	// TODO: 在此添加消息处理程序代码和/或调用默认值

	CDialogEx::OnGetMinMaxInfo(lpMMI);
	if (windowMinSize.cx == 0 && windowMinSize.cy == 0)
	{
		CRect windowRect;
		GetWindowRect(&windowRect);
		windowMinSize.SetSize(windowRect.Width(), windowRect.Height());
	}
	lpMMI->ptMinTrackSize.x = windowMinSize.cx;
	lpMMI->ptMinTrackSize.y = windowMinSize.cy;
}

void CFakeWordMfcDlg::wordTextChangeDetecterWorker()
{
	while (!objDeconstructing)
	{
		if (!wordLastText.size())
		{
			// 之前是空的话先更新一次，实际上有一个\r字符
			auto rstGetActiveDoc = wordCtrller->getActiveDocument();
			if (!rstGetActiveDoc)
				continue;
			try {
				Word::_DocumentPtr docPtr = rstGetActiveDoc.getData();
				Word::RangePtr mainTextRangePtr = docPtr->Range();
				std::wstring currentText = mainTextRangePtr->Text;
				wordLastText = currentText;
			} catch (...) {}
		}
		else
			processTextChange();
		Sleep(100);
	}
}

Result<bool> CFakeWordMfcDlg::onWordWindowSelectionChangeEvent(DISPPARAMS* params, WordEventSink::UserDataType userData)
{
	IDispatch* disp = params->rgvarg[0].pdispVal;
	CComPtr<IDispatch> selDisp(disp);
	DISPID dispid;
	OLECHAR* name = (OLECHAR*)L"Text";
	HRESULT hr = selDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_NAME_USER_DEFAULT, &dispid);
	if (FAILED(hr))
		return Result<bool>::CreateFromHRESULT(hr, true);
	VARIANT result = { 0 };
	DISPPARAMS invokeDispParams = { 0 };
	hr = selDisp->Invoke(dispid, IID_NULL, LOCALE_NAME_USER_DEFAULT, DISPATCH_PROPERTYGET, &invokeDispParams, &result, nullptr, nullptr);
	if (FAILED(hr))
		return Result<bool>::CreateFromHRESULT(hr, true);
	std::wstring selectedText;
	if (result.bstrVal)
		selectedText += result.bstrVal;
	VariantClear(&result);
	std::cout << "Selected text: " << wstr2str_2ANSI(selectedText) << std::endl;
	return processTextChange();
}

Result<bool> CFakeWordMfcDlg::processTextChange()
{
	auto rstGetActiveDoc = wordCtrller->getActiveDocument();
	if (!rstGetActiveDoc)
		return ErrorResult<bool>(rstGetActiveDoc);
	Result<bool> rst(true, 0, true);
	try {
		Word::_DocumentPtr docPtr = rstGetActiveDoc.getData();
		docPtr->Application->ScreenUpdating = VARIANT_FALSE;
		try {
			docPtr->Application->UndoRecord->StartCustomRecord(_bstr_t(L"Fish update"));
			try {
				Word::RangePtr mainTextRangePtr = docPtr->Range();
				std::wstring currentText = mainTextRangePtr->Text;
				auto lenCur = currentText.size() - 1;
				auto lenLast = wordLastText.size() - 1;

				// 寻找差异位置
				std::pair<std::wstring::iterator, std::wstring::iterator> diffResult;
				std::reference_wrapper<std::wstring> txtFullText(processedFishFileTxtWordText);
				switch (replaceMode)
				{
				case ReplaceMode::Work:
				{
					diffResult = std::mismatch(currentText.begin(), currentText.end(), processedWorkFileTxtWordText.begin(), processedWorkFileTxtWordText.end());
					txtFullText = processedWorkFileTxtWordText;
					break;
				}
				default:
				case ReplaceMode::SlackOff:
				{
					diffResult = std::mismatch(currentText.begin(), currentText.end(), processedFishFileTxtWordText.begin(), processedFishFileTxtWordText.end());
					txtFullText = processedFishFileTxtWordText;
					break;
				}
				}
				// 返回第一个不匹配字符的位置
				auto diffStart = std::distance(currentText.begin(), diffResult.first);
				//if (lenCur > lenLast)
				//{
				//	// 新增小说文本
				//	wordTextUpdateTxtSubTextRange.start = diffStart;
				//	//wordTextUpdateTxtSubTextRange.start = lenLast - 1;
				//	//wordTextUpdateTxtSubTextRange.length = lenCur - lenLast;
				//	wordTextUpdateTxtSubTextRange.end = lenCur;
				//	wordTextUpdateTxtSubTextRange.newLength = (lenCur - diffStart) * GetEditNumber(IDC_EDITOIRATIO);
				//	updateWordText(docPtr, txtFullText.get(), wordTextUpdateTxtSubTextRange.start, wordTextUpdateTxtSubTextRange.end, wordTextUpdateTxtSubTextRange.newLength);
				//}
				//else if (lenCur < lenLast)
				//{
				//	wordTextUpdateTxtSubTextRange.start = 0;
				//	wordTextUpdateTxtSubTextRange.end = lenCur;
				//	wordTextUpdateTxtSubTextRange.newLength = lenCur - 0;
				//	updateWordText(docPtr, txtFullText.get(), wordTextUpdateTxtSubTextRange.start, wordTextUpdateTxtSubTextRange.end, wordTextUpdateTxtSubTextRange.newLength);
				//}
				auto& rangeStart = wordTextUpdateTxtSubTextRange.start = diffStart;
				auto& rangeEnd = wordTextUpdateTxtSubTextRange.end = lenCur;
				auto& newLen = wordTextUpdateTxtSubTextRange.newLength = (static_cast<long double>(lenCur - diffStart)) * oiRatio;
				if (!(rangeStart == rangeEnd && newLen == 0))
				{
					updateWordText(docPtr, txtFullText.get(), rangeStart, rangeEnd, newLen);
					Word::RangePtr newMainTextRangePtr = docPtr->Range();
					std::wstring newCurText = newMainTextRangePtr->Text;
					wordLastText = newCurText;
				}
			}
			COM_CATCH_NOMSGBOX(L"", rst)
			COM_CATCH_ALL_NOMSGBOX(L"", rst)
			docPtr->Application->UndoRecord->EndCustomRecord();
		}
		COM_CATCH_NOMSGBOX(L"", rst)
		COM_CATCH_ALL_NOMSGBOX(L"", rst)
		docPtr->Application->ScreenUpdating = VARIANT_TRUE;
	}
	COM_CATCH_NOMSGBOX(L"", rst)
	COM_CATCH_ALL_NOMSGBOX(L"", rst)
	return rst;
}

// 一定会更新，即使rangeStart == rangeEnd && newLen == 0
void CFakeWordMfcDlg::updateWordText(Word::_DocumentPtr doc, const std::wstring& txtFullText, size_t rangeStart, size_t rangeEnd, size_t newLen)
{
	try {
		// 修改输入的文本
		//Word::_DocumentPtr doc = wordCtrller->getActiveDocument().getData();
		_variant_t vFrom(rangeStart);
		_variant_t vEnd(rangeEnd);
		Word::RangePtr range = doc->Range(&vFrom, &vEnd);
		if (rangeStart < txtFullText.size())
		{
			if (rangeStart + newLen > txtFullText.size())
				newLen = txtFullText.size() - rangeStart;
			auto addText = txtFullText.substr(rangeStart, newLen);
			auto&& rangeText = range->Text;
			std::wstring oldText = rangeText.GetBSTR()  ? rangeText.GetBSTR() : L"";
			//range->Text = _bstr_t(addText.c_str());
			range->Delete();
			range->InsertAfter(addText.c_str());
			Word::SelectionPtr selection = doc->Application->Selection;
			if (selection) {
				_variant_t vSelEnd(Word::wdStory);
				selection->EndKey(&vSelEnd, &vtMissing);
				//selection->SetRange(from + length, from + length);
				//selection->SetRange(from + length, from + length);
			}
			//selection->Text =  _bstr_t(txtFullText.substr(0, from + length).c_str());
			//selection->InsertAfter(_bstr_t(txtFullText.substr(0, from + length).c_str()));
			//variant_t vCollapseEnd(Word::wdCollapseEnd);
			//selection->Collapse(&vCollapseEnd);
			std::cout << "Document text updating: Range: [" << rangeStart << "," << rangeEnd << ") "
				<< "Old text: " << wstr2str_2ANSI(oldText) << " -> "
				<< "Txt full string substr: " << wstr2str_2ANSI(addText)
				<< std::endl;
		}
		else
		{
			range->Text = L"";
		}
	}
	COM_CATCH_NORESULT(L"Word文本更新失败！")
	COM_CATCH_ALL_NORESULT(L"Word文本更新失败！")
}

Result<bool> CFakeWordMfcDlg::onWordDocumentBeforeClose(DISPPARAMS* params, WordEventSink::UserDataType userData)
{
	destroyStateBoxDlg();
	try {
		IDispatch* disp = params->rgvarg[0].pdispVal;
		CComPtr<IDispatch> docDisp(disp);
		DISPID dispid;
		HWND windowId = 0;
		OLECHAR* name = (OLECHAR*)L"Name";
		HRESULT hr = docDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_NAME_USER_DEFAULT, &dispid);
		if (FAILED(hr))
		{
			_com_error err(hr);
			return ErrorResult<bool>(hr, err.ErrorMessage());
		}
		VARIANT docNameVariant = { 0 };
		DISPPARAMS invokeDispParams = { 0 };
		hr = docDisp->Invoke(dispid, IID_NULL, LOCALE_NAME_USER_DEFAULT, DISPATCH_PROPERTYGET, &invokeDispParams, &docNameVariant, NULL, NULL);
		if (FAILED(hr))
		{
			_com_error err(hr);
			return ErrorResult<bool>(hr, err.ErrorMessage());
		}
		std::wstring docNameWStr = docNameVariant.bstrVal;
		std::wstringstream docNameStream;
		docNameStream << L"Document: " << docNameWStr << L" is being closed.";
		std::cout << wstr2str_2ANSI(docNameStream.str()) << std::endl;
		std::thread([msg = docNameStream.str()] {::MessageBox(0, msg.c_str(), L"Info", MB_ICONINFORMATION); }).detach();
	}
	catch (...)
	{
		return ErrorResult<bool>(0, L"Unknown error.");
	}
	return SuccessResult(true);
}

Result<bool> CFakeWordMfcDlg::onWordWindowActivate(DISPPARAMS* params, WordEventSink::UserDataType userData)
{
	if (isStateBoxDlgExist())
	{
		embeddedStateDlg->show();
		updateStateBoxDlg(embeddedStateDlg);
	}
	// 原型: HRESULT WindowActivate(struct _Document* Doc, struct Window* Wn);
	Result<bool> rst(false, E_POINTER, L"指针异常");
	auto paramCount = params->cArgs;
	if (paramCount < 2)
		return Result<bool>::CreateFromHRESULT(E_INVALIDARG, 0);
	IDispatch* documentDisp = params->rgvarg[0].pdispVal;
	IDispatch* windowDisp = params->rgvarg[1].pdispVal;
	if (!documentDisp)
		return rst;
	Word::_DocumentPtr docPtr;
	Word::WindowPtr windowPtr;
	HRESULT hr = documentDisp->QueryInterface(&docPtr);
	if (FAILED(hr))
	{
		_com_error err(hr);
		std::cout << "documentDisp->QueryInterface(&docPtr): 0x"
			<< std::hex << hr << std::dec
			<< wstr2str_2ANSI(err.ErrorMessage())
			<< std::endl;
		return ErrorResult<bool>(hr, err.ErrorMessage());
	}
	hr = windowDisp->QueryInterface(&windowPtr);
	if (FAILED(hr))
	{
		_com_error err(hr);
		std::cout << "windowDisp->QueryInterface(&windowPtr): 0x"
			<< std::hex << hr << std::dec
			<< wstr2str_2ANSI(err.ErrorMessage())
			<< std::endl;
		return ErrorResult<bool>(hr, err.ErrorMessage());
	}
	try {
		_bstr_t docNameBstr = docPtr->GetName();
		// 日志，消息框
		std::wstring docNameWStr;
		std::wstringstream docNameStream;
		docNameStream << L"Window is activated. ";
		if (docNameBstr.GetBSTR())
		{
			docNameWStr = docNameBstr;
			docNameStream << L"Document: " << docNameWStr << ".";
		}
		else
		{
			docNameStream << L"Error while getting document name.";
		}
		std::cout << wstr2str_2ANSI(docNameStream.str()) << std::endl;
		//std::thread([msg = docNameStream.str()] {::MessageBox(0, msg.c_str(), L"Info", MB_ICONINFORMATION); }).detach();

		// 获取窗口句柄
		HWND hwnd = reinterpret_cast<HWND>(windowPtr->GetHwnd());
		if (embeddedStateDlg && embeddedStateDlg->GetSafeHwnd())
			::SetParent(embeddedStateDlg->GetSafeHwnd(), hwnd);
	}
	COM_CATCH_NOMSGBOX(L"Word事件 WindowActivate 处理出错", rst)
	COM_CATCH_ALL_NOMSGBOX(L"Word事件 WindowActivate 处理出错", rst)
	return rst;
}

BOOL CFakeWordMfcDlg::enumWordChildWindowsProc(HWND hwnd, LPARAM params)
{
	int windowClassNameBufferLen = 256;
	WCHAR* windowClassName = new WCHAR[windowClassNameBufferLen];
	int windowTitleNameBufferLen = 256;
	WCHAR* windowTitleName = new WCHAR[windowTitleNameBufferLen];
	for (int rst = 0; (rst = ::GetClassName(hwnd, windowClassName, windowClassNameBufferLen)) != 0 && rst == windowClassNameBufferLen - 1;)
	{
		delete[] windowClassName;
		windowClassNameBufferLen += 256;
		windowClassName = new WCHAR[windowClassNameBufferLen];
	}
	for (int rst = 0; (rst = ::GetWindowText(hwnd, windowTitleName, windowTitleNameBufferLen)) != 0 && rst == windowTitleNameBufferLen - 1;)
	{
		delete[] windowTitleName;
		windowTitleNameBufferLen += 256;
		windowTitleName = new WCHAR[windowTitleNameBufferLen];
	}
	std::cout << "Enum Word Child Windows: \n\tWindow title: " << wstr2str_2ANSI(windowTitleName)
		<< "\n\t Window class: " << wstr2str_2ANSI(windowClassName)
		<< "\n\t Window HWND: 0x" << std::hex<< hwnd << std::dec 
		<< std::endl;
	// 如果是_WwB或_WwG，深入枚举其子窗口

	EnumChildWindows(hwnd, [](HWND childHwnd, LPARAM) -> BOOL {
		wchar_t childClassName[256];
		wchar_t childText[256];

		::GetClassName(childHwnd, childClassName, 256);
		::GetWindowText(childHwnd, childText, 256);

		std::wcout << L"  └─子窗口: 0x" << std::hex << (DWORD_PTR)childHwnd
			<< L" 类名: " << childClassName
			<< L" 标题: " << (childText[0] ? childText : L"(空)")
			<< std::dec << std::endl;
		return TRUE;
		}, 0);

	if (wcscmp(windowClassName, L"_WwG") == 0)
	{
		std::cout << "Find Word Document Window: 0x" << std::hex << hwnd << std::dec << std::endl;
		wordDocumentHwnd = hwnd;
		delete[] windowClassName;
		delete[] windowTitleName;
		return FALSE;
	}
	delete[] windowClassName;
	delete[] windowTitleName;
	return TRUE;
}

void CFakeWordMfcDlg::findWordDocumentHWND()
{
	// 遍历子窗口
	wordDocumentHwnd = NULL;
	EnumWindowsStruct enumWindowsStrtuct{
		this,
		0
	};
	EnumChildWindows(wordAppHwnd, [](HWND hwnd, LPARAM lParams) {
		return reinterpret_cast<EnumWindowsStruct*>(lParams)->ptr->enumWordChildWindowsProc(hwnd, lParams);
		}, reinterpret_cast<LPARAM>(&enumWindowsStrtuct));
	if (!wordDocumentHwnd)
	{
		MessageBeep(MB_ICONERROR);
		MessageBox(L"无法获取Word文档编辑框句柄，软件功能可能受影响", L"", MB_ICONERROR);
	}
}

// 带错误处理的获取编辑框long double类型的数值
long double CFakeWordMfcDlg::GetEditNumber(int id)
{
	CString numStr;
	GetDlgItem(id)->GetWindowText(numStr);
	long double num = 0;
	try {
		num = std::stold(numStr.GetString());
	}
	catch (...) {

	}
	return num;
}

// 任意支持类型(std::to_wstring(T))的编辑框数字文本设置
template <typename T>
void CFakeWordMfcDlg::SetEditNumber(int id, T&& number)
{
	GetDlgItem(id)->SetWindowText(std::to_wstring(std::forward<T>(number)).c_str());
}


void CFakeWordMfcDlg::OnOK()
{
	//CDialog::OnOK();
}

void CFakeWordMfcDlg::OnCancel()
{
	//CDialog::OnCancel();
}

void CFakeWordMfcDlg::OnClose()
{
	// TODO: 在此添加消息处理程序代码和/或调用默认值
	UnregisterHotKey(GetSafeHwnd(), ID_HOTKEY_DEFAULT_SWITCH_REPLACE_MODE); // 取消快捷键
	m_pToolTip->Activate(FALSE);
	auto toolTip = m_pToolTip;
	m_pToolTip = nullptr;
	toolTip->DestroyToolTipCtrl();
	destroyStateBoxDlg();
	delete wordCtrller;
	CDialogEx::OnCancel(); // 这里关闭对话框
	CDialogEx::OnClose();
}


void CFakeWordMfcDlg::OnHotKey(UINT nHotKeyId, UINT nKey1, UINT nKey2)
{
	// TODO: 在此添加消息处理程序代码和/或调用默认值
	switch (nHotKeyId)
	{
	case ID_HOTKEY_DEFAULT_SWITCH_REPLACE_MODE:
	{
		this->replaceMode.next();
		updateStateBoxDlg(embeddedStateDlg);
		break;
	}
	default:
		break;
	}
	CDialogEx::OnHotKey(nHotKeyId, nKey1, nKey2);
}


BOOL CFakeWordMfcDlg::PreTranslateMessage(MSG* pMsg)
{
	// TODO: 在此添加专用代码和/或调用基类
	if (m_pToolTip && m_pToolTip->GetSafeHwnd() != NULL)
	{
		m_pToolTip->RelayEvent(pMsg);
	}
	return CDialogEx::PreTranslateMessage(pMsg);
}


void CFakeWordMfcDlg::OnClickedBtnSetting()
{
	// TODO: 在此添加控件通知处理程序代码
	CSettingDialog dlg(switchReplaceModeHotkey, wordDetectMode, oiRatio, this);
	INT_PTR rst = dlg.DoModal();
	if (rst == IDOK)
	{
		HotKey hk = dlg.getSwitchReplaceModeHotKey();
		auto OIRatio = dlg.getOIRatio();
		WordDetectMode detectMode = dlg.getWordDetectMode();
		switchReplaceModeHotkey = hk;
		wordDetectMode = detectMode;
		oiRatio = OIRatio;
		// 保存到配置文件
		configManager.writeConfigNumber(configSection, configKeys[0], static_cast<int>(wordDetectMode));
		configManager.writeConfigNumber(configSection, configKeys[1], oiRatio);
		configManager.writeConfigNumber(configSection, configKeys[2], switchReplaceModeHotkey.virtualKeyCode);
		configManager.writeConfigNumber(configSection, configKeys[3], switchReplaceModeHotkey.modifiers);
		// 更新热键
		UnregisterHotKey(GetSafeHwnd(), ID_HOTKEY_DEFAULT_SWITCH_REPLACE_MODE);
		RegisterHotKey(GetSafeHwnd(), ID_HOTKEY_DEFAULT_SWITCH_REPLACE_MODE, switchReplaceModeHotkey.modifiers, switchReplaceModeHotkey.virtualKeyCode);
		// 更新状态显示窗口的状态
		updateStateBoxDlg(embeddedStateDlg);
	}
}
