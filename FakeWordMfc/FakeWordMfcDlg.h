
// FakeWordMfcDlg.h: 头文件
//

#pragma once
#include "WordController.h"
#include "TextFileReader.h"
#include "IniConfigManager.h"
#include <thread>
#include <vector>
#include "Apptypes.h"
#include "CStateBoxDlg.h"
#include <random>
#include <unordered_set>
#undef max
#undef min

// CFakeWordMfcDlg 对话框
class CFakeWordMfcDlg : public CDialogEx
{
private:
	// 当前对象正在析构
	bool objDeconstructing = false;

	// 配置文件管理
	IniConfigManager configManager{L"config.ini", false};
	std::wstring configSection = L"Config";
	std::vector<std::wstring> configKeys{{
		L"WordDetectMode",
		L"OIRatio",
		L"SwitchReplaceModeHotkey_VirtualKeyCode",
		L"SwitchReplaceModeHotkey_Modifiers"
	}};

	CSize windowMinSize;
	// Word管理
	WordController* wordCtrller = nullptr;
	// 当前工作模式
	ReplaceMode replaceMode = ReplaceMode::SlackOff;

	// -----------------设置---------------------
	// 切换模式的快捷键
	HotKey switchReplaceModeHotkey{ DEFAULT_HOTKEY_SWITCHREPLACEMODE };
	// 输出输入比例
	long double oiRatio = DEFAULT_OIRATIO;
	// Word改变检测模式
	WordDetectMode wordDetectMode = DEFAULT_WORDDETECTMODE;
	// 小说和工作文件
	SupportedFile selectedFishFile;
	SupportedFile selectedWorkFile;
	TextFileReader fishFileTxtReader;
	std::wstring processedFishFileTxtWordText; // 为与word比较特地处理后的文本
	TextFileReader workFileTxtReader;
	std::wstring processedWorkFileTxtWordText; // 为与word比较特地处理后的文本

	std::wstring getCurrentReplacingFileName() const {
		std::wstring workingFileName;
		auto replaceMode = this->replaceMode;
		switch (replaceMode)
		{
		default:
		case ReplaceMode::SlackOff:
		{
			workingFileName = selectedFishFile.fileName;
			break;
		}
		case ReplaceMode::Work:
		{
			workingFileName = selectedWorkFile.fileName;
			break;
		}
		}
		return workingFileName;
	}

	void updateStateBoxDlg(const std::unique_ptr<CStateBoxDlg>& dlg) {
		if (dlg)
			dlg->updateState(replaceMode, getCurrentReplacingFileName(), switchReplaceModeHotkey);
	}
	void newStateBoxDlg(HWND target) {
		auto parentCwnd = CWnd::FromHandle(target);
		std::unique_ptr<CStateBoxDlg> stateBox = std::make_unique<CStateBoxDlg>();
		auto& dlg = embeddedStateDlg = std::move(stateBox);
		dlg->Create(0);
		//::SetParent(dlg->GetSafeHwnd(), target);
		dlg->SetParent(parentCwnd);
		CRect windowRect;
		dlg->GetWindowRect(windowRect);
		windowRect.MoveToXY(0, 0);
		dlg->MoveWindow(windowRect);
		//::MoveWindow(dlg->GetSafeHwnd(), windowRect.left, windowRect.top, windowRect.Width(), windowRect.Height(), TRUE);
		dlg->show();
	}
	bool isStateBoxDlgExist() {
		return embeddedStateDlg && embeddedStateDlg->GetSafeHwnd() && IsWindow(embeddedStateDlg->GetSafeHwnd());
	}
	void destroyStateBoxDlg() {
		embeddedStateDlg.release();
	}
	// 嵌入Word的状态显示窗口，每个Word独一份
	// 映射中的键，唯一id
	typedef unsigned long long IdType;
	typedef std::numeric_limits<IdType> IdLimits;
	std::unordered_set<IdType> wordWindowIdSet;
	std::random_device rdWordWindowId;
	std::mt19937_64 mtWordWindowId{ rdWordWindowId() };
	std::unique_ptr<CStateBoxDlg> embeddedStateDlg;
	std::mutex mtxEmbeddedStateDlgMap; // 互斥锁
	IdType getRandomId(std::random_device& rd, std::mt19937_64& mt) {
		// 均匀分布
		std::uniform_int_distribution<IdType> uniIntDist(0, IdLimits::max()); // [min, max]
		return uniIntDist(mt);
	}
	IdType getAvailableWordWindowId() {
		IdType id;
		while (wordWindowIdSet.count(id = getRandomId(rdWordWindowId, mtWordWindowId)))
			;
		return id;
	}
	std::unordered_set<IdType> injectIdSet;
	std::random_device rdInjectIdSet;
	std::mt19937_64 mtInjectIdSet{ rdInjectIdSet() };
	IdType getAvailableInjectId() {
		IdType id;
		while (injectIdSet.count(id = getRandomId(rdInjectIdSet, mtInjectIdSet)))
			;
		return id;
	}

	// 更新填充小说文本的时候要截取的小说字符串范围
	struct SubTextRange {
		unsigned long long start = 0;
		unsigned long long end = 0;
		unsigned long long newLength = 0;
	} wordTextUpdateTxtSubTextRange; // 仅有一个位置使用，可以作为函数的局部变量
	// word
	std::wstring wordLastText = L"\r"; // word 上一次的文本内容，用于检测文本变化
	std::thread wordTextChangeDetecterThread; // 检测word文本变化的线程
	void wordTextChangeDetecterWorker(); // 检测文本变化的线程工作函数
	/**
	 * 收到word的WindowSelectionChange事件时候触发，在打开Word时进行事件的注册
	 */
	Result<bool> onWordWindowSelectionChangeEvent(DISPPARAMS* params, WordEventSink::UserDataType userData);

	Result<bool> processTextChange();

	/**
	 * 更新word文本的函数
	 * 
	 */
	void updateWordText(Word::_DocumentPtr doc, const std::wstring& txtFullText, size_t rangeStart, size_t rangeEnd, size_t newLen);
	/**
	 * 收到word的文档关闭事件时触发，在打开Word时进行事件的注册
	 */
	Result<bool> onWordDocumentBeforeClose(DISPPARAMS* params, WordEventSink::UserDataType userData);
	/**
	 * 收到word的窗口激活时触发，在打开Word时进行事件的注册
	 * 原型 HRESULT WindowActivate(struct _Document* Doc, struct Window* Wn);
	 */
	Result<bool> onWordWindowActivate(DISPPARAMS* params, WordEventSink::UserDataType userData);


	/**
	 * 遍历窗口，获取word文档区域句柄
	 */
	HWND wordAppHwnd = NULL;
	HWND wordDocumentHwnd = NULL;
	BOOL enumWordChildWindowsProc(HWND hwnd, LPARAM params);
	//LRESULT hookWordDocumentProc(int code, WPARAM wParam, LPARAM lParam);
	void findWordDocumentHWND();

private:
	// Mfc
	
	long double GetEditNumber(int id);
	template <typename T>
	void SetEditNumber(int id, T&& number);



// 构造
public:
	~CFakeWordMfcDlg() {
		objDeconstructing = true;
		if (wordTextChangeDetecterThread.joinable())
			wordTextChangeDetecterThread.join();
        if (wordCtrller)
		    delete wordCtrller;
	}
	CFakeWordMfcDlg(CWnd* pParent = nullptr);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_FAKEWORDMFC_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	// 提示
	CMFCToolTipCtrl* m_pToolTip = nullptr;
	CButton m_BtnOpenFile;
	CEdit m_EditFishFilePath;
	CEdit m_EditWorkFilePath;
	CButton m_BtnSelectFishFile;
	CButton m_BtnSelectWorkFile;
	afx_msg void OnBnClickedBtnOpenFile();
	afx_msg void OnBnClickedBtnSelectFishFile();
	afx_msg void OnBnClickedBtnSelectWorkFile();
	afx_msg void OnBnClickedBtnOpenWord();
	afx_msg void OnGetMinMaxInfo(MINMAXINFO* lpMMI);
	virtual void OnOK();
	virtual void OnCancel();
	afx_msg void OnClose();
	afx_msg void OnHotKey(UINT nHotKeyId, UINT nKey1, UINT nKey2);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnClickedBtnSetting();
};

