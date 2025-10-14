#pragma once
#include "afxdialogex.h"
#include <string>
#include "AppTypes.h"

// CStateBoxDlg 对话框

class CStateBoxDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CStateBoxDlg)

public:
	void show() {
		this->ShowWindow(SW_SHOW);
	}
	void hide() {
		this->ShowWindow(SW_HIDE);
	}
	void showMinimized() {
		this->ShowWindow(SW_MINIMIZE); // 最小化指定的窗口，并按 Z 顺序激活下一个顶级窗口。
		//this->ShowWindow(SW_SHOWMINIMIZED); // 激活窗口并将其显示为最小化窗口。
	}
	void showMaximized() {
		this->ShowWindow(SW_MAXIMIZE); // 激活窗口并显示最大化的窗口。
	}
	BOOL Create(CWnd* pParentWnd = NULL)
	{
		return CDialogEx::Create(this->m_lpszTemplateName, pParentWnd);
	}
	
	void setReplaceModeName(const std::wstring& mode) {
		m_StateReplaceMode.SetWindowText(mode.c_str());
	}

	void setFileName(const std::wstring& fileName) {
		m_FileName.SetWindowText(fileName.c_str());
	}

	void setSwitchReplaceModeHoyKey(const HotKey& hotKey) {
		auto hk = hotKey.toCtrlHotKey();
		m_HotKeyReplaceMode.SetHotKey(hk.virtualKeyCode, hk.modifiers);
	}

	void updateState(const std::wstring& mode, const std::wstring& fileName, const HotKey& hotKey) {
		setReplaceModeName(mode);
		setFileName(fileName);
		setSwitchReplaceModeHoyKey(hotKey);
	}

public:
	CStateBoxDlg(CWnd* pParent = nullptr);   // 标准构造函数
	virtual ~CStateBoxDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG_STATEBOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	CHotKeyCtrl m_HotKeyReplaceMode;
	CStatic m_FileName;
	CStatic m_StateReplaceMode;
	afx_msg void OnDestroy();
	afx_msg void OnClose();
};
