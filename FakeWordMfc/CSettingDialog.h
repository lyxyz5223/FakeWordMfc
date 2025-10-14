#pragma once
#include "afxdialogex.h"
#include <string>
#include "IniConfigManager.h"
#include "AppTypes.h"
#include <iostream>

// CSettingDialog 对话框

class CSettingDialog : public CDialogEx
{
	DECLARE_DYNAMIC(CSettingDialog)

private:
	// 带错误处理的获取编辑框long double类型的数值
	long double GetEditNumber(CWnd& edit) const {
		CString numStr;
		edit.GetWindowText(numStr);
		long double num = 0;
		try {
			num = std::stold(numStr.GetString());
		}
		catch (...) {

		}
		return num;
	}
	long double GetEditNumber(int id) const {
		return GetEditNumber(*GetDlgItem(id));
	}

	// 任意支持类型(std::to_wstring(T))的编辑框数字文本设置
	template <typename T>
	void SetEditNumber(CWnd& edit, T&& number) {
		edit.SetWindowText(std::to_wstring(std::forward<T>(number)).c_str());
	}
	template <typename T>
	void SetEditNumber(int id, T&& number) {
		SetEditNumber(*GetDlgItem(id), number);
	}


public:
	HotKey getSwitchReplaceModeHotKey() const {
		return m_switchReplaceModeHotKey;
	}

	WordDetectMode getWordDetectMode() const {
		return m_wordDetectMode;
	}

	long double getOIRatio() const {
		return m_oiRatio;
	}

	// 只有窗口存在时能够使用
	HotKey getSwitchReplaceModeHotKeyByUI() const {
		// 虚拟密钥代码采用低序字节，修饰符标志位于高阶字节中
		HotKey hk = { 0, 0 };
		auto dhk = m_ReplaceModeHotKey.GetHotKey();
		hk = HotKey::fromCtrlHotKey(dhk);
		//std::cout << std::hex << dhk << " " << hk.modifiers << " " << (char)hk.virtualKeyCode << std::dec << std::endl;
		return hk;
	}
	// 只有窗口存在时能够使用
	WordDetectMode getWordDetectModeByUI() const {
		return WordDetectMode::fromNumber(m_DetectModeComboBox.GetCurSel());
	}
	// 只有窗口存在时能够使用
	long double getOIRatioByUI() const {
		return GetEditNumber(IDC_EDIT_OIRATIO);
	}

private:
	HotKey m_switchReplaceModeHotKey = DEFAULT_HOTKEY_SWITCHREPLACEMODE;
	WordDetectMode m_wordDetectMode = DEFAULT_WORDDETECTMODE;
	long double m_oiRatio = DEFAULT_OIRATIO;
	CSize m_windowMinSize;

public:
	CSettingDialog(HotKey switchReplaceModeHotKey, WordDetectMode wordDetectMode, long double oiRatio,
		CWnd* pParent = nullptr);   // 标准构造函数
	virtual ~CSettingDialog();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG_SETTING };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	CHotKeyCtrl m_ReplaceModeHotKey;
	CComboBox m_DetectModeComboBox;
	CEdit m_OIRatioEdit;
	afx_msg void OnDeltaposSpinOiratio(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnChangeEditOIRatio();
	afx_msg void OnChangeHotKeyReplaceMode();
	virtual BOOL OnInitDialog();
	afx_msg void OnSelChangeComboDetectMode();
	afx_msg void OnGetMinMaxInfo(MINMAXINFO* lpMMI);
};
