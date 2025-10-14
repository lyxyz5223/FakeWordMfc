// CSettingDialog.cpp: 实现文件
//

#include "pch.h"
#include "FakeWordMfc.h"
#include "afxdialogex.h"
#include "CSettingDialog.h"


// CSettingDialog 对话框

IMPLEMENT_DYNAMIC(CSettingDialog, CDialogEx)

CSettingDialog::CSettingDialog(HotKey switchReplaceModeHotKey, WordDetectMode wordDetectMode, long double oiRatio,
	CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DIALOG_SETTING, pParent)
	, m_switchReplaceModeHotKey(switchReplaceModeHotKey)
	, m_wordDetectMode(wordDetectMode)
	, m_oiRatio(oiRatio)
{

}

CSettingDialog::~CSettingDialog()
{
}

void CSettingDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_HOTKEY_REPLACEMODE, m_ReplaceModeHotKey);
	DDX_Control(pDX, IDC_COMBO_DETECTMODE, m_DetectModeComboBox);
	DDX_Control(pDX, IDC_EDIT_OIRATIO, m_OIRatioEdit);
}


BEGIN_MESSAGE_MAP(CSettingDialog, CDialogEx)
	ON_NOTIFY(UDN_DELTAPOS, IDC_SPIN_OIRATIO, &CSettingDialog::OnDeltaposSpinOiratio)
	ON_EN_CHANGE(IDC_EDIT_OIRATIO, &CSettingDialog::OnChangeEditOIRatio)
	ON_CBN_SELCHANGE(IDC_COMBO_DETECTMODE, &CSettingDialog::OnSelChangeComboDetectMode)
	ON_EN_CHANGE(IDC_HOTKEY_REPLACEMODE, &CSettingDialog::OnChangeHotKeyReplaceMode)
	ON_WM_GETMINMAXINFO()
END_MESSAGE_MAP()


// CSettingDialog 消息处理程序

// 会调用OnChangeEditOIRatio
void CSettingDialog::OnDeltaposSpinOiratio(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMUPDOWN pNMUpDown = reinterpret_cast<LPNMUPDOWN>(pNMHDR);
	*pResult = 0;
	auto num = GetEditNumber(m_OIRatioEdit);
	num -= pNMUpDown->iDelta;
	SetEditNumber(m_OIRatioEdit, num);
}


void CSettingDialog::OnChangeEditOIRatio()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	CString numStr;
	m_OIRatioEdit.GetWindowText(numStr);
	long double num = 0;
	try {
		num = std::stold(numStr.GetString());
		m_oiRatio = num;
	}
	catch (...) {
		SetEditNumber(m_OIRatioEdit, num);
		m_oiRatio = num;
	}
	if (num < 0)
	{
		SetEditNumber(m_OIRatioEdit, 0);
		m_oiRatio = 0;
	}

}

void CSettingDialog::OnChangeHotKeyReplaceMode()
{
	m_switchReplaceModeHotKey = getSwitchReplaceModeHotKeyByUI();
}


BOOL CSettingDialog::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	CtrlHotKey chk = m_switchReplaceModeHotKey.toCtrlHotKey();
	m_ReplaceModeHotKey.SetHotKey(chk.virtualKeyCode, chk.modifiers);
	m_DetectModeComboBox.SetCurSel(static_cast<int>(m_wordDetectMode));
	SetEditNumber(m_OIRatioEdit, m_oiRatio);

	// 设置热键规则
	m_ReplaceModeHotKey.SetRules(HKCOMB_NONE, HOTKEYF_CONTROL | HOTKEYF_ALT);

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}


void CSettingDialog::OnSelChangeComboDetectMode()
{
	m_wordDetectMode = getWordDetectModeByUI();
}


void CSettingDialog::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
	// TODO: 在此添加消息处理程序代码和/或调用默认值

	CDialogEx::OnGetMinMaxInfo(lpMMI);
	if (m_windowMinSize.cx == 0 && m_windowMinSize.cy == 0)
	{
		CRect windowRect;
		GetWindowRect(&windowRect);
		m_windowMinSize.SetSize(windowRect.Width(), windowRect.Height());
	}
	lpMMI->ptMinTrackSize.x = m_windowMinSize.cx;
	lpMMI->ptMinTrackSize.y = m_windowMinSize.cy;
}
