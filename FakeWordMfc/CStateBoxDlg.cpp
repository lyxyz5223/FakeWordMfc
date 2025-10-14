// CStateBoxDlg.cpp: 实现文件
//

#include "pch.h"
#include "FakeWordMfc.h"
#include "afxdialogex.h"
#include "CStateBoxDlg.h"


// CStateBoxDlg 对话框

IMPLEMENT_DYNAMIC(CStateBoxDlg, CDialogEx)

CStateBoxDlg::CStateBoxDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DIALOG_STATEBOX, pParent)
{

}

CStateBoxDlg::~CStateBoxDlg()
{
}

void CStateBoxDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_HOTKEY_REPLACEMODE, m_HotKeyReplaceMode);
	DDX_Control(pDX, IDC_FILENAME, m_FileName);
	DDX_Control(pDX, IDC_STATE_REPLACEMODE, m_StateReplaceMode);
}


BEGIN_MESSAGE_MAP(CStateBoxDlg, CDialogEx)
	ON_WM_DESTROY()
	ON_WM_CLOSE()
END_MESSAGE_MAP()


// CStateBoxDlg 消息处理程序


void CStateBoxDlg::OnDestroy()
{
	CDialogEx::OnDestroy();

	// TODO: 在此处添加消息处理程序代码
	delete this;
}


void CStateBoxDlg::OnClose()
{
	// TODO: 在此添加消息处理程序代码和/或调用默认值

	CDialogEx::OnClose();
}
