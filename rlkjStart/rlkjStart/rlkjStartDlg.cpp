
// rlkjStartDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "rlkjStart.h"
#include "rlkjStartDlg.h"
#include "afxdialogex.h"
#include "ZSPDStartupManager.h"
#include <windows.h>  
#include <Winsvc.h> 
#include"tlhelp32.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CrlkjStartDlg 对话框



CrlkjStartDlg::CrlkjStartDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CrlkjStartDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CrlkjStartDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CrlkjStartDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDOK3, &CrlkjStartDlg::OnBnClickedOk3)
	ON_BN_CLICKED(IDOK, &CrlkjStartDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CrlkjStartDlg 消息处理程序

BOOL CrlkjStartDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO:  在此添加额外的初始化代码
	m_imagePre.LoadFromResource(AfxGetResourceHandle(), MAKEINTRESOURCE(IDB_BITMAPBACKGROUND));
	GetClientRect(&m_rect);
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CrlkjStartDlg::OnPaint()
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
		CPaintDC dc(this);
		m_imagePre.Draw(dc.m_hDC, 0, 0, m_rect.Width(), m_rect.Height());
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CrlkjStartDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CrlkjStartDlg::OnBnClickedOk3()
{
	// TODO:  在此添加控件通知处理程序代码
	if (IsOpen(_T("acad.exe")))
	{
		AfxMessageBox(_T("正在使用CAD，请先关闭！"));
		return;
	}

	TCHAR szFilePath[MAX_PATH + 1] = { 0 };
	GetModuleFileName(NULL, szFilePath, MAX_PATH);
	(_tcsrchr(szFilePath, _T('\\')))[1] = 0; 
	
	ZSPDStartupManager regMgr;
	regMgr.SetACADVersion(_T("R19.0"));
	regMgr.SetModulePath(szFilePath);

	regMgr.Startup();
	OnCancel();
}


void CrlkjStartDlg::OnBnClickedOk()
{
	// TODO:  在此添加控件通知处理程序代码
	if (IsOpen(_T("acad.exe")))
	{
		AfxMessageBox(_T("正在使用CAD，请先关闭！"));
		return;
	}

	TCHAR szFilePath[MAX_PATH + 1] = { 0 };
	GetModuleFileName(NULL, szFilePath, MAX_PATH);
	(_tcsrchr(szFilePath, _T('\\')))[1] = 0; 
	ZSPDStartupManager regMgr;
	regMgr.SetACADVersion(_T("R19.0"));
	regMgr.SetModulePath(szFilePath);

	regMgr.clear();
	CDialogEx::OnOK();
}

BOOL CrlkjStartDlg::IsOpen(TCHAR* strProcess)
{
	PROCESSENTRY32 pe32;
	pe32.dwSize = sizeof(pe32);
	HANDLE hProcessSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
	if (hProcessSnap == INVALID_HANDLE_VALUE)
	{
		printf("CreatToolhelp32Snapeshot调用失败!\n");
		return FALSE;
	}

	BOOL bMore = Process32First(hProcessSnap, &pe32);//获得第一个进程的句柄。注意是第一个所以还要while()


	while (bMore)
	{
		bMore = Process32Next(hProcessSnap, &pe32);//这里要注意，必须循环，因为你只运行一次函数，只能取到进程


		CString strFindFile;
		strFindFile.Format(_T("%s"), strProcess);
		CString strExeFile;
		strExeFile.Format(_T("%s"), pe32.szExeFile);

		if (strExeFile.CompareNoCase(strFindFile) == 0)
		{
			CloseHandle(hProcessSnap);
			return TRUE;
		}
	}
	CloseHandle(hProcessSnap);
	return FALSE;
}