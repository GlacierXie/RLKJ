
// rlkjStartDlg.cpp : ʵ���ļ�
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


// CrlkjStartDlg �Ի���



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


// CrlkjStartDlg ��Ϣ�������

BOOL CrlkjStartDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO:  �ڴ���Ӷ���ĳ�ʼ������
	m_imagePre.LoadFromResource(AfxGetResourceHandle(), MAKEINTRESOURCE(IDB_BITMAPBACKGROUND));
	GetClientRect(&m_rect);
	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CrlkjStartDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CPaintDC dc(this);
		m_imagePre.Draw(dc.m_hDC, 0, 0, m_rect.Width(), m_rect.Height());
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CrlkjStartDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CrlkjStartDlg::OnBnClickedOk3()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	if (IsOpen(_T("acad.exe")))
	{
		AfxMessageBox(_T("����ʹ��CAD�����ȹرգ�"));
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
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	if (IsOpen(_T("acad.exe")))
	{
		AfxMessageBox(_T("����ʹ��CAD�����ȹرգ�"));
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
		printf("CreatToolhelp32Snapeshot����ʧ��!\n");
		return FALSE;
	}

	BOOL bMore = Process32First(hProcessSnap, &pe32);//��õ�һ�����̵ľ����ע���ǵ�һ�����Ի�Ҫwhile()


	while (bMore)
	{
		bMore = Process32Next(hProcessSnap, &pe32);//����Ҫע�⣬����ѭ������Ϊ��ֻ����һ�κ�����ֻ��ȡ������


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