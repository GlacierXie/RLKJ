
// rlkjStartDlg.h : ͷ�ļ�
//

#pragma once


// CrlkjStartDlg �Ի���
class CrlkjStartDlg : public CDialogEx
{
// ����
public:
	CrlkjStartDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_RLKJSTART_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;
	CImage    m_imagePre;        //Ԥ��ͼƬ
	CRect     m_rect;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk3();
	afx_msg void OnBnClickedOk();

private:
	BOOL IsOpen(TCHAR* strProcess);
};
