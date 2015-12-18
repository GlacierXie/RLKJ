

#include "stdafx.h"
#include "RLOperExcel.h"
#include <TlHelp32.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CRLOperExcel::CRLOperExcel()
	:m_bIsXlsProcRedu(FALSE)
{
	if (!AfxOleInit())
	{
		AfxMessageBox(_T("��ʼ��ole��ʧ��!"));
		return ;
	}
}

CRLOperExcel::~CRLOperExcel()
{
	/*�ͷ���Դ*/
	try
	{    //δ��ʼ���Ļ�Excel App  �ر��Լ������쳣
		XlsCloseAlert(); //�ر���ʾ��Ϣ!
		try
		{
			m_xlsWorkBooks.Close();  //�˳�����
		}
		catch(...)
		{
			XlsTerminateExcelProcess();
		}
		if (m_xlsWorkBooks.m_lpDispatch != NULL)
		{
			m_xlsWorkBooks.ReleaseDispatch();
		}	
		try
		{
			m_xlsAppLication.Quit();
		}
		catch(...)
		{
			XlsTerminateExcelProcess();
		}        
		m_xlsAppLication.ReleaseDispatch();
	}
	catch (_com_error &e)
	{
		XlsErrInfo(e);
		XlsTerminateExcelProcess();
	}
	AfxOleTerm(FALSE);//�ر�ole�� 
}

//�ӱ��ģ�崴��Excel�ļ�
BOOL CRLOperExcel::CreateXlsFromXltFile(const CString &strXltFileName, BOOL bIsShowExcel)
{
	static int iAppCout = 1;  //Ӧ�ó��������
	if (iAppCout > 1)
	{
		AfxMessageBox(_T("�������ظ���ʼ��!"));
		return FALSE;
	}
	if (m_bIsXlsProcRedu)
	{
		XlsQuit();//�ж��Ƿ���EXCEL����
	}
	if (!m_xlsAppLication.CreateDispatch(_T("Excel.Application"), NULL))
	{
		AfxMessageBox(_T("����Excel������ʧ��!"));
		return FALSE;
	}
	if (strXltFileName.IsEmpty())
	{
		return FALSE;
	}
	m_xlsAppLication.put_UserControl(TRUE);
	XlsOpen(strXltFileName,_T(""));
	m_xlsAppLication.put_Visible(bIsShowExcel);
	iAppCout++;
	return TRUE;
}

//��������ʼ��Excel
int CRLOperExcel::CreateXls(BOOL bIsShowExcel)
{
	try
	{
		if (m_bIsXlsProcRedu)
		{
			XlsQuit();//�ж��Ƿ���EXCEL����
		}
		if (!m_xlsAppLication.CreateDispatch(_T("Excel.Application"), NULL))
		{
			AfxMessageBox(_T("����Excel������ʧ��!"));
			return -1;
		}
		/*�жϵ�ǰExcel�İ汾*/
// 		CString &strExcelVersion = m_xlsAppLication.get_Version();
// 		int iStart = 0;
// 		strExcelVersion = strExcelVersion.Tokenize(_T("."), iStart);
// 		if (_T("11") == strExcelVersion)
// 		{
// 			AfxMessageBox(_T("��ǰExcel�İ汾��2003��"));
// 		}
// 		else if (_T("12") == strExcelVersion)
// 		{
// 			AfxMessageBox(_T("��ǰExcel�İ汾��2007��"));
// 		}
// 		else
// 		{
// 			AfxMessageBox(_T("��ǰExcel�İ汾�������汾��"));
// 		}

		m_xlsAppLication.put_UserControl(FALSE);
		//XlsOpen(strXltFileName,_T(""));
		m_xlsAppLication.put_Visible(bIsShowExcel);

	}
	catch (_com_error &e)
	{
		XlsErrInfo(e);
	}

	return 1;

}

// ��һ�鵥Ԫ���в�������
void CRLOperExcel::XlsSetRangeData(const CString &Row, const CString &Col, const CString &strValue)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (Row.IsEmpty() || Col.IsEmpty())
	{
		return;
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet    xlsSheet;
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp);
		lpDisp = xlsSheet.get_Range((_variant_t)Row, (_variant_t)Col);
	}
	else
	{
		return;
	}
	CRange  range;
	if (lpDisp != NULL)
	{
		range.AttachDispatch(lpDisp);
		range.put_Value2((_variant_t)strValue);
	}
	range.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	lpDisp = NULL;
}

//���һ�鵥Ԫ���е�����   jd ������
CString CRLOperExcel::XlsGetRangeStrData(const CString &Row, const CString &Col)
{
	if (!xlsAppIsInit())
	{
		return _T("erro");
	}
	if (Row.IsEmpty() || Col.IsEmpty())
	{
		return _T("erro");
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet    xlsSheet;
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp);
		lpDisp = xlsSheet.get_Range((_variant_t)Row, (_variant_t)Col);
	}
	else
	{
		return _T("erro");
	}
	CRange  range;
	CString strTmp = _T("");
	if (lpDisp != NULL)
	{
		range.AttachDispatch(lpDisp);
		COleVariant vResult =range.get_Value2();

		// 		if (vResult.vt == VT_BSTR)
		// 		{
		// 			strTmp = vResult.bstrVal;  //�ַ���
		// 		}
	}
	xlsSheet.ReleaseDispatch();
	range.ReleaseDispatch();
	return strTmp;
}

double CRLOperExcel::XlsGetRangeDoubleData(const CString &Row, const CString &Col)
{
	if (!xlsAppIsInit())
	{
		return -1;
	}
	if (Row.IsEmpty() || Col.IsEmpty())
	{
		return -1.0;  //
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet    xlsSheet;
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp);
		lpDisp = xlsSheet.get_Range((_variant_t)Row, (_variant_t)Col);
	}
	else
	{
		return -1.0;
	}

	CRange  range;
	double dResult = 0.0;
	if (lpDisp != NULL)
	{
		range.AttachDispatch(lpDisp);
		COleVariant vResult =range.get_Value2();
		// 		COleSafeArray oleSafeRead(vResult);
		// 		if (vResult.vt == VT_R8)
		// 		{
		// 			dResult = vResult.dblVal;  //�ַ���
		// 		}

	}
	xlsSheet.ReleaseDispatch();
	range.ReleaseDispatch();
	return dResult;
}

//��һ�鵥Ԫ���в�������
void CRLOperExcel::XlsSetRangeData(const CString &Row, const CString &Col, double dValue)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (Row.IsEmpty() || Col.IsEmpty())
	{
		return;
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet  xlsSheet;
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp);
		lpDisp = xlsSheet.get_Range((_variant_t)Row, (_variant_t)Col);
	}
	else
	{
		return;
	}
	CRange  range;
	if (lpDisp != NULL)
	{
		range.AttachDispatch(lpDisp);
		range.put_Value2((_variant_t)dValue);
	}	
	range.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	lpDisp = NULL;
}

//�����ڱ߿�
void CRLOperExcel::XlsDrawInBorders(const CString &Row, const CString &Col,XlBordersLineStyle xlLineStyle, XlBorderWeight XlsWeight,long coloIndex)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (Row.IsEmpty() || Col.IsEmpty())
	{
		return;
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet    xlsSheet;
	if (lpDisp == NULL)
	{
		return;
	}
	xlsSheet.AttachDispatch(lpDisp);
	if (xlsSheet.m_lpDispatch == NULL)
	{
		return;
	}
	lpDisp = xlsSheet.get_Range(_variant_t(Row), _variant_t(Col));
	CRange  rang; 
	rang.AttachDispatch(lpDisp);
	CBorders xlsBords;   
	xlsBords.AttachDispatch(rang.get_Borders());
	CBorder  xlsBord;   //�ڱ߿�
	for (int i = 11; i<= 12; i++)
	{
        /*
		xlInsideHorizontal = 12,
		xlInsideVertical = 11,
		*/
		lpDisp = xlsBords.get_Item(_variant_t((long)i));   
		if (lpDisp == NULL)
		{
			return;
		}
		xlsBord.AttachDispatch(lpDisp);
		xlsBord.put_Weight(_variant_t((long)XlsWeight));
		xlsBord.put_ColorIndex(_variant_t(coloIndex));
		xlsBord.put_LineStyle(_variant_t(xlLineStyle));
	}
	xlsBord.ReleaseDispatch();
	xlsBords.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	rang.ReleaseDispatch(); 
	lpDisp = NULL;
}

//������߿�
void CRLOperExcel::XlsDrawOutBorders(const CString &Row, const CString &Col,XlBordersLineStyle xlLineStyle, XlBorderWeight XlsWeight,long coloIndex)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (Row.IsEmpty() || Col.IsEmpty())
	{
		return;
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet    xlsSheet;
	if (lpDisp == NULL)
	{
		return;
	}
	xlsSheet.AttachDispatch(lpDisp);
	if (xlsSheet.m_lpDispatch == NULL)
	{
		return;
	}
	lpDisp = xlsSheet.get_Range(_variant_t(Row), _variant_t(Col));
	if (lpDisp == NULL)
	{
		return;
	}
	CRange  rang; 
	rang.AttachDispatch(lpDisp);
	lpDisp = rang.get_Borders();
	if (lpDisp == NULL)
	{
		return;
	}
	CBorders xlsBords;   //��
	xlsBords.AttachDispatch(lpDisp);
	//�߿�
	CBorder xlsBordTmp;

	//����Left�߿�
	lpDisp = xlsBords.get_Item(_variant_t((long)xlEdgeLeft));
	if (lpDisp == NULL)
	{
		return;
	}
 	xlsBordTmp.AttachDispatch(lpDisp);  
 	xlsBordTmp.put_ColorIndex(_variant_t(coloIndex));  
 	xlsBordTmp.put_Weight(_variant_t((long)XlsWeight));
	xlsBordTmp.put_LineStyle(_variant_t(xlLineStyle));
    
	//����Top�߿�
	lpDisp = xlsBords.get_Item(_variant_t((long)xlEdgeTop));
	if (lpDisp == NULL)
	{
		return;
	}
	xlsBordTmp.AttachDispatch(lpDisp);  
	xlsBordTmp.put_ColorIndex(_variant_t(coloIndex));  
	xlsBordTmp.put_Weight(_variant_t((long)XlsWeight));
	xlsBordTmp.put_LineStyle(_variant_t(xlLineStyle));

	//����Bottom�߿�
	lpDisp = xlsBords.get_Item(_variant_t((long)xlEdgeBottom));
	if (lpDisp == NULL)
	{
		return;
	}
	xlsBordTmp.AttachDispatch(lpDisp); 
	xlsBordTmp.put_ColorIndex(_variant_t(coloIndex));  
	xlsBordTmp.put_Weight(_variant_t((long)XlsWeight));
	xlsBordTmp.put_LineStyle(_variant_t(xlLineStyle));

	//����Right�߿�
	lpDisp = xlsBords.get_Item(_variant_t((long)xlEdgeRight));
	if (lpDisp == NULL)
	{
		return;
	}
	xlsBordTmp.AttachDispatch(lpDisp); 
 	xlsBordTmp.put_ColorIndex(_variant_t(coloIndex));  
 	xlsBordTmp.put_Weight(_variant_t((long)XlsWeight));
    xlsBordTmp.put_LineStyle(_variant_t(xlLineStyle));

	xlsBordTmp.ReleaseDispatch();
	xlsBords.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	rang.ReleaseDispatch();
	lpDisp = NULL;
}

//���Ƶ�����Ԫ��߿�
void CRLOperExcel::XlsDrawCellBorders(const CString &Cell, XlBordersLineStyle xlLineStyle, XlBorderWeight XlsWeight,long coloIndex,
	XlBordersIndex xlBordIndex)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (Cell.IsEmpty())
	{
		return;
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet    xlsSheet;
	if (lpDisp == NULL)
	{
		return;
	}
	xlsSheet.AttachDispatch(lpDisp);
	if (xlsSheet.m_lpDispatch == NULL)
	{
		return;
	}
	lpDisp = xlsSheet.get_Range(_variant_t(Cell), _variant_t(Cell));
	CRange  rang; 
	rang.AttachDispatch(lpDisp);
	CBorders xlsBords;   
	xlsBords.AttachDispatch(rang.get_Borders());

	CBorder  xlsBord;   //�ڱ߿�
	lpDisp = xlsBords.get_Item(_variant_t((long)xlBordIndex)); //�߿���ʽ����
	if (lpDisp == NULL)
	{
		return;
	}
	xlsBord.AttachDispatch(lpDisp);
	xlsBord.put_Weight(_variant_t((long)XlsWeight));
	xlsBord.put_ColorIndex(_variant_t(coloIndex));
	xlsBord.put_LineStyle(_variant_t(xlLineStyle));

	xlsBord.ReleaseDispatch();
	xlsBords.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	rang.ReleaseDispatch();
	lpDisp = NULL;
}

//���õ�Ԫ���doubleֵ
void CRLOperExcel::XlsSetCellData(const CString &cell, double dValue)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (cell.IsEmpty())
	{
		return;
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet  xlsSheet;
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp);
		if (xlsSheet.m_lpDispatch == NULL)
		{
			AfxMessageBox(_T("��ȡsheet�ӿ�ʧ��!"));
			return;
		}
		lpDisp = xlsSheet.get_Range((_variant_t)cell, (_variant_t)cell);
	}
	else
	{
		return;
	}

	CRange  range;
	if (lpDisp != NULL)
	{
		range.AttachDispatch(lpDisp);
		range.put_Value2((_variant_t)dValue);
	}	
	range.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	lpDisp = NULL;
}

//���õ�Ԫ����ַ���ֵ
void CRLOperExcel::XlsSetCellData(const CString &cell, const CString& strValue)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (cell.IsEmpty() || strValue.IsEmpty())
	{
		return;
	}

	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet  xlsSheet;
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp);
		lpDisp = xlsSheet.get_Range((_variant_t)cell, (_variant_t)cell);
	}
	else
	{
		return;
	}

	CRange  range;
	if (lpDisp != NULL)
	{
		range.AttachDispatch(lpDisp);
		range.put_Value2((_variant_t)strValue);
	}	
	range.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	lpDisp = NULL;
}

//���ڸ�����Ԫ���ַ���õ�double�͵���ֵ
BOOL CRLOperExcel::XlsGetCellValue(const CString &addr, double &dValue)
{
	if (!xlsAppIsInit())
	{
		return FALSE;
	}
	if (addr.IsEmpty())
	{
		return FALSE;
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet    xlsSheet;
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp);
		lpDisp = xlsSheet.get_Range((_variant_t)addr, (_variant_t)addr);
	}
	else
	{
		return FALSE;
	}
	CRange  range;
	if (lpDisp != NULL)
	{
		range.AttachDispatch(lpDisp);
		COleVariant vResult =range.get_Value2();

		if (vResult.vt == VT_R8)
		{
			dValue = vResult.dblVal;  //double
		}
	}
	xlsSheet.ReleaseDispatch();
	range.ReleaseDispatch();
	return TRUE;
}

//���ڸ�����Ԫ���ַ���õ�CString�͵���ֵ
BOOL CRLOperExcel::XlsGetCellValue(const CString &addr, CString &strValue)
{
	if (!xlsAppIsInit())
	{
		return FALSE;
	}
	if (addr.IsEmpty())
	{
		return FALSE;
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet    xlsSheet;
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp);
		lpDisp = xlsSheet.get_Range((_variant_t)addr, (_variant_t)addr);
	}
	else
	{
		return FALSE;
	}
	CRange  range;
	if (lpDisp != NULL)
	{
		range.AttachDispatch(lpDisp);
		COleVariant vResult =range.get_Value2();

		if (vResult.vt == VT_BSTR)
		{
			strValue = vResult.bstrVal;  //�ַ���
		}
	}
	xlsSheet.ReleaseDispatch();
	range.ReleaseDispatch();
	return TRUE;
}

//���Ϊ
void CRLOperExcel::XlsSaveAs(const CString &fileName,XlFileFormat NewFileFormat /*= xlExcel5*/)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (fileName.IsEmpty())
	{
		AfxMessageBox(_T("�ļ�������Ϊ��!"));
		return;
	}
	//XlFileFormat NewFileFormat = xlExcel12;
	CWorkbook xlsWorkBook;
	LPDISPATCH lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveWorkbook();
	if (lpDisp != NULL)
	{
		xlsWorkBook.AttachDispatch(lpDisp);
		if (xlsWorkBook.m_lpDispatch != NULL)
		{
			XlsCloseAlert();//�ر��ļ����Ǿ���
			xlsWorkBook.SaveAs(_variant_t(fileName), _variant_t((long)NewFileFormat), vtMissing, vtMissing, vtMissing, 
				vtMissing, 0, vtMissing, vtMissing, vtMissing, 
				vtMissing, vtMissing);
		}
		else
		{
			AfxMessageBox(_T("�����ļ�ʧ��!"));
			return;
		}
	}
	else
	{
		AfxMessageBox(_T("�����ļ�ʧ��!"));
		return;
	}
	xlsWorkBook.ReleaseDispatch();
	lpDisp = NULL;
}

//�˳�
void CRLOperExcel::XlsQuit()
{
	if (!xlsAppIsInit())
	{
		return;
	}
	try
	{
		m_xlsAppLication.Quit();
		m_bIsXlsProcRedu = FALSE;
	}
	catch(_com_error &e)
	{
		XlsErrInfo(e);
		return;
	}

}

void CRLOperExcel::XlsQuit1()
{
	// 	try{
	// 		pXL->Quit(); 
	// 		ProcRedun = FALSE;//ֻ�����ڴ�����һ��EXCEL ����
	// 	}catch(_com_error &e){
	// 		XlsErrInfo(e);
	// 		return;
	// 	}


}


//ǿ�ƹر�Ӧ�ó������
void CRLOperExcel::XlsTerminateExcelProcess(CString strProName)
{
	if (strProName.IsEmpty())
	{
		AfxMessageBox(_T("�������ֲ���Ϊ��!"));
		return;
	}
	HANDLE SnapShot, ProcessHandle;  
	SHFILEINFO shSmall;  
	PROCESSENTRY32 ProcessInfo;   
	CString strExeFile; //��������
	strProName.MakeLower();  //
	SnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);  
	if(SnapShot != NULL)   
	{  
		ProcessInfo.dwSize = sizeof(ProcessInfo);  // ����ProcessInfo�Ĵ�С  
		BOOL Status = Process32First(SnapShot, &ProcessInfo); 

		while(Status)  
		{  
			// ��ȡ�����ļ���Ϣ  
			SHGetFileInfo(ProcessInfo.szExeFile, 0, &shSmall, sizeof(shSmall), SHGFI_ICON|SHGFI_SMALLICON);  

			// �������Ƿ���Ҫ�ر�  
			strExeFile = ProcessInfo.szExeFile;
			strExeFile.MakeLower();  //Сд

			if(strProName.Compare(strExeFile) == 0)   
			{  
				// ��ȡ���̾����ǿ�йر�  
				ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, FALSE, ProcessInfo.th32ProcessID);  
				TerminateProcess(ProcessHandle, 1);  
				//break;  
			}  

			// ��ȡ��һ�����̵���Ϣ  
			Status = Process32Next(SnapShot, &ProcessInfo);  
		}  
	} 

}

//����һ��sheet
void CRLOperExcel::XlsAddSheet(const CString &strSheetName)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (strSheetName.IsEmpty())
	{
		return;
	}

	LPDISPATCH lpDisp = NULL;
	/*�õ��������е�Sheet������*/
	lpDisp = m_xlsAppLication.get_ActiveWorkbook();
	if (lpDisp == NULL)
	{
		return;
	}
	CWorkbook xlsWorkBook;
	xlsWorkBook.AttachDispatch(lpDisp);
	lpDisp = xlsWorkBook.get_Sheets();
	if (lpDisp == NULL)
	{
		return;
	}
	CWorksheets xlsWorkSheets;
	xlsWorkSheets.AttachDispatch(lpDisp);  //��ǰ�����sheets����

	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet    xlsSheet;
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp);
	}
	else
	{
		return;
	}
	/*��һ��Sheet���粻���ڣ�������һ��Sheet*/
	try
	{
		/*��һ�����е�Sheet*/
		lpDisp = xlsWorkSheets.get_Item(_variant_t(strSheetName));
		if (lpDisp != NULL)
		{
			xlsSheet.AttachDispatch(lpDisp);
		}		
	}
	catch(...)
	{
		/*����һ���µ�Sheet*/
		lpDisp = xlsWorkSheets.Add(vtMissing, vtMissing, _variant_t((long)1), vtMissing);
		if(lpDisp != NULL)
		{
			xlsSheet.AttachDispatch(lpDisp);
			xlsSheet.put_Name(strSheetName);
		}
	}
	xlsWorkBook.ReleaseDispatch();
	xlsWorkSheets.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	lpDisp = NULL;
}

//������sheet
void CRLOperExcel::XlsSetSheetName(long iIndex,const CString &strSheetName)
{
	if (!xlsAppIsInit())
	{
		return;
	}

	if (strSheetName.IsEmpty())
	{
		AfxMessageBox(_T("���ֲ���Ϊ��"));
		return ;
	}
	if (iIndex <= 0)
	{
		AfxMessageBox(_T("sheet����������ڵ�1"));
		return;
	}
	CWorksheets xlsWorksheets = XlsGetActiveWorkBookSheets();
	LPDISPATCH lpDisp = NULL;
	lpDisp = xlsWorksheets.get_Item(_variant_t(iIndex));
	CWorksheet xlsWorkSheet;
	if (lpDisp == NULL)
	{
		return;
	}
	xlsWorkSheet.AttachDispatch(lpDisp);
	if (xlsWorkSheet.m_lpDispatch == NULL)
	{
		return;
	}
	xlsWorkSheet.put_Name(strSheetName); //������

	xlsWorkSheet.ReleaseDispatch();
	xlsWorksheets.ReleaseDispatch();
	XlsReleaseActiveBookSheets();  //����ɶԳ���
	lpDisp = NULL;
}

//���ص�ǰ���������sheets����
long CRLOperExcel::XlsGetSheetCount()
{
	if (!xlsAppIsInit())
	{
		return -1;
	}
	LPDISPATCH lpDisp = NULL;
	long longCout = (long)0.0;
	lpDisp = m_xlsAppLication.get_ActiveWorkbook();
	CWorkbook xlsWorkBook;
	if (lpDisp == NULL)
	{
		return -1;
	}
	xlsWorkBook.AttachDispatch(lpDisp);
	CWorksheets    xlsSheets;
	lpDisp = xlsWorkBook.get_Sheets();
	if (lpDisp == NULL)
	{
		return -2;
	}
	xlsSheets.AttachDispatch(lpDisp); //��ǰ���������sheets����
	longCout = xlsSheets.get_Count();
	
	xlsSheets.ReleaseDispatch();
	xlsWorkBook.ReleaseDispatch();
	lpDisp = NULL;
	return longCout;
}

//�õ�sheet������
CString CRLOperExcel::XlsGetSheetName(int iIndex)
{
	if (!xlsAppIsInit())
	{
		return _T("erro");
	}
	CWorksheet xlsSheet;
	LPDISPATCH lpDisp = NULL;
	/*�õ��������е�Sheet������*/
	lpDisp = m_xlsAppLication.get_ActiveWorkbook();
	if (lpDisp == NULL)
	{
		return _T("erro");
	}
	CWorkbook xlsWorkBook;
	xlsWorkBook.AttachDispatch(lpDisp);    //�����workbook
	lpDisp = xlsWorkBook.get_Sheets();
	if (lpDisp == NULL)
	{
		return _T("erro");
	}
	CWorksheets xlsWorkSheets;
	xlsWorkSheets.AttachDispatch(lpDisp);  //��ǰ����Ĺ������м����sheets����

	lpDisp = xlsWorkSheets.get_Item(_variant_t((long)iIndex));
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp, true);
	}
	CString strSheetName = xlsSheet.get_Name();

	xlsSheet.ReleaseDispatch();
	xlsWorkSheets.ReleaseDispatch();
	xlsWorkBook.ReleaseDispatch();
	lpDisp = NULL;
	return strSheetName;
}

//��ȡworkBook����
CString CRLOperExcel::XlsGetWorkBookName(int iIndex)
{
	if (!xlsAppIsInit())
	{
		return _T("erro");
	}
	CWorkbook xlsWorkBook;
	LPDISPATCH lpDisp = NULL;
	CString strWorkBookName = _T("");
	if (iIndex <= m_xlsWorkBooks.get_Count() && iIndex >= 1)
	{
		lpDisp = m_xlsWorkBooks.get_Item(_variant_t((long)iIndex));
		if (lpDisp != NULL)
		{
			xlsWorkBook.AttachDispatch(lpDisp);
		}
		strWorkBookName = xlsWorkBook.get_Name();
	}
	xlsWorkBook.ReleaseDispatch();
	lpDisp = NULL;
	return strWorkBookName;
}

//����workBook����
int CRLOperExcel::XlsGetWorkBookIndex(const CString& strWorkName)
{
	if (!xlsAppIsInit())
	{
		return -1;
	}
	int iIndex = -1;
	if (!strWorkName.IsEmpty())
	{
		int i =1;
		for (i = 1; i <= XlsGetWorkBookCout(); i++)
		{
			if (strWorkName.Compare(XlsGetWorkBookName(i)) == 0)
			{
				iIndex = i;
				break;
			}
		}
	}
	else
	{
		iIndex =-1;
	}
	return iIndex;
}

//����workBook����
int CRLOperExcel::XlsGetWorkBookCout()
{
	if (!xlsAppIsInit())
	{
		return -1;
	}
	return m_xlsWorkBooks.get_Count();
}

//ɾ��sheet
void CRLOperExcel::XlsDelSheet(const CString &strSheetName)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	XlsCloseAlert();
	CWorksheet xlsSheet;
	if (1 == XlsGetSheetCount())
	{
		AfxMessageBox(_T("�����������ٺ���һ��sheet,������ɾ��ȫ��!"));
		return;
	}
	LPDISPATCH lpDisp = NULL;
 	/*�õ��������е�Sheet������*/
 	lpDisp = m_xlsAppLication.get_ActiveWorkbook();
 	if (lpDisp == NULL)
 	{
 		return;
 	}
 	CWorkbook xlsWorkBook;
 	xlsWorkBook.AttachDispatch(lpDisp);    //�����workbook
 	lpDisp = xlsWorkBook.get_Sheets();
 	if (lpDisp == NULL)
 	{
 		return;
 	}

	CWorksheets xlsWorkSheets;
	xlsWorkSheets.AttachDispatch(lpDisp);  //��ǰ����Ĺ������м����sheets����
	if (xlsWorkSheets.m_lpDispatch == NULL)
	{
		AfxMessageBox(_T("��ȡǰ����Ĺ������м����sheets����ʧ��"));
		return;
	}

	int i = 1; //������1��ʼ 
	for (i = 1; i <= XlsGetSheetCount(); i++)
	{
		CString &strSheetNameTmp = XlsGetSheetName(i);
		if (strSheetName.Compare(strSheetNameTmp) == 0)
		{
			lpDisp = xlsWorkSheets.get_Item(_variant_t((long)i));
			if (lpDisp != NULL)
			{
				xlsSheet.AttachDispatch(lpDisp, true);
				xlsSheet.Delete();
				break;
			}
		}
		else
		{   
			CString &strTmp = _T("ɾ��") + strSheetName +_T("ʧ��");
			AfxMessageBox(strTmp);
			return;
		}
	}
	xlsWorkSheets.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	xlsWorkBook.ReleaseDispatch();
	lpDisp = NULL;
}

//��������ɾ��sheet
void CRLOperExcel::XlsDelSheet(int Pos)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	XlsCloseAlert();
	if (Pos <= 0)
	{
		AfxMessageBox(_T("Sheet�����ű�����ڵ���1"));
		return;
	}
	CWorksheet xlsSheet;
	if (1 == XlsGetSheetCount())
	{
		AfxMessageBox(_T("�����������ٺ���һ��sheet,������ɾ��ȫ��!"));
		return;
	}
 	LPDISPATCH lpDisp = NULL;
	/*�õ��������е�Sheet������*/
	lpDisp = m_xlsAppLication.get_ActiveWorkbook();
	if (lpDisp == NULL)
	{
		return;
	}
	CWorkbook xlsWorkBook;
	xlsWorkBook.AttachDispatch(lpDisp);    //�����workbook
	lpDisp = xlsWorkBook.get_Sheets();
	if (lpDisp == NULL)
	{
		return;
	}

	CWorksheets xlsWorkSheets;
	xlsWorkSheets.AttachDispatch(lpDisp);  //��ǰ����Ĺ������м����sheets����

	if (xlsWorkSheets.m_lpDispatch == NULL)
	{
		AfxMessageBox(_T("��ȡǰ����Ĺ������м����sheets����ʧ��"));
		return;
	}

	lpDisp = xlsWorkSheets.get_Item(_variant_t((long)Pos));
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp, true);
		if (xlsSheet.m_lpDispatch == NULL)
		{
			return;
		}		
		xlsSheet.Delete();
	}
	else
	{
		AfxMessageBox(_T("Sheets����Ϊ��!"));//ɾ��ʧ��
	}
	xlsWorkSheets.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	lpDisp = NULL;
}

//�رո����ļ�ʱ��ľ���
void CRLOperExcel::XlsCloseAlert()
{
	if (!xlsAppIsInit())
	{
		return;
	}
	m_xlsAppLication.put_AlertBeforeOverwriting(FALSE); //�ر���ʾ
	m_xlsAppLication.put_DisplayAlerts(FALSE);
}

//�ж�Excel�Ƿ��ʼ��
BOOL CRLOperExcel::xlsAppIsInit()
{
	LPDISPATCH lpDisp = NULL;
	BOOL bRec = FALSE;
	lpDisp = m_xlsAppLication.m_lpDispatch;
	if (lpDisp == NULL)
	{
		AfxMessageBox(_T("���½��ļ�Excelδ��ʼ��!"));
		bRec = FALSE;
	}
	else
	{
		bRec = TRUE;
	}
	lpDisp = NULL;
	return bRec;
}

//����ַ������Ƿ�Ϸ�  ���������п� B:B
BOOL CRLOperExcel::XlsCheckStrIsValue(CString &str, CString &strResult /*out*/)
{
	BOOL bRec = FALSE;
	if (str.IsEmpty())
	{
		AfxMessageBox(_T("�����������ֵ!"));
		bRec = FALSE;
	}
	else
	{
		str.MakeUpper();
		int i = 0;
		int iIndex =0;
		int iCount = 0; //����
		for (i = 0; i< str.GetLength(); i++)
		{
			 if (str.GetAt(i) == ':')
			 {
				 iIndex++;
				 iCount = i;
			 } 
		}
		if (iIndex == 1)
		{
			/*B1:B1*/
			CString strFirst = str.Mid(0,iCount);
			strFirst.Trim();
			CString strSec = str.Mid(iCount+1);
			strSec.Trim();
			if (strFirst.Compare(strSec) == 0) //������� �������ñ����ʱ�����
			{
				strResult = strFirst + _T(":") + strSec;
				bRec = TRUE;
			}
			else
			{
				AfxMessageBox(_T("�����ַ������Ϸ�"));
				bRec = FALSE;
			}
		}
		else
		{
		   AfxMessageBox(_T("�����ַ������Ϸ�"));
		   bRec = FALSE;
		}
	}
	return bRec;
}

//�����п�
void CRLOperExcel::XlsSetColumnWidth(CString strRow/*_variant_t("B:B") or _variant_t("B1:B1")*/, double width/*0.0-255.0*/)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	CString strTmp;
	if (!XlsCheckStrIsValue(strRow, strTmp))
	{
		return;
	}
	
	/*�õ��������е�Sheet������*/
	if (width  <= 255.0000 && width >= 0.00 )
	{  
		LPDISPATCH lpDisp = NULL;
		lpDisp = m_xlsAppLication.get_ActiveSheet();
		CWorksheet    xlsSheet;
		ASSERT(lpDisp);
		xlsSheet.AttachDispatch(lpDisp);
		lpDisp = xlsSheet.get_Range(_variant_t(strTmp), _variant_t(strTmp));
		CRange  range;
		if (lpDisp != NULL)
		{
			range.AttachDispatch(lpDisp);
			range.put_ColumnWidth(_variant_t((long)width));
		}
		range.ReleaseDispatch();
		xlsSheet.ReleaseDispatch();
		lpDisp = NULL;
	}
	else
	{
		AfxMessageBox(_T("���ȡֵ��0.0--255.0֮��"));
		return;
	}
}

//�����и�
void CRLOperExcel::XlsSetRowHeight(int Row, double height)
{
	if (!xlsAppIsInit())
	{
		return;
	}

	CRange  range;
	LPDISPATCH lpDisp = NULL;
	/*�õ��������е�Sheet������*/
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet    xlsSheet;
	ASSERT(lpDisp);
	xlsSheet.AttachDispatch(lpDisp);
	range.AttachDispatch(xlsSheet.get_Rows(),true);
	range.AttachDispatch(range.get_Item(_variant_t((long)Row), vtMissing).pdispVal, true);
	range.put_RowHeight(_variant_t((long)height));
	range.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	lpDisp = NULL;
}

//����sheet�ı�׼��Ԫ���Ⱥ͸߶�
void CRLOperExcel::XlsSetHW(double Height,double Width)
{
	if (!xlsAppIsInit())
	{
		return;
	}
// 	LPDISPATCH lpDisp = NULL;
//   	
// 	lpDisp = m_xlsAppLication.get_ActiveSheet();
// 	CWorksheet xlsWorkSheet;
// 	xlsWorkSheet.AttachDispatch(lpDisp);
// 	CRange rang;
// 
// 	rang.put_UseStandardHeight();
}


//��һ��excel �ļ�
int CRLOperExcel::XlsOpen(const CString &strOpenFileName, const CString &prjName)
{
	if (!xlsAppIsInit())
	{
		return -1;
	}
	if (strOpenFileName .IsEmpty())
	{
		return -1;
	}
	LPDISPATCH lpDisp = NULL;
	LPDISPATCH lpDisp1 = NULL;
	lpDisp = m_xlsAppLication.get_Workbooks();
	if (lpDisp == NULL)
	{
		return -1;
	}
	/*�õ�����������*/
	m_xlsWorkBooks.AttachDispatch(lpDisp); 
	CWorkbook xlsWorkBook;
	try
	{
		/*��һ��������*/
		lpDisp = m_xlsWorkBooks.Open(strOpenFileName, 
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 
			vtMissing, vtMissing, vtMissing, vtMissing);

		xlsWorkBook.AttachDispatch(lpDisp);  
	}
	catch(...)
	{
		/*����һ���µĹ�����*/
		lpDisp = m_xlsWorkBooks.Add(vtMissing);
		if (lpDisp != NULL)
		{
			xlsWorkBook.AttachDispatch(lpDisp);
		}
	}
	xlsWorkBook.ReleaseDispatch();
	lpDisp = NULL;
	return 1;
}

//������һ��sheet
long CRLOperExcel::XlsActivateNextSheet()
{  
	if (!xlsAppIsInit())
	{
		return -1;
	}
	LPDISPATCH lpDisp = NULL;
	long longCout = (long)1.0;
	lpDisp = m_xlsAppLication.get_ActiveWorkbook();
	CWorkbook xlsWorkBook;
	if (lpDisp == NULL)
	{
		return -2;
	}
	xlsWorkBook.AttachDispatch(lpDisp);

	lpDisp = xlsWorkBook.get_ActiveSheet();
	if (lpDisp != NULL)
	{
		CWorksheet     xlsSheet;
		CWorksheets    xlsSheets;
		xlsSheets.AttachDispatch(xlsWorkBook.get_Sheets()); //��ǰ���������sheets����
		int iSheetCout = xlsSheets.get_Count();   //��ǰ���������sheet������
		xlsSheet.AttachDispatch(lpDisp);
		longCout = xlsSheet.get_Index();
		if (longCout >= 1 && longCout < iSheetCout/*XlsGetSheetCount()*/)
		{
			//������һ��sheet;
			longCout += (long)1.0;
			xlsSheet.AttachDispatch(xlsSheets.get_Item(_variant_t((long)longCout)), true);	
			xlsSheet.Activate();
		}
		else if (longCout == iSheetCout/*XlsGetSheetCount()*/)
		{
			//�����һ��sheet;
			xlsSheet.AttachDispatch(xlsSheets.get_Item(_variant_t((long)1)), true);
			xlsSheet.Activate();
			longCout = (long)1.0;
		}
		xlsSheet.ReleaseDispatch();
		xlsSheets.ReleaseDispatch();
	}
	lpDisp = NULL;
	return longCout;
}

//������һ��WorkBook����
long CRLOperExcel::XlsActivateNextWorkBook()
{
	if (!xlsAppIsInit())
	{
		return -1;
	}
	LPDISPATCH lpDisp = NULL;
	long longCout = (long)1.0;
	lpDisp  = m_xlsAppLication.get_ActiveWorkbook();
	CWorkbook  xlsWorkBook;
	if (lpDisp != NULL)
	{
		xlsWorkBook.AttachDispatch(lpDisp);
		CString strName = xlsWorkBook.get_Name();
		longCout = XlsGetWorkBookIndex(strName);

		if (longCout >= 1 && longCout < XlsGetWorkBookCout())
		{
			//������һ��WorkBook;
			longCout += (long)1.0;
			xlsWorkBook.AttachDispatch(m_xlsWorkBooks.get_Item(_variant_t((long)longCout)),true);
			xlsWorkBook.Activate();
		}
		else if (longCout == XlsGetWorkBookCout())
		{  
			//�����һ��WorkBook;
			xlsWorkBook.AttachDispatch(m_xlsWorkBooks.get_Item(_variant_t((long)1)),true);
			xlsWorkBook.Activate();
		}	
	}
	else
	{
		AfxMessageBox(_T("��ǰ�����ڴ򿪵Ĺ�����!"));
		return FALSE;
	}

	lpDisp  = NULL;
	xlsWorkBook.ReleaseDispatch();
	return longCout;
}

//��������ɾ��workBook
void CRLOperExcel::XlsDelWorkBook(const CString& strWorkName)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	XlsCloseAlert();
	int i = 1;
	for (i = 1; i <= XlsGetWorkBookCout(); i++)
	{
		if (strWorkName.Compare(XlsGetWorkBookName(i)) == 0)
		{
			COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
			LPDISPATCH lpDisp = NULL;
			lpDisp = m_xlsWorkBooks.get_Item(_variant_t((long)i));
			CWorkbook xlsWorkBook;
			if (lpDisp == NULL)
			{
				return;
			}
			xlsWorkBook.AttachDispatch(lpDisp);
			xlsWorkBook.Close(covOptional,covOptional,covOptional);
			xlsWorkBook.ReleaseDispatch();
			lpDisp = NULL;
			break;
		}
	}
}

//��������ɾ��������
void CRLOperExcel::XlsDelWorkBook(long iIndex)
{  
	if (!xlsAppIsInit())
	{
		return;
	}
	XlsCloseAlert();
	if (iIndex >= 1 && iIndex <= XlsGetWorkBookCout())
	{
		COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		LPDISPATCH lpDisp = NULL;
		lpDisp = m_xlsWorkBooks.get_Item(_variant_t((long)iIndex));
		CWorkbook xlsWorkBook;
		if (lpDisp == NULL)
		{
			return;
		}
		xlsWorkBook.AttachDispatch(lpDisp);
		xlsWorkBook.Close(covOptional,covOptional,covOptional);
		xlsWorkBook.ReleaseDispatch();
		lpDisp = NULL;
	}
	else
	{
		AfxMessageBox(_T("��ǰû�й���������ɾ��!"));
		return;
	}
}

//�ɵ�Ԫ���name���ԣ��õ���Ԫ���ַ(xlA1)
int CRLOperExcel::XlsName2Add(CString name, CString &col,UINT &row)
{
	//jd 
	return 1;
}


//set the excel app 's visible property, true for visible, vise versa
BOOL CRLOperExcel::XlsSetVisible(BOOL bIsVisible)
{
	if (!xlsAppIsInit())
	{
		return FALSE;
	}

	BOOL bRec = FALSE;
	BOOL bVis1 = m_xlsAppLication.get_Visible();
	m_xlsAppLication.put_Visible(bIsVisible);
	BOOL bVis2 = m_xlsAppLication.get_Visible();
	if (bVis1 != bVis2)
	{
		bRec = TRUE;
	}
	else
	{
		bRec = FALSE;
	}

	return bRec;
}

//�õ�ģ�����������
void CRLOperExcel::XlsGetLNum()
{
}


//��һ��ģ���е��������������������жϸ�ģ���Ƿ�Ϊ�кͲ��е�ģ��
//���кϲ��з���2��û�кϲ��з���1
int CRLOperExcel::XlsGetStep(CStringArray &nameArray)
{
	//jd 
	return 1;
}


void CRLOperExcel::XlsSetPage(int pageNum)
{

}

//���ù��
void CRLOperExcel::XlsSetCursor(XlsMouseCur nflag)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	switch(nflag)
	{
	case CRLOperExcel::xlWaitCur:
		m_xlsAppLication.put_Cursor(xlWait);
		break;
	case CRLOperExcel::xlDefaultCur:
		m_xlsAppLication.put_Cursor(xlDefault);
		break;
	case CRLOperExcel::xlIBeamCur:
		m_xlsAppLication.put_Cursor(xlIBeam);
		break;
	case CRLOperExcel::xlNorthwestArrowCur:
		m_xlsAppLication.put_Cursor(xlNorthwestArrow);
		break;
	default:
		break;
	}
}


//���к� ������
void CRLOperExcel::XlsRunMacro(const CString &macro)
{
	if (!xlsAppIsInit())
	{
		return;
	}

	m_xlsAppLication.Run(_variant_t(macro),vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,
			vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,
	    	vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing);
	//���ֶ������쳣������ᵼ���ڴ�й¶	
}

//�õ��������к������
BOOL CRLOperExcel::XlsGetMacroName(CStringArray &arrMarcoName)
{
	if (!xlsAppIsInit())
	{
		return FALSE;
	}
	arrMarcoName.RemoveAll();  //�����������
	LPDISPATCH lpDisp = NULL; 
	lpDisp = m_xlsAppLication.get_ActiveWorkbook();
	if (lpDisp == NULL)
	{
		return FALSE;
	}
	CWorkbook  xlsWorkBook;
	xlsWorkBook.AttachDispatch(lpDisp);
	if (xlsWorkBook.m_lpDispatch == NULL)
	{
		return FALSE;
	}
	BOOL bRec = xlsWorkBook.get_HasVBProject();  //�Ƿ����VB����
	if (!bRec)
	{
		return FALSE;
	}
	lpDisp = xlsWorkBook.get_VBProject();
	if (lpDisp == NULL)
	{
		return FALSE;
	}
	CVBProject xlsVBProj;
	xlsVBProj.AttachDispatch(lpDisp, TRUE);
	if (xlsVBProj.m_lpDispatch == NULL)
	{
		return FALSE;
	}
	CVBComponents xlsVBCompoents = xlsVBProj.get_VBComponents();
	if (xlsVBCompoents.m_lpDispatch == NULL)
	{
		return FALSE;
	}
	long lCompCount = xlsVBCompoents.get_Count();
	long lLineCount = 0;
	long lProcKind = vbext_pk_Proc;
	CVBComponent xlsVBCompoent;
	CCodeModule xlsCodeModule;
	long i ,j;
	CString sProcName;//,
	CString sItem;
	/*
	Constant        Value     Description
	-------------   -----     ------------------------------------------------
	vbext_pk_Get    3         Procedure that returns the value of a property 
	vbext_pk_Let    1         Procedure that assigns a value to a property

	vbext_pk_Set    2         Procedure that sets a reference to an object 
	vbext_pk_Proc   0         All procedures other than property procedures
	*/
	for(i = 1; i <= lCompCount; i++)
	{
		xlsVBCompoent = xlsVBCompoents.Item(_variant_t(i));
		xlsCodeModule = xlsVBCompoent.get_CodeModule();

		//If the component contains any lines of code, then
		//retrieve the name of each procedure (Functions and Subs)
		
		lLineCount = xlsCodeModule.get_CountOfLines();
		j=1;
		while( j< lLineCount)
		{
			sProcName = xlsCodeModule.get_ProcOfLine(j, &lProcKind);
			if(!sProcName.IsEmpty())
			{
				//sItem.Format(_T("%s\t\t%s"), xlsVBCompoent.get_Name(), sProcName);
				sItem = xlsVBCompoent.get_Name();
				arrMarcoName.Add(sProcName); //eg: macro1
				j = j + xlsCodeModule.get_ProcCountLines(sProcName, lProcKind);
				bRec = TRUE;
			}
			else 
			{				
				j++; 
			}
		}
	}
	xlsVBProj.ReleaseDispatch();
	xlsVBCompoents.ReleaseDispatch();
	xlsVBCompoent.ReleaseDispatch();
	xlsCodeModule.ReleaseDispatch();
	xlsWorkBook.ReleaseDispatch();
	lpDisp = NULL;
	return bRec; 
}

//���ļ��м��غ�
void CRLOperExcel::XlsAddMacroFromFlie(CString &strMacroFilePath)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (strMacroFilePath.IsEmpty())
	{   
		AfxMessageBox(_T("�ļ�·������Ϊ��!"));
		return;
	}
	LPDISPATCH lpDisp = NULL; 
	lpDisp = m_xlsAppLication.get_ActiveWorkbook();
	if (lpDisp == NULL)
	{
		return ;
	}
	CWorkbook  xlsWorkBook;
	xlsWorkBook.AttachDispatch(lpDisp);
	if (xlsWorkBook.m_lpDispatch == NULL)
	{
		return ;
	}
	BOOL bRec = xlsWorkBook.get_HasVBProject();  //�Ƿ����VB����
	if (!bRec)
	{
		//return ;  //������
	}
	else
	{

	}
	lpDisp = xlsWorkBook.get_VBProject();
	if (lpDisp == NULL)
	{
		return ;
	}

 	CVBProject xlsVBProject;
	xlsVBProject.AttachDispatch(lpDisp);
	lpDisp = xlsVBProject.get_VBComponents(); 
	if (lpDisp == NULL)
	{
		return ;
	}
	CVBComponents xlsVbComponents;
	xlsVbComponents.AttachDispatch(lpDisp);
	lpDisp = xlsVbComponents.Add(1);        // ditto
	if (lpDisp == NULL)
	{
		return ;
	}
	CVBComponent xlsVbComponent;
	xlsVbComponent.AttachDispatch(lpDisp);
	lpDisp = xlsVbComponent.get_CodeModule();
	if (lpDisp == NULL)
	{
		return ;
	}
	CCodeModule  xlsVbModule;  // ditto
	xlsVbModule.AttachDispatch(lpDisp);
	// Load the macro from the file into the VBA module
	// of the Excel document.
	 xlsVbModule.AddFromFile(strMacroFilePath); // Load the macro into the _CodeModule.

	 xlsVbComponent.ReleaseDispatch();
	 xlsVbComponents.ReleaseDispatch();
	 xlsVBProject.ReleaseDispatch();
	 xlsVbModule.ReleaseDispatch();
	 xlsWorkBook.ReleaseDispatch();
	 lpDisp = NULL;

}


int CRLOperExcel::CreateXlsNoInfo()
{

	// 	try {
	// 		if(ProcRedun) XlsQuit();//�ж��Ƿ���EXCEL����
	// 		
	//                
	// 		if(S_OK != pXL.GetActiveObject(L"Excel.Application"))
	// 		pXL.CreateInstance(L"Excel.Application");
	// //		DisplayFullScreen for ������ʾ
	// //		pXL->DisplayFullScreen = VARIANT_TRUE;
	// //		pXL->DisplayFullScreen = VARIANT_FALSE;
	// 
	// 		pXL->Visible = VARIANT_FALSE;
	// 
	// 		pBooks = pXL->Workbooks;
	// 		
	// 		ProcRedun = TRUE; //��EXCEL���̴����ˣ���
	// 
	// 	} catch(_com_error &e){
	// 		XlsErrInfo(e);
	// 	}

	return 1;
}


//close workbook
int CRLOperExcel::XlsCloseBook(CWorkbook &workBook)
{
	if (!xlsAppIsInit())
	{
		return -1;
	}

	XlsCloseAlert(); //�ر���ʾ��Ϣ!
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	//book.Close(covOptional,COleVariant(strPathName),covOptional);
	try
	{
		workBook.Close(covOptional,covOptional,covOptional);
	}
	catch(...)
	{
		XlsTerminateExcelProcess();
	}
	//�ڲ��ͷŽӿ�
	workBook.ReleaseDispatch();
	return 1;
}


int CRLOperExcel::XlsCloseBook()
{
	if (!xlsAppIsInit())
	{
		return -1;
	}

	XlsCloseAlert(); //�ر���ʾ��Ϣ!
	try
	{
		m_xlsWorkBooks.Close();  //�˳�����
	}
	catch(...)
	{
		XlsTerminateExcelProcess();
	}
	return 1;
}


//����:
//����:
//���:
//����:
BOOL CRLOperExcel::XlsAddLine(const CString &v_hBegin, const CString &v_hEnd, XlInsertFormatOrigin xlsInsertFormat)
{	
	if (!xlsAppIsInit())
	{
		return FALSE;
	}

	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	if (lpDisp == NULL)
	{
		return FALSE;
	}
	CWorksheet  xlsSheet;
	xlsSheet.AttachDispatch(lpDisp);

	lpDisp = xlsSheet.get_Range(_variant_t(v_hBegin),_variant_t(v_hEnd));
	ASSERT(lpDisp);
	CRange rang;
	rang.AttachDispatch(lpDisp); 
	// CRange rang;
	//rang.AttachDispatch(m_xlsAppLication.get_ActiveCell());
	//���ڱ༭״̬��һֱ�ȴ�
	rang.AttachDispatch(rang.get_EntireRow());
	/*
	enum XlInsertFormatOrigin
	{
	xlFormatFromLeftOrAbove = 0,
	xlFormatFromRightOrBelow = 1
	};
	*/
	rang.Insert(vtMissing, _variant_t(xlsInsertFormat));
	xlsSheet.ReleaseDispatch();
	rang.ReleaseDispatch();
	lpDisp = NULL;
	return TRUE;
}

//����:
//����:
//���:
//����:
BOOL CRLOperExcel::XmlGetRangeWidthHeight(RangePtr v_hRangePtr,double &v_dWidth,double &v_dHeight)
{

	return true;
}

void CRLOperExcel::XlsErrInfo(_com_error &e)
{

}

//��õ�ǰ����Ĺ�����
 CWorkbook CRLOperExcel::XlsGetActiveBook()
 {  
	 LPDISPATCH lpDisp = NULL;
	 if (!xlsAppIsInit())
	 {
		 return lpDisp = NULL;
	 }
	
	 lpDisp = m_xlsAppLication.get_ActiveWorkbook();
	 m_xlsWorkBook.AttachDispatch(lpDisp);
	 if (m_xlsWorkBook.m_lpDispatch == NULL)
	 {
		 return lpDisp = NULL;
	 }
	 return m_xlsWorkBook;
 }

 //�ͷŵ�ǰ����Ĺ������ӿ�
 void CRLOperExcel::XlsReleaseActiveBook()
 {
	 if (!xlsAppIsInit())
	 {
		 return;
	 }

	 if (m_xlsWorkBook.m_lpDispatch != NULL)
	 {
		 m_xlsWorkBook.ReleaseDispatch();
	 }
 }

//�ϲ���Ԫ��
void CRLOperExcel::XlsMergeCells(int &iRow, int &iCol,int &iMergeCellRow,int &iMergeCellCol)
{
	//�ϲ���Ԫ��Ĵ���
	//�����жϵڣ�iRow��iCol������Ԫ���Ƿ�Ϊ�ϲ���Ԫ���Լ����ڣ�iRow��iCol������Ԫ����кϲ�
	//������
	if (!xlsAppIsInit())
	{
		return;
	}

	if (!(iRow >=1 && iCol >= 1))
	{
		AfxMessageBox(_T("����������������ڵ���1��������������!"));
		return;
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();

	if (lpDisp == NULL)
	{
		return;
	}
	CWorksheet  xlsSheet;
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp);
	}

	CRange unionRange;
	CRange range;
	range.AttachDispatch(xlsSheet.get_Cells()); 
	unionRange.AttachDispatch(range.get_Item(COleVariant((long)iRow),COleVariant((long)iRow)).pdispVal );

	VARIANT vResult=unionRange.get_MergeCells();    
	if(vResult.boolVal==-1)             //�Ǻϲ��ĵ�Ԫ��    
	{
		//�ϲ���Ԫ������� 
		range.AttachDispatch (unionRange.get_Rows());
		long iUnionRowNum=range.get_Count(); 

		//�ϲ���Ԫ�������
		range.AttachDispatch (unionRange.get_Columns());
		long iUnionColumnNum=range.get_Count ();   

		//�ϲ��������ʼ�У���
		long iUnionStartRow=unionRange.get_Row();       //��ʼ�У���1��ʼ
		long iUnionStartCol=unionRange.get_Column();    //��ʼ�У���1��ʼ

	}
	else if(vResult.boolVal==0)   
	{
		//���Ǻϲ��ĵ�Ԫ��
		//����һ����Ԫ��ϲ���iMergeCellRow�У�iMergeCellCol��
		range.AttachDispatch(xlsSheet.get_Cells()); 
		unionRange.AttachDispatch(range.get_Item (COleVariant((long)iRow),COleVariant((long)iRow)).pdispVal );
		unionRange.AttachDispatch(unionRange.get_Resize(COleVariant((long)iMergeCellRow),COleVariant((long)iMergeCellCol)));
		unionRange.Merge(COleVariant((long)0));   //�ϲ���Ԫ��
	}
	unionRange.ReleaseDispatch();
	range.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	lpDisp = NULL;
}


//�����Ԫ������
void CRLOperExcel::XlsClearRangeContens(const CString& strCell)
{
	
	if (!xlsAppIsInit())
	{
		return;
	}

	if (strCell.IsEmpty())
	{
		return;
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet    xlsSheet;
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp);
		lpDisp = xlsSheet.get_Range((_variant_t)strCell, (_variant_t)strCell);
	}
	else
	{
		return;
	}
	CRange  range;
	if (lpDisp != NULL)
	{
		range.AttachDispatch(lpDisp);
		range.Select();
		range.ClearContents();
	}
	range.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	lpDisp = NULL;

}

//�Զ���� �����кϲ���Ԫ��  �����ִ��ʧ��
void CRLOperExcel::XlsAutoFillCell(const CString &strRow, const CString &strCol,const CString& strValue,
	XlAutoFillType xlsAtuoFillType/* = xlFillDefault*/)
{
	
	if (!xlsAppIsInit())
	{
		return;
	}

	if (strRow.IsEmpty())
	{
		return;
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet    xlsSheet;
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp);
		lpDisp = xlsSheet.get_Range((_variant_t)strRow, (_variant_t)strRow);
	}
	else
	{
		return;
	}
	CRange  range;
	CRange  rangeDes; //Ҫ��������
	if (lpDisp != NULL)
	{
		range.AttachDispatch(lpDisp);
		range.put_Value2(_variant_t(strValue));
		lpDisp = xlsSheet.get_Range((_variant_t)strRow, (_variant_t)strCol);
		rangeDes.AttachDispatch(lpDisp);
		range.AutoFill(rangeDes.m_lpDispatch, xlsAtuoFillType);
	}
	range.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	rangeDes.ReleaseDispatch();
	lpDisp = NULL;
}

//����Ȳ�����P1��P2, P3
void CRLOperExcel::XlsAutoFillArithmeticSequenceCell(const CString &strRow, const CString &strCol,
	const CString& strFirstValue,const CString& strSecValue, XlAutoFillType xlsAtuoFillType /* = xlFillDefault */)
{
	
	if (!xlsAppIsInit())
	{
		return;
	}

	if (strRow.IsEmpty())
	{
		return;
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet    xlsSheet;
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp);
		lpDisp = xlsSheet.get_Range((_variant_t)strRow, (_variant_t)strRow);
	}
	else
	{
		return;
	}
	CRange  range;
	CRange  rangeDes; //Ҫ��������
	if (lpDisp != NULL)
	{
		range.AttachDispatch(lpDisp);
		range.put_Value2(_variant_t(strFirstValue));
		range.put_Value2(_variant_t(strSecValue));
		lpDisp = xlsSheet.get_Range((_variant_t)strRow, (_variant_t)strCol);
		rangeDes.AttachDispatch(lpDisp);
		range.AutoFill(rangeDes.m_lpDispatch, xlsAtuoFillType);
	}
	range.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	rangeDes.ReleaseDispatch();
	lpDisp = NULL;
}

//�Ȳ����� 1 2 3
void CRLOperExcel::XlsAutoFillNoArithSequenceCell(const CString &strFirstRow/*A1*/, const CString &strSecRow/*A2*/, 
	const CString &strCol/*A20*/,double dFirtValue, double dSecValue,XlAutoFillType xlsAtuoFillType /*= xlFillDefault*/)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	
	if (strFirstRow.IsEmpty())
	{
		return;
	}
	LPDISPATCH  lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	CWorksheet    xlsSheet;
	if (lpDisp != NULL)
	{
		xlsSheet.AttachDispatch(lpDisp);
		lpDisp = xlsSheet.get_Range((_variant_t)strFirstRow, (_variant_t)strFirstRow);
	}
	else
	{
		return;
	}
	CRange  range1;
	CRange  range2;
	CRange  rangeDes; //Ҫ��������
	CRange  rangTmp;
	if (lpDisp != NULL)
	{
		range1.AttachDispatch(lpDisp);
		range1.put_Value2(_variant_t((double)dFirtValue));
	    lpDisp = xlsSheet.get_Range((_variant_t)strSecRow, (_variant_t)strSecRow);
		if (lpDisp == NULL)
		{
			return;
		}
		range2.AttachDispatch(lpDisp);
		range2.put_Value2(_variant_t((double)dSecValue));
		if (lpDisp == NULL)
		{
			return;
		}
		lpDisp = xlsSheet.get_Range((_variant_t)strFirstRow, (_variant_t)strSecRow);
		rangTmp.AttachDispatch(lpDisp);
		lpDisp = xlsSheet.get_Range((_variant_t)strFirstRow, (_variant_t)strCol);
		if (lpDisp == NULL)
		{
			return;
		}
		rangeDes.AttachDispatch(lpDisp);
		rangTmp.AutoFill(rangeDes.m_lpDispatch, xlsAtuoFillType);
	}
	
	range1.ReleaseDispatch();
	range2.ReleaseDispatch();
	rangTmp.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	rangeDes.ReleaseDispatch();
	lpDisp = NULL;
}

/*�õ���ǰ����Ĺ������е�Sheets������*/
//�˽ӿڵ��ú��ͷ��ǳɶԵġ�
CWorksheets CRLOperExcel::XlsGetActiveWorkBookSheets()
{
	LPDISPATCH lpDisp = NULL;
	if (!xlsAppIsInit())
	{
		return lpDisp = NULL;
	}
	
	/*�õ��������е�Sheets������*/
	lpDisp = m_xlsAppLication.get_ActiveWorkbook();
	if (lpDisp == NULL)
	{
		return lpDisp = NULL;
	}
	CWorkbook xlsWorkBook;
	xlsWorkBook.AttachDispatch(lpDisp);
	if (xlsWorkBook.m_lpDispatch == NULL)
	{
		return lpDisp = NULL;
	}
	lpDisp = xlsWorkBook.get_Sheets();
	if (lpDisp == NULL)
	{
		return lpDisp = NULL;
	}
	m_xlsWorkSheets.AttachDispatch(lpDisp);
	if (m_xlsWorkSheets.m_lpDispatch == NULL)
	{
		return lpDisp = NULL;
	}
	xlsWorkBook.ReleaseDispatch();
    lpDisp = NULL;
	return m_xlsWorkSheets;
}

//�˽ӿڵ��ú��ͷ��ǳɶԵĳ��֡�
void CRLOperExcel::XlsReleaseActiveBookSheets()
{
	if (m_xlsWorkSheets.m_lpDispatch != NULL)
	{
		m_xlsWorkSheets.ReleaseDispatch();
	}
}

//���õ�Ԫ����뷽ʽ
void CRLOperExcel::XlsSetCellAlignment(int row, int colum, CellHorizonAlignment horizon, CellVerticalAlignment vertical)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	/*
	row  colum  �ж�
	*/
	if (!(row >= 1 && colum >= 1))
	{
		AfxMessageBox(_T("����������������ڵ���1��������������!"));
		return;
	}

	CWorksheet xlsWorkSheet = m_xlsAppLication.get_ActiveSheet();
	if (xlsWorkSheet.m_lpDispatch == NULL)
	{
		return;
	}
	CRange range;
	LPDISPATCH lpDisp = NULL;
	lpDisp = xlsWorkSheet.get_Cells();  //���б��
	if (lpDisp == NULL)
	{
		return;
	}

	range.AttachDispatch(lpDisp);
	if (range.m_lpDispatch == NULL)
	{
		return;
	}
	
	range.get_Item(_variant_t((long)row), _variant_t((long)colum));
	switch(horizon)
	{
	case CRLOperExcel::xlLeft:
		range.put_HorizontalAlignment(_variant_t(xlLeft));
		break;
	case CRLOperExcel::xlRight:
		range.put_HorizontalAlignment(_variant_t(xlRight));
		break;
	case CRLOperExcel::xlCenter:
		range.put_HorizontalAlignment(_variant_t(xlCenter));
		break;
	case CRLOperExcel::xlFill:
		range.put_HorizontalAlignment(_variant_t(xlFill));
		break;
	case CRLOperExcel::xlJustify:
		range.put_HorizontalAlignment(_variant_t(xlJustify));
		break;
	case CRLOperExcel::xlCenterAcrossSelection:
		range.put_HorizontalAlignment(_variant_t(xlCenterAcrossSelection));
		break;
	case CRLOperExcel::xlDistributed:
		range.put_HorizontalAlignment(_variant_t(xlDistributed));
		break;
	default:
		break;
	}
	switch (vertical)
	{
	case CRLOperExcel::xlTop:
		range.put_VerticalAlignment(_variant_t(xlTop));
		break;
	case CRLOperExcel::xlBottom:
		range.put_VerticalAlignment(_variant_t(xlBottom));
		break;
	case CRLOperExcel::xlJustify:
		range.put_VerticalAlignment(_variant_t(xlJustify));
		break;
	case CRLOperExcel::xlDistributed:
		range.put_VerticalAlignment(_variant_t(xlDistributed));
		break;
	default:
		break;
	}
	range.ReleaseDispatch();
	xlsWorkSheet.ReleaseDispatch();
	lpDisp = NULL;
}

//���õ�Ԫ��Ķ��뷽ʽ
void CRLOperExcel::XlsSetCellAlignment(const CString& cellRow, const CString& cellCol, CellHorizonAlignment horizon, CellVerticalAlignment vertical)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (cellRow.IsEmpty() || cellCol.IsEmpty())
	{
		AfxMessageBox(_T("��������в���Ϊ��!"));
		return;
	}
	CWorksheet xlsWorkSheet = m_xlsAppLication.get_ActiveSheet();
	if (xlsWorkSheet.m_lpDispatch == NULL)
	{
		return;
	}
	CRange range;
	LPDISPATCH lpDisp = NULL;
	lpDisp = xlsWorkSheet.get_Range(_variant_t(cellRow),_variant_t(cellCol));
	if (lpDisp == NULL)
	{
		return;
	}
	range.AttachDispatch(lpDisp);
	if (range.m_lpDispatch == NULL)
	{
		return;
	}
	switch(horizon)  //ˮƽ
	{
	case CRLOperExcel::xlLeft:
		range.put_HorizontalAlignment(_variant_t(xlLeft));
		break;
	case CRLOperExcel::xlRight:
		range.put_HorizontalAlignment(_variant_t(xlRight));
		break;
	case CRLOperExcel::xlCenter:
		range.put_HorizontalAlignment(_variant_t(xlCenter));
		break;
	case CRLOperExcel::xlFill:
		range.put_HorizontalAlignment(_variant_t(xlFill));
		break;
	case CRLOperExcel::xlJustify:
		range.put_HorizontalAlignment(_variant_t(xlJustify));
		break;
	case CRLOperExcel::xlCenterAcrossSelection:
		range.put_HorizontalAlignment(_variant_t(xlCenterAcrossSelection));
		break;
	case CRLOperExcel::xlDistributed:
		range.put_HorizontalAlignment(_variant_t(xlDistributed));
		break;
	default:
		break;
	}
	switch (vertical)  //��ֱ
	{
	case CRLOperExcel::xlTop:
		range.put_VerticalAlignment(_variant_t(xlTop));
		break;
	case CRLOperExcel::xlBottom:
		range.put_VerticalAlignment(_variant_t(xlBottom));
		break;
	case CRLOperExcel::xlvJustify:
		range.put_VerticalAlignment(_variant_t(xlvJustify));
		break;
	case CRLOperExcel::xlvDistributed:
		range.put_VerticalAlignment(_variant_t(xlvDistributed));
		break;
	default:
		break;
	}
	range.ReleaseDispatch();
	xlsWorkSheet.ReleaseDispatch();
	lpDisp = NULL;
}

//����ͼƬ
void CRLOperExcel::XlsInsertPictrue(const CString &strPicPath,float xPos, float yPos,float picWidth, float picHeight)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (strPicPath.IsEmpty())
	{
		AfxMessageBox(_T("�ļ�·������Ϊ��!"));
		return;
	}
	const float fAccuracy = 0.0000000000;
	if (xPos < fAccuracy || yPos < fAccuracy || picWidth < fAccuracy || picHeight < fAccuracy)
	{
		AfxMessageBox(_T("����Ĳ���������ڵ���0"));
		return;
	}
	CShapes xlsShapes;
	LPDISPATCH lpDisp = NULL;
	lpDisp = m_xlsAppLication.get_ActiveSheet();
	if (lpDisp == NULL)
	{
		return;
	}
	CWorksheet xlsWorkSheet;
	xlsWorkSheet.AttachDispatch(lpDisp);

	lpDisp = xlsWorkSheet.get_Shapes();
	if (lpDisp == NULL)
	{
		return;
	}
	xlsShapes.AttachDispatch(lpDisp);
	if (xlsShapes.m_lpDispatch == NULL)
	{
		return;
	}
	xlsShapes.AddPicture(strPicPath,false,true,xPos,yPos,10,10);
	lpDisp = xlsShapes.get_Range(_variant_t((long)1));
	CShape xlsShape;
	if (lpDisp == NULL)
	{
		return;
	}

	xlsShape.AttachDispatch(lpDisp);
	if (xlsShape.m_lpDispatch == NULL)
	{
		return;
	}
	xlsShape.put_Name(_T("JD"));  //������
	xlsShape.put_Width(picWidth);
	xlsShape.put_Height(picHeight);

	xlsShapes.ReleaseDispatch();
	xlsShape.ReleaseDispatch();
	xlsWorkSheet.ReleaseDispatch();
	lpDisp = NULL;
}

boost::optional<CWorkbook> CRLOperExcel::OpenFile(const CString& fileName)
{
	boost::optional<CWorkbook> ret;
	if (!xlsAppIsInit() || fileName.IsEmpty())
		return ret;

	LPDISPATCH lpDisp = m_xlsAppLication.get_Workbooks();
	if (lpDisp == NULL)
		return ret;
	
	CWorkbooks workBooks(lpDisp);
	CWorkbook  workBook;
	try
	{
		LPDISPATCH lpDisp = workBooks.Open(fileName, 
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 
			vtMissing, vtMissing, vtMissing, vtMissing);

		if (lpDisp)
		{
			workBook.AttachDispatch(lpDisp);  
			ret.reset(workBook);
		}
	}
	catch(...) {}

	return ret;
}

boost::optional<CWorkbook> CRLOperExcel::NewFile()
{
	boost::optional<CWorkbook> ret;
	if (!xlsAppIsInit())
		return ret;

	LPDISPATCH lpDisp = m_xlsAppLication.get_Workbooks();
	if (lpDisp == NULL)
		return ret;

	CWorkbooks workBooks(lpDisp);
	CWorkbook  workBook;

	LPDISPATCH lpNewBook = workBooks.Add(vtMissing);
	if (lpNewBook)
	{
		workBook.AttachDispatch(lpNewBook);  
		ret.reset(workBook);
	}

	return ret;
}

bool CRLOperExcel::SaveAs(CWorkbook& workBook, const CString& fileName)
{
	if (!xlsAppIsInit() || fileName.IsEmpty() || workBook.m_lpDispatch == NULL)
		return false;

	workBook.SaveAs(_variant_t(fileName), vtMissing, vtMissing, vtMissing, vtMissing, 
		vtMissing, 0, vtMissing, vtMissing, vtMissing, vtMissing, vtMissing);

	return true;
}

void CRLOperExcel::Close(CWorkbook& workBook)
{
	workBook.Close(_variant_t(false), vtMissing, vtMissing);
}

int CRLOperExcel::GetSheetCount(CWorkbook& workBook)
{
	int ret(0);

	CWorksheets sheets(workBook.get_Worksheets());
	if (sheets.m_lpDispatch == NULL)
		return ret;

	ret = (int)sheets.get_Count();
	return ret;
}

int CRLOperExcel::AddSheetBack(CWorkbook& workBook)
{
	int ret(0);

	CWorksheets sheets(workBook.get_Worksheets());
	if (sheets.m_lpDispatch == NULL)
		return ret;

	LPDISPATCH lpNewSheet = NULL;
	if (sheets.get_Count() == 0)
		lpNewSheet = sheets.Add(vtMissing, vtMissing, vtMissing, vtMissing);
	else
	{
		LPDISPATCH lpdisp = sheets.get_Item(_variant_t(sheets.get_Count()));
		if (lpdisp == NULL)
			return ret;

		lpNewSheet = sheets.Add(vtMissing, _variant_t(lpdisp), vtMissing, vtMissing);
		lpdisp->Release();
	}
	if (lpNewSheet == NULL)
		return ret;

	CWorksheet newSheet(lpNewSheet);
	ret = (int)newSheet.get_Index();
	return ret;
}

void CRLOperExcel::SetSheetName(CWorkbook& workBook, int nSheetIndex, const CString& name)
{
	CWorksheets sheets(workBook.get_Worksheets());
	if (sheets.m_lpDispatch == NULL)
		return;

	CWorksheet sheet(sheets.get_Item(_variant_t(nSheetIndex)));
	if (sheet.m_lpDispatch == NULL)
		return;

	try {
		sheet.put_Name(name);
	}
	catch (...) {}
}

bool CRLOperExcel::CopySheetContents(CWorkbook& srcBook, int nSrcSheet, CWorkbook& destBook, int nDestSheet)
{
	try
	{
		if (!xlsAppIsInit())
			return false;

		CWorksheets srcSheets(srcBook.get_Worksheets());
		if (srcSheets.m_lpDispatch == NULL)
			return false;

		// ����Դ��Ԫ��
		CWorksheet srcSheet(srcSheets.get_Item(_variant_t(nSrcSheet)));
		if (srcSheet.m_lpDispatch == NULL)
			return false;

		CRange srcAllCells(srcSheet.get_Cells());
		if (srcAllCells.m_lpDispatch == NULL)
			return false;
		srcAllCells.Copy(vtMissing);

		// ճ����Ŀ�깤����
		CWorksheets destSheets(destBook.get_Worksheets());
		if (destSheets.m_lpDispatch == NULL)
			return false;

		CWorksheet destSheet(destSheets.get_Item(_variant_t(nDestSheet)));
		if (destSheet.m_lpDispatch == NULL)
			return false;
		destSheet.Paste(vtMissing, vtMissing);
		// ճ�������б�ճ���ĵ�Ԫ���ڱ�ѡ��״̬��ȡ������״ֻ̬ѡ�е�һ����Ԫ��
		CRange range(destSheet.get_Range(_variant_t(_T("A1")), vtMissing));
		if (range.m_lpDispatch)
		{
			try {
				range.Select();
			} catch (...) {

			}
		}

		// ȡ������ģʽ
		m_xlsAppLication.put_CutCopyMode(FALSE);
	}
	catch (...)
	{
		return false;
	}

	return true;
}

void CRLOperExcel::ActiveSheet(CWorkbook& workBook, int nSheetIndex)
{
	CWorksheets sheets(workBook.get_Worksheets());
	if (sheets.m_lpDispatch == NULL)
		return;

	CWorksheet sheet(sheets.get_Item(_variant_t(nSheetIndex)));
	if (sheet.m_lpDispatch == NULL)
		return;

	sheet.Activate();
}

void CRLOperExcel::SetCellText(CWorkbook& workBook, int nSheetIndex, int row, int column, const CString& text)
{
	CWorksheets sheets(workBook.get_Worksheets());
	if (sheets.m_lpDispatch == NULL)
		return;

	CWorksheet sheet(sheets.get_Item(_variant_t(nSheetIndex)));
	if (sheet.m_lpDispatch == NULL)
		return;

	CRange allCells(sheet.get_Cells());
	if (allCells.m_lpDispatch == NULL)
		return;

	CRange cell((LPDISPATCH)_variant_t(allCells.get_Item(_variant_t(row), _variant_t(column))));
	if (cell.m_lpDispatch == NULL)
		return;

	cell.put_Value2(_variant_t(text));
}

int CRLOperExcel::GetMergeCellsCount(CWorkbook& workBook, int nSheetIndex, int row, int column)
{
	CWorksheets sheets(workBook.get_Worksheets());
	if (sheets.m_lpDispatch == NULL)
		return 0;

	CWorksheet sheet(sheets.get_Item(_variant_t(nSheetIndex)));
	if (sheet.m_lpDispatch == NULL)
		return 0;

	CRange allCells(sheet.get_Cells());
	if (allCells.m_lpDispatch == NULL)
		return 0;

	CRange cell((LPDISPATCH)_variant_t(allCells.get_Item(_variant_t(row), _variant_t(column))));
	if (cell.m_lpDispatch == NULL)
		return 0;

	CRange mergeArea(cell.get_MergeArea());
	if (mergeArea.m_lpDispatch == NULL)
		return 0;
	return (int)mergeArea.get_Count();
}

void CRLOperExcel::SetShrinkToFit(CWorkbook& workBook, int nSheetIndex, int row, int column)
{
	CWorksheets sheets(workBook.get_Worksheets());
	if (sheets.m_lpDispatch == NULL)
		return;

	CWorksheet sheet(sheets.get_Item(_variant_t(nSheetIndex)));
	if (sheet.m_lpDispatch == NULL)
		return;

	CRange allCells(sheet.get_Cells());
	if (allCells.m_lpDispatch == NULL)
		return;

	CRange cell((LPDISPATCH)_variant_t(allCells.get_Item(_variant_t(row), _variant_t(column))));
	if (cell.m_lpDispatch == NULL)
		return;

	cell.put_ShrinkToFit(_variant_t(true));
}
