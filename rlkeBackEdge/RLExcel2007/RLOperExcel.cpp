

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
		AfxMessageBox(_T("初始化ole库失败!"));
		return ;
	}
}

CRLOperExcel::~CRLOperExcel()
{
	/*释放资源*/
	try
	{    //未初始化的话Excel App  关闭自己处理异常
		XlsCloseAlert(); //关闭提示信息!
		try
		{
			m_xlsWorkBooks.Close();  //退出进程
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
	AfxOleTerm(FALSE);//关闭ole库 
}

//从表格模板创建Excel文件
BOOL CRLOperExcel::CreateXlsFromXltFile(const CString &strXltFileName, BOOL bIsShowExcel)
{
	static int iAppCout = 1;  //应用程序的数量
	if (iAppCout > 1)
	{
		AfxMessageBox(_T("不允许重复初始化!"));
		return FALSE;
	}
	if (m_bIsXlsProcRedu)
	{
		XlsQuit();//判断是否有EXCEL进程
	}
	if (!m_xlsAppLication.CreateDispatch(_T("Excel.Application"), NULL))
	{
		AfxMessageBox(_T("启动Excel服务器失败!"));
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

//创建，初始化Excel
int CRLOperExcel::CreateXls(BOOL bIsShowExcel)
{
	try
	{
		if (m_bIsXlsProcRedu)
		{
			XlsQuit();//判断是否有EXCEL进程
		}
		if (!m_xlsAppLication.CreateDispatch(_T("Excel.Application"), NULL))
		{
			AfxMessageBox(_T("启动Excel服务器失败!"));
			return -1;
		}
		/*判断当前Excel的版本*/
// 		CString &strExcelVersion = m_xlsAppLication.get_Version();
// 		int iStart = 0;
// 		strExcelVersion = strExcelVersion.Tokenize(_T("."), iStart);
// 		if (_T("11") == strExcelVersion)
// 		{
// 			AfxMessageBox(_T("当前Excel的版本是2003。"));
// 		}
// 		else if (_T("12") == strExcelVersion)
// 		{
// 			AfxMessageBox(_T("当前Excel的版本是2007。"));
// 		}
// 		else
// 		{
// 			AfxMessageBox(_T("当前Excel的版本是其他版本。"));
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

// 向一组单元格中插入数据
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

//获得一组单元格中的数据   jd 待完善
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
		// 			strTmp = vResult.bstrVal;  //字符串
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
		// 			dResult = vResult.dblVal;  //字符串
		// 		}

	}
	xlsSheet.ReleaseDispatch();
	range.ReleaseDispatch();
	return dResult;
}

//向一组单元格中插入数据
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

//绘制内边框
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
	CBorder  xlsBord;   //内边框
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

//绘制外边框
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
	CBorders xlsBords;   //外
	xlsBords.AttachDispatch(lpDisp);
	//边框
	CBorder xlsBordTmp;

	//绘制Left边框
	lpDisp = xlsBords.get_Item(_variant_t((long)xlEdgeLeft));
	if (lpDisp == NULL)
	{
		return;
	}
 	xlsBordTmp.AttachDispatch(lpDisp);  
 	xlsBordTmp.put_ColorIndex(_variant_t(coloIndex));  
 	xlsBordTmp.put_Weight(_variant_t((long)XlsWeight));
	xlsBordTmp.put_LineStyle(_variant_t(xlLineStyle));
    
	//绘制Top边框
	lpDisp = xlsBords.get_Item(_variant_t((long)xlEdgeTop));
	if (lpDisp == NULL)
	{
		return;
	}
	xlsBordTmp.AttachDispatch(lpDisp);  
	xlsBordTmp.put_ColorIndex(_variant_t(coloIndex));  
	xlsBordTmp.put_Weight(_variant_t((long)XlsWeight));
	xlsBordTmp.put_LineStyle(_variant_t(xlLineStyle));

	//绘制Bottom边框
	lpDisp = xlsBords.get_Item(_variant_t((long)xlEdgeBottom));
	if (lpDisp == NULL)
	{
		return;
	}
	xlsBordTmp.AttachDispatch(lpDisp); 
	xlsBordTmp.put_ColorIndex(_variant_t(coloIndex));  
	xlsBordTmp.put_Weight(_variant_t((long)XlsWeight));
	xlsBordTmp.put_LineStyle(_variant_t(xlLineStyle));

	//绘制Right边框
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

//绘制单个单元格边框
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

	CBorder  xlsBord;   //内边框
	lpDisp = xlsBords.get_Item(_variant_t((long)xlBordIndex)); //边框样式索引
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

//设置单元格的double值
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
			AfxMessageBox(_T("获取sheet接口失败!"));
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

//设置单元格的字符串值
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

//对于给定单元格地址，得到double型的数值
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

//对于给定单元格地址，得到CString型的数值
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
			strValue = vResult.bstrVal;  //字符串
		}
	}
	xlsSheet.ReleaseDispatch();
	range.ReleaseDispatch();
	return TRUE;
}

//另存为
void CRLOperExcel::XlsSaveAs(const CString &fileName,XlFileFormat NewFileFormat /*= xlExcel5*/)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (fileName.IsEmpty())
	{
		AfxMessageBox(_T("文件名不能为空!"));
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
			XlsCloseAlert();//关闭文件覆盖警告
			xlsWorkBook.SaveAs(_variant_t(fileName), _variant_t((long)NewFileFormat), vtMissing, vtMissing, vtMissing, 
				vtMissing, 0, vtMissing, vtMissing, vtMissing, 
				vtMissing, vtMissing);
		}
		else
		{
			AfxMessageBox(_T("保存文件失败!"));
			return;
		}
	}
	else
	{
		AfxMessageBox(_T("保存文件失败!"));
		return;
	}
	xlsWorkBook.ReleaseDispatch();
	lpDisp = NULL;
}

//退出
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
	// 		ProcRedun = FALSE;//只考虑内存中有一个EXCEL 进程
	// 	}catch(_com_error &e){
	// 		XlsErrInfo(e);
	// 		return;
	// 	}


}


//强制关闭应用程序进程
void CRLOperExcel::XlsTerminateExcelProcess(CString strProName)
{
	if (strProName.IsEmpty())
	{
		AfxMessageBox(_T("进程名字不能为空!"));
		return;
	}
	HANDLE SnapShot, ProcessHandle;  
	SHFILEINFO shSmall;  
	PROCESSENTRY32 ProcessInfo;   
	CString strExeFile; //进程名字
	strProName.MakeLower();  //
	SnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);  
	if(SnapShot != NULL)   
	{  
		ProcessInfo.dwSize = sizeof(ProcessInfo);  // 设置ProcessInfo的大小  
		BOOL Status = Process32First(SnapShot, &ProcessInfo); 

		while(Status)  
		{  
			// 获取进程文件信息  
			SHGetFileInfo(ProcessInfo.szExeFile, 0, &shSmall, sizeof(shSmall), SHGFI_ICON|SHGFI_SMALLICON);  

			// 检测进程是否需要关闭  
			strExeFile = ProcessInfo.szExeFile;
			strExeFile.MakeLower();  //小写

			if(strProName.Compare(strExeFile) == 0)   
			{  
				// 获取进程句柄，强行关闭  
				ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, FALSE, ProcessInfo.th32ProcessID);  
				TerminateProcess(ProcessHandle, 1);  
				//break;  
			}  

			// 获取下一个进程的信息  
			Status = Process32Next(SnapShot, &ProcessInfo);  
		}  
	} 

}

//增加一个sheet
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
	/*得到工作簿中的Sheet的容器*/
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
	xlsWorkSheets.AttachDispatch(lpDisp);  //当前激活的sheets容器

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
	/*打开一个Sheet，如不存在，就新增一个Sheet*/
	try
	{
		/*打开一个已有的Sheet*/
		lpDisp = xlsWorkSheets.get_Item(_variant_t(strSheetName));
		if (lpDisp != NULL)
		{
			xlsSheet.AttachDispatch(lpDisp);
		}		
	}
	catch(...)
	{
		/*创建一个新的Sheet*/
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

//重命名sheet
void CRLOperExcel::XlsSetSheetName(long iIndex,const CString &strSheetName)
{
	if (!xlsAppIsInit())
	{
		return;
	}

	if (strSheetName.IsEmpty())
	{
		AfxMessageBox(_T("名字不能为空"));
		return ;
	}
	if (iIndex <= 0)
	{
		AfxMessageBox(_T("sheet索引必须大于等1"));
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
	xlsWorkSheet.put_Name(strSheetName); //重命名

	xlsWorkSheet.ReleaseDispatch();
	xlsWorksheets.ReleaseDispatch();
	XlsReleaseActiveBookSheets();  //必须成对出现
	lpDisp = NULL;
}

//返回当前激活工作簿中sheets数量
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
	xlsSheets.AttachDispatch(lpDisp); //当前激活工作簿中sheets数量
	longCout = xlsSheets.get_Count();
	
	xlsSheets.ReleaseDispatch();
	xlsWorkBook.ReleaseDispatch();
	lpDisp = NULL;
	return longCout;
}

//得到sheet的名字
CString CRLOperExcel::XlsGetSheetName(int iIndex)
{
	if (!xlsAppIsInit())
	{
		return _T("erro");
	}
	CWorksheet xlsSheet;
	LPDISPATCH lpDisp = NULL;
	/*得到工作簿中的Sheet的容器*/
	lpDisp = m_xlsAppLication.get_ActiveWorkbook();
	if (lpDisp == NULL)
	{
		return _T("erro");
	}
	CWorkbook xlsWorkBook;
	xlsWorkBook.AttachDispatch(lpDisp);    //激活的workbook
	lpDisp = xlsWorkBook.get_Sheets();
	if (lpDisp == NULL)
	{
		return _T("erro");
	}
	CWorksheets xlsWorkSheets;
	xlsWorkSheets.AttachDispatch(lpDisp);  //当前激活的工作簿中激活的sheets容器

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

//获取workBook名字
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

//返回workBook索引
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

//返回workBook数量
int CRLOperExcel::XlsGetWorkBookCout()
{
	if (!xlsAppIsInit())
	{
		return -1;
	}
	return m_xlsWorkBooks.get_Count();
}

//删除sheet
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
		AfxMessageBox(_T("工作表中至少含有一个sheet,不允许删除全部!"));
		return;
	}
	LPDISPATCH lpDisp = NULL;
 	/*得到工作簿中的Sheet的容器*/
 	lpDisp = m_xlsAppLication.get_ActiveWorkbook();
 	if (lpDisp == NULL)
 	{
 		return;
 	}
 	CWorkbook xlsWorkBook;
 	xlsWorkBook.AttachDispatch(lpDisp);    //激活的workbook
 	lpDisp = xlsWorkBook.get_Sheets();
 	if (lpDisp == NULL)
 	{
 		return;
 	}

	CWorksheets xlsWorkSheets;
	xlsWorkSheets.AttachDispatch(lpDisp);  //当前激活的工作簿中激活的sheets容器
	if (xlsWorkSheets.m_lpDispatch == NULL)
	{
		AfxMessageBox(_T("获取前激活的工作簿中激活的sheets容器失败"));
		return;
	}

	int i = 1; //索引从1开始 
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
			CString &strTmp = _T("删除") + strSheetName +_T("失败");
			AfxMessageBox(strTmp);
			return;
		}
	}
	xlsWorkSheets.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	xlsWorkBook.ReleaseDispatch();
	lpDisp = NULL;
}

//根据索引删除sheet
void CRLOperExcel::XlsDelSheet(int Pos)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	XlsCloseAlert();
	if (Pos <= 0)
	{
		AfxMessageBox(_T("Sheet索引号必须大于等于1"));
		return;
	}
	CWorksheet xlsSheet;
	if (1 == XlsGetSheetCount())
	{
		AfxMessageBox(_T("工作表中至少含有一个sheet,不允许删除全部!"));
		return;
	}
 	LPDISPATCH lpDisp = NULL;
	/*得到工作簿中的Sheet的容器*/
	lpDisp = m_xlsAppLication.get_ActiveWorkbook();
	if (lpDisp == NULL)
	{
		return;
	}
	CWorkbook xlsWorkBook;
	xlsWorkBook.AttachDispatch(lpDisp);    //激活的workbook
	lpDisp = xlsWorkBook.get_Sheets();
	if (lpDisp == NULL)
	{
		return;
	}

	CWorksheets xlsWorkSheets;
	xlsWorkSheets.AttachDispatch(lpDisp);  //当前激活的工作簿中激活的sheets容器

	if (xlsWorkSheets.m_lpDispatch == NULL)
	{
		AfxMessageBox(_T("获取前激活的工作簿中激活的sheets容器失败"));
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
		AfxMessageBox(_T("Sheets不能为空!"));//删除失败
	}
	xlsWorkSheets.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	lpDisp = NULL;
}

//关闭覆盖文件时候的警告
void CRLOperExcel::XlsCloseAlert()
{
	if (!xlsAppIsInit())
	{
		return;
	}
	m_xlsAppLication.put_AlertBeforeOverwriting(FALSE); //关闭提示
	m_xlsAppLication.put_DisplayAlerts(FALSE);
}

//判断Excel是否初始化
BOOL CRLOperExcel::xlsAppIsInit()
{
	LPDISPATCH lpDisp = NULL;
	BOOL bRec = FALSE;
	lpDisp = m_xlsAppLication.m_lpDispatch;
	if (lpDisp == NULL)
	{
		AfxMessageBox(_T("请新建文件Excel未初始化!"));
		bRec = FALSE;
	}
	else
	{
		bRec = TRUE;
	}
	lpDisp = NULL;
	return bRec;
}

//检查字符输入是否合法  用于设置列宽 B:B
BOOL CRLOperExcel::XlsCheckStrIsValue(CString &str, CString &strResult /*out*/)
{
	BOOL bRec = FALSE;
	if (str.IsEmpty())
	{
		AfxMessageBox(_T("不允许输入空值!"));
		bRec = FALSE;
	}
	else
	{
		str.MakeUpper();
		int i = 0;
		int iIndex =0;
		int iCount = 0; //索引
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
			if (strFirst.Compare(strSec) == 0) //必须相等 否则设置表格宽度时候出错
			{
				strResult = strFirst + _T(":") + strSec;
				bRec = TRUE;
			}
			else
			{
				AfxMessageBox(_T("输入字符串不合法"));
				bRec = FALSE;
			}
		}
		else
		{
		   AfxMessageBox(_T("输入字符串不合法"));
		   bRec = FALSE;
		}
	}
	return bRec;
}

//设置列宽
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
	
	/*得到工作簿中的Sheet的容器*/
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
		AfxMessageBox(_T("宽度取值在0.0--255.0之间"));
		return;
	}
}

//设置行高
void CRLOperExcel::XlsSetRowHeight(int Row, double height)
{
	if (!xlsAppIsInit())
	{
		return;
	}

	CRange  range;
	LPDISPATCH lpDisp = NULL;
	/*得到工作簿中的Sheet的容器*/
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

//设置sheet的标准单元格宽度和高度
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


//打开一个excel 文件
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
	/*得到工作簿容器*/
	m_xlsWorkBooks.AttachDispatch(lpDisp); 
	CWorkbook xlsWorkBook;
	try
	{
		/*打开一个工作簿*/
		lpDisp = m_xlsWorkBooks.Open(strOpenFileName, 
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 
			vtMissing, vtMissing, vtMissing, vtMissing);

		xlsWorkBook.AttachDispatch(lpDisp);  
	}
	catch(...)
	{
		/*增加一个新的工作簿*/
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

//激活下一个sheet
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
		xlsSheets.AttachDispatch(xlsWorkBook.get_Sheets()); //当前激活工作簿的sheets数量
		int iSheetCout = xlsSheets.get_Count();   //当前激活工作簿中sheet的数量
		xlsSheet.AttachDispatch(lpDisp);
		longCout = xlsSheet.get_Index();
		if (longCout >= 1 && longCout < iSheetCout/*XlsGetSheetCount()*/)
		{
			//激活下一个sheet;
			longCout += (long)1.0;
			xlsSheet.AttachDispatch(xlsSheets.get_Item(_variant_t((long)longCout)), true);	
			xlsSheet.Activate();
		}
		else if (longCout == iSheetCout/*XlsGetSheetCount()*/)
		{
			//激活第一个sheet;
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

//激活下一个WorkBook窗口
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
			//激活下一个WorkBook;
			longCout += (long)1.0;
			xlsWorkBook.AttachDispatch(m_xlsWorkBooks.get_Item(_variant_t((long)longCout)),true);
			xlsWorkBook.Activate();
		}
		else if (longCout == XlsGetWorkBookCout())
		{  
			//激活第一个WorkBook;
			xlsWorkBook.AttachDispatch(m_xlsWorkBooks.get_Item(_variant_t((long)1)),true);
			xlsWorkBook.Activate();
		}	
	}
	else
	{
		AfxMessageBox(_T("当前无正在打开的工作簿!"));
		return FALSE;
	}

	lpDisp  = NULL;
	xlsWorkBook.ReleaseDispatch();
	return longCout;
}

//根据名字删除workBook
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

//根据索引删除工作簿
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
		AfxMessageBox(_T("当前没有工作簿可以删除!"));
		return;
	}
}

//由单元格的name属性，得到单元格地址(xlA1)
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

//得到模板的数据行数
void CRLOperExcel::XlsGetLNum()
{
}


//把一个模板中的所有属性名传进来，判断该模板是否为有和并行的模板
//是有合并行返回2，没有合并行返回1
int CRLOperExcel::XlsGetStep(CStringArray &nameArray)
{
	//jd 
	return 1;
}


void CRLOperExcel::XlsSetPage(int pageNum)
{

}

//设置光标
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


//运行宏 待完善
void CRLOperExcel::XlsRunMacro(const CString &macro)
{
	if (!xlsAppIsInit())
	{
		return;
	}

	m_xlsAppLication.Run(_variant_t(macro),vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,
			vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,
	    	vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing);
	//不手动处理异常，否则会导致内存泄露	
}

//得到工作簿中宏的名字
BOOL CRLOperExcel::XlsGetMacroName(CStringArray &arrMarcoName)
{
	if (!xlsAppIsInit())
	{
		return FALSE;
	}
	arrMarcoName.RemoveAll();  //清除所有内容
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
	BOOL bRec = xlsWorkBook.get_HasVBProject();  //是否存在VB工程
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

//从文件中加载宏
void CRLOperExcel::XlsAddMacroFromFlie(CString &strMacroFilePath)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (strMacroFilePath.IsEmpty())
	{   
		AfxMessageBox(_T("文件路径不能为空!"));
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
	BOOL bRec = xlsWorkBook.get_HasVBProject();  //是否存在VB工程
	if (!bRec)
	{
		//return ;  //待完善
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
	// 		if(ProcRedun) XlsQuit();//判断是否有EXCEL进程
	// 		
	//                
	// 		if(S_OK != pXL.GetActiveObject(L"Excel.Application"))
	// 		pXL.CreateInstance(L"Excel.Application");
	// //		DisplayFullScreen for 正常显示
	// //		pXL->DisplayFullScreen = VARIANT_TRUE;
	// //		pXL->DisplayFullScreen = VARIANT_FALSE;
	// 
	// 		pXL->Visible = VARIANT_FALSE;
	// 
	// 		pBooks = pXL->Workbooks;
	// 		
	// 		ProcRedun = TRUE; //有EXCEL进程存在了！！
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

	XlsCloseAlert(); //关闭提示信息!
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
	//内部释放接口
	workBook.ReleaseDispatch();
	return 1;
}


int CRLOperExcel::XlsCloseBook()
{
	if (!xlsAppIsInit())
	{
		return -1;
	}

	XlsCloseAlert(); //关闭提示信息!
	try
	{
		m_xlsWorkBooks.Close();  //退出进程
	}
	catch(...)
	{
		XlsTerminateExcelProcess();
	}
	return 1;
}


//功能:
//输入:
//输出:
//返回:
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
	//处于编辑状态会一直等待
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

//功能:
//输入:
//输出:
//返回:
BOOL CRLOperExcel::XmlGetRangeWidthHeight(RangePtr v_hRangePtr,double &v_dWidth,double &v_dHeight)
{

	return true;
}

void CRLOperExcel::XlsErrInfo(_com_error &e)
{

}

//获得当前激活的工作簿
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

 //释放当前激活的工作簿接口
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

//合并单元格
void CRLOperExcel::XlsMergeCells(int &iRow, int &iCol,int &iMergeCellRow,int &iMergeCellCol)
{
	//合并单元格的处理
	//包括判断第（iRow，iCol）个单元格是否为合并单元格，以及将第（iRow，iCol）个单元格进行合并
	//待完善
	if (!xlsAppIsInit())
	{
		return;
	}

	if (!(iRow >=1 && iCol >= 1))
	{
		AfxMessageBox(_T("行列索引都必须大于等于1，请检查您的输入!"));
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
	if(vResult.boolVal==-1)             //是合并的单元格    
	{
		//合并单元格的行数 
		range.AttachDispatch (unionRange.get_Rows());
		long iUnionRowNum=range.get_Count(); 

		//合并单元格的列数
		range.AttachDispatch (unionRange.get_Columns());
		long iUnionColumnNum=range.get_Count ();   

		//合并区域的起始行，列
		long iUnionStartRow=unionRange.get_Row();       //起始行，从1开始
		long iUnionStartCol=unionRange.get_Column();    //起始列，从1开始

	}
	else if(vResult.boolVal==0)   
	{
		//不是合并的单元格
		//将第一个单元格合并成iMergeCellRow行，iMergeCellCol列
		range.AttachDispatch(xlsSheet.get_Cells()); 
		unionRange.AttachDispatch(range.get_Item (COleVariant((long)iRow),COleVariant((long)iRow)).pdispVal );
		unionRange.AttachDispatch(unionRange.get_Resize(COleVariant((long)iMergeCellRow),COleVariant((long)iMergeCellCol)));
		unionRange.Merge(COleVariant((long)0));   //合并单元格
	}
	unionRange.ReleaseDispatch();
	range.ReleaseDispatch();
	xlsSheet.ReleaseDispatch();
	lpDisp = NULL;
}


//清除单元格内容
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

//自动填充 不能有合并单元格  否则会执行失败
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
	CRange  rangeDes; //要填充的区域
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

//插入等差数列P1，P2, P3
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
	CRange  rangeDes; //要填充的区域
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

//等差数列 1 2 3
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
	CRange  rangeDes; //要填充的区域
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

/*得到当前激活的工作簿中的Sheets的容器*/
//此接口调用和释放是成对的。
CWorksheets CRLOperExcel::XlsGetActiveWorkBookSheets()
{
	LPDISPATCH lpDisp = NULL;
	if (!xlsAppIsInit())
	{
		return lpDisp = NULL;
	}
	
	/*得到工作簿中的Sheets的容器*/
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

//此接口调用和释放是成对的出现。
void CRLOperExcel::XlsReleaseActiveBookSheets()
{
	if (m_xlsWorkSheets.m_lpDispatch != NULL)
	{
		m_xlsWorkSheets.ReleaseDispatch();
	}
}

//设置单元格对齐方式
void CRLOperExcel::XlsSetCellAlignment(int row, int colum, CellHorizonAlignment horizon, CellVerticalAlignment vertical)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	/*
	row  colum  判断
	*/
	if (!(row >= 1 && colum >= 1))
	{
		AfxMessageBox(_T("行列索引都必须大于等于1，请检查您的输入!"));
		return;
	}

	CWorksheet xlsWorkSheet = m_xlsAppLication.get_ActiveSheet();
	if (xlsWorkSheet.m_lpDispatch == NULL)
	{
		return;
	}
	CRange range;
	LPDISPATCH lpDisp = NULL;
	lpDisp = xlsWorkSheet.get_Cells();  //所有表格
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

//设置单元格的对齐方式
void CRLOperExcel::XlsSetCellAlignment(const CString& cellRow, const CString& cellCol, CellHorizonAlignment horizon, CellVerticalAlignment vertical)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (cellRow.IsEmpty() || cellCol.IsEmpty())
	{
		AfxMessageBox(_T("输入的行列不能为空!"));
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
	switch(horizon)  //水平
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
	switch (vertical)  //垂直
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

//插入图片
void CRLOperExcel::XlsInsertPictrue(const CString &strPicPath,float xPos, float yPos,float picWidth, float picHeight)
{
	if (!xlsAppIsInit())
	{
		return;
	}
	if (strPicPath.IsEmpty())
	{
		AfxMessageBox(_T("文件路径不能为空!"));
		return;
	}
	const float fAccuracy = 0.0000000000;
	if (xPos < fAccuracy || yPos < fAccuracy || picWidth < fAccuracy || picHeight < fAccuracy)
	{
		AfxMessageBox(_T("输入的参数必须大于等于0"));
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
	xlsShape.put_Name(_T("JD"));  //待完善
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

		// 复制源单元格
		CWorksheet srcSheet(srcSheets.get_Item(_variant_t(nSrcSheet)));
		if (srcSheet.m_lpDispatch == NULL)
			return false;

		CRange srcAllCells(srcSheet.get_Cells());
		if (srcAllCells.m_lpDispatch == NULL)
			return false;
		srcAllCells.Copy(vtMissing);

		// 粘贴到目标工作表
		CWorksheets destSheets(destBook.get_Worksheets());
		if (destSheets.m_lpDispatch == NULL)
			return false;

		CWorksheet destSheet(destSheets.get_Item(_variant_t(nDestSheet)));
		if (destSheet.m_lpDispatch == NULL)
			return false;
		destSheet.Paste(vtMissing, vtMissing);
		// 粘贴后所有被粘贴的单元格处于被选定状态，取消这种状态只选中第一个单元格
		CRange range(destSheet.get_Range(_variant_t(_T("A1")), vtMissing));
		if (range.m_lpDispatch)
		{
			try {
				range.Select();
			} catch (...) {

			}
		}

		// 取消复制模式
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
