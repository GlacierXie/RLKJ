
#pragma  once 
#ifndef  RLOPER_EXCEL_H
#define  RLOPER_EXCEL_H

#include "ExcelLibHeader/CApplication.h"
#include "ExcelLibHeader/CRange.h"
#include "ExcelLibHeader/CRanges.h"
#include "ExcelLibHeader/CSheets.h"
#include "ExcelLibHeader/CWorkbook.h"
#include "ExcelLibHeader/CWorkbooks.h"
#include "ExcelLibHeader/CWorksheet.h"
#include "ExcelLibHeader/CWorksheets.h"
#include "ExcelLibHeader/CShape.h"
#include "ExcelLibHeader/CShapes.h"
#include "ExcelLibHeader/CBorder.h"
#include "ExcelLibHeader/CBorders.h"
#include "VBHeader/CApplication0.h"
#include "VBHeader/CCodeModule.h"
#include "VBHeader/CVBProject.h"
#include "VBHeader/CVBComponents.h"
#include "VBHeader/CVBComponent.h"
#include <boost/optional.hpp>


class/* AFX_CLASS_IMPORT*/ CRLOperExcel  
{
public:
	CRLOperExcel();
	virtual ~CRLOperExcel();
public:

	//workBook����
	//������һ�������� ����������ID
	long    XlsActivateNextWorkBook();
	CString XlsGetWorkBookName(int iIndex);
	int     XlsGetWorkBookIndex(const CString& strWorkName);
	int     XlsGetWorkBookCout();
	void    XlsDelWorkBook(long iIndex);
	void    XlsDelWorkBook(const CString& strWorkName);

	/*
	��ȡ���ͷű���ɶԳ���
	*/
	CWorkbook XlsGetActiveBook();
	void    XlsReleaseActiveBook();

	int  XlsCloseBook(CWorkbook &workBook);			//close workbook
	int  XlsCloseBook();

	//����һ��
	BOOL XlsAddLine(const CString& v_hBegin, const CString& v_hEnd, XlInsertFormatOrigin xlsInsertFormat = xlFormatFromRightOrBelow);
	BOOL XmlGetRangeWidthHeight(RangePtr v_hRangePtr,double &v_dWidth,double &v_dHeight);
	
	void XlsRunMacro(const CString &macro);//���к�
	
	BOOL  XlsGetMacroName(CStringArray &arrMarcoName);//�õ���ǰExcel�д��ڵĺ�����
	
	void XlsAddMacroFromFlie(CString &strMacroFilePath);//���ļ��м���VBA��
	
	enum XlsMouseCur
	{
		xlIBeamCur = 3,
		xlDefaultCur = -4143,
		xlNorthwestArrowCur = 1,
		xlWaitCur = 2
	};
   void XlsSetCursor(XlsMouseCur nflag);

	//����XLs�ļ�
	/*
	����Excel�ļ������ַ�ʽ  
	1���ӱ��ģ��*.xlt�д���
	2���г����Զ�����
	�����Ƿ����ͬʱ��ʼ��������
	*/

	/*----------------------------------------------------------------------------*
	* ��  ��:  �ӱ��ģ�崴��XLs�ļ�
	* ��  ��:  	[in]  strXltFileName  xlt�ļ�ȫ·��
	*           [in]  bIsShowExcel    �Ƿ���ʾ�����ɹ���Excel�ļ�             
	* ����ֵ:   �ɹ����� -- TRUE      ʧ�ܷ��� -- FALSE
	* ��  ʷ:  1.0 ���� 2013-10-31   create
	-----------------------------------------------------------------------------*/
	BOOL CreateXlsFromXltFile(const CString &strXltFileName, BOOL bIsShowExcel = TRUE);
	int CreateXls(BOOL bIsShowExcel = TRUE);
	int CreateXlsNoInfo();

    //�򿪺��½����� 
	int  XlsOpen(const CString &strOpenFileName, const CString &prjName);

	int  XlsName2Add(CString name,  CString &col,UINT &row);

	/*----------------------------------------------------------------------------*
	* ��  ��:  ��һ�鵥Ԫ���в�������
	* ��  ��:  	[in]  Row  ��n��
	*           [in]  Col  ��n��
	*           [in]  strValue  ������ַ�������
	* ����ֵ:  
	* ��  ʷ:  1.0 ���� 2013-10-31   create
	-----------------------------------------------------------------------------*/
	void XlsSetRangeData(const CString &Row, const CString &Col, const CString &strValue);
	void XlsSetRangeData(const CString &Row, const CString &Col, double dValue);

	CString XlsGetRangeStrData(const CString &Row, const CString &Col);
	double XlsGetRangeDoubleData(const CString &Row, const CString &Col);
	
	enum XlBordersLineStyle
	{
		xlContinuous = 1,
		xlDash = -4115,
		xlDashDot = 4,
		xlDashDotDot = 5,
		xlDot = -4118,
		xlDouble = -4119,
		xlSlantDashDot = 13,
		xlLineStyleNone = -4142
	};
	
	/*
	enum XlBordersIndex
	{
	xlInsideHorizontal = 12,
	xlInsideVertical = 11,
	xlDiagonalDown = 5,
	xlDiagonalUp = 6,
	xlEdgeBottom = 9,
	xlEdgeLeft = 7,
	xlEdgeRight = 10,
	xlEdgeTop = 8
	};
	*/
	enum XlBorderWeight
	{
		xlHairline = 1,
		xlMedium = -4138,
		xlThick = 4,
		xlThin = 2
	};
	/*----------------------------------------------------------------------------*
	* ��  ��:   ���Ʊ߿�
	* ��  ��:  	[in]  Row  ��n��
	*           [in]  Col  ��n��
	*           [in]  xlLineStyle  �߿�������ʽ
	*           [in]  XlsWeight     �߿�Ŀ��
	*           [in]  coloIndex     ��ɫ���� ��������ڣ�0-100��֮��
	* ����ֵ:  
	* ��  ʷ:  1.0 ���� 2013-10-31   create
	-----------------------------------------------------------------------------*/
	//�����ڱ߿�
	void XlsDrawInBorders(const CString &Row, const CString &Col,XlBordersLineStyle xlLineStyle, 
		XlBorderWeight XlsWeight,long coloIndex/*0-100*/);

	//������߿�
	void XlsDrawOutBorders(const CString &Row, const CString &Col,XlBordersLineStyle xlLineStyle, 
		XlBorderWeight XlsWeight,long coloIndex/*0-100*/);
		
    /*
	enum XlBordersIndex
	{
	xlInsideHorizontal = 12,
	xlInsideVertical = 11,
	xlDiagonalDown = 5, // ����
	xlDiagonalUp = 6,   //  '/'
	xlEdgeBottom = 9,
	xlEdgeLeft = 7,
	xlEdgeRight = 10,
	xlEdgeTop = 8
	};
	*/
	//���Ƶ�����Ԫ��߿���ʽ
	void XlsDrawCellBorders(const CString &Cell, XlBordersLineStyle xlLineStyle, 
		XlBorderWeight XlsWeight,long coloIndex/*0-100*/, XlBordersIndex xlBordIndex);

	/*----------------------------------------------------------------------------*
	* ��  ��:  �򵥸���Ԫ���в�������
	* ��  ��:  [in]  Cell  ��n����Ԫ��
	*          [in]  strValue  ������ַ�������
	* ����ֵ:  
	* ��  ʷ:  1.0 ���� 2013-10-31   create
	-----------------------------------------------------------------------------*/
	void XlsSetCellData(const CString &cell, double dValue);
	void XlsSetCellData(const CString &cell, const CString& strValue);

	/*----------------------------------------------------------------------------*
	* ��  ��:  ���ڸ�����Ԫ���ַ��double�͵���ֵ
	* ��  ��:  [in]  addr        ��n����Ԫ��                     eg: addr = A1  ��ֵ�� 100.0
	*          [out] dValue      ����double���͵���ֵ                dValue =100.00; 
	* ����ֵ:  �ɹ�����---TRUE,    ʧ�ܷ���--FALSE
	* ��  ʷ:  1.0 ���� 2013-11-4   create
	-----------------------------------------------------------------------------*/
	BOOL  XlsGetCellValue(const CString &addr, double &dValue);
	BOOL  XlsGetCellValue(const CString &addr,  CString& strValue);
	
	//�����Ԫ������
	void XlsClearRangeContens(const CString& strCell);
	
	/*----------------------------------------------------------------------------*
	* ��  ��:  �Զ��������
	* ��  ��:  [in]  strRow                          
	*          [in]  strCol      
	*          [in]  strValue        ������ֵ 
	*          [in]  xlsAtuoFillType �����ʽ
	* ��  ʷ:  1.0 ���� 2013-11-7 16:36:12   create
	-----------------------------------------------------------------------------*/
	void XlsAutoFillCell(const CString &strRow, const CString &strCol,const CString& strValue,
		XlAutoFillType xlsAtuoFillType = xlFillDefault);

	/*----------------------------------------------------------------------------*
	* ��  ��:  ����Ȳ�����
	* ��  ��:  [in]  strRow                          
	*          [in]  strCol      
	*          [in]  strFirstValue  �Ȳ����е�һ��  eg �� p1
	*          [in]  strSecValue    �Ȳ����еڶ���        p2
	*          [in]  xlsAtuoFillType �����ʽ
	* ��  ʷ:  1.0 ���� 2013-11-7 16:36:12   create
	-----------------------------------------------------------------------------*/
	void XlsAutoFillArithmeticSequenceCell(const CString &strRow, const CString &strCol,const CString& strFirstValue,const CString& strSecValue,
		XlAutoFillType xlsAtuoFillType = xlFillDefault);

	void XlsAutoFillNoArithSequenceCell(const CString &strFirstRow, const CString &strSecRow, const CString &strCol,double dFirtValue,double dSecValue,
		XlAutoFillType xlsAtuoFillType = xlFillDefault);

	//�˳�
	void XlsQuit1();
	void XlsQuit();

	//ǿ�ƹر�EXCLE.exe����
	void XlsTerminateExcelProcess(CString strProName = _T("EXCEL.EXE"));

	/*----------------------------------------------------------------------------*
	* ��  ��:  �ϲ���Ԫ��
	* ��  ��:  [in]  iRow  ��n�е�Ԫ��
	*          [in]  iCol  ��n�е�Ԫ��
	*          [in]  iMergeCellRow  Ҫ�ϲ�������
	*          [in]  iMergeCellCol  Ҫ�ϲ�������
	* ����ֵ:  
	* ��  ʷ:  1.0 ���� 2013-11-1   create
	-----------------------------------------------------------------------------*/
	void XlsMergeCells(int &iRow, int &iCol,int &iMergeCellRow,int &iMergeCellCol);
	
	void XlsGetLNum();                               //�õ�ģ�����������
	void XlsGetLNum_dai();
	int  XlsGetStep(CStringArray &nameArray);	     //�ж�ģ���Ƿ��кϲ��У��з���2���޷���1
	void XlsSetPage(int pageNum);				    //���ݸ�����ҳ������EXCELģ���и���ҳ��
	
	//Sheet ��������
	void XlsAddSheet(const CString &strSheetName);       //����sheet

	//����sheet������(������sheet)
	void XlsSetSheetName(long iIndex/*in (iIndex ȡֵ >= 1)*/, const CString &strSheetName/*in*/);

	 //������һ��sheet  ���ظ�sheet���� 
	long XlsActivateNextSheet();    

	/*
	    ע�⣺������������ɶ�ʹ��   XlsGetActiveWorkBookSheets(); 
	*                                XlsReleaseActiveBookSheets();
	*
	*/
	CWorksheets  XlsGetActiveWorkBookSheets();           //���ص�ǰ���������sheets �ӿ�  
	void  XlsReleaseActiveBookSheets();                  //�ͷŵ�ǰ���������sheets �ӿ�

	void XlsDelSheet(const CString &strSheetName);      //ɾ��sheet
	void XlsDelSheet(int Pos);	
	long XlsGetSheetCount();							//�õ�sheet������
	CString XlsGetSheetName(int iIndex);               //�õ�sheet������

	enum CellHorizonAlignment
	{
		xlCenter = -4108,
		xlLeft   = -4131,
		xlRight  = -4152,
		xlFill   = 5,
		xlJustify= -4130,
		xlCenterAcrossSelection = 7,
		xlDistributed= -4117
	};
	enum CellVerticalAlignment
	{
       xlTop = -4160,
	   xlBottom = -4107,
	   xlvJustify= -4130,
	   xlvDistributed= -4117
	};
	
	//���õ�Ԫ����뷽ʽ
	void XlsSetCellAlignment(int row, int colum, CellHorizonAlignment horizon, CellVerticalAlignment vertical);
	void XlsSetCellAlignment(const CString& cellRow, const CString& cellCol,  CellHorizonAlignment horizon, CellVerticalAlignment vertical);
	
	//����ͼƬ
	void XlsInsertPictrue(const CString &strPicPath,float xPos, float yPos, float picWidth, float picHeight);
	 
	//Excle��������
	BOOL XlsSetVisible(BOOL bIsVisible);
	void XlsSaveAs(const CString &fileName,XlFileFormat NewFileFormat = xlExcel5);

	void XlsSetColumnWidth(CString strRow/*_variant_t("B:B")*/, double width);   //���ñ����п��
	void XlsSetRowHeight(int Row, double height);
	void XlsCloseAlert();								//�رո����ļ�ʱ�ľ���


	// ���ļ�
	boost::optional<CWorkbook> OpenFile(const CString& fileName);
	// �½��ļ�
	boost::optional<CWorkbook> NewFile();
	// �����ļ�
	bool SaveAs(CWorkbook& workBook, const CString& fileName);
	// �ر�CWorkbook
	void Close(CWorkbook& workBook);
	// ��ù������й�����ĸ���
	int GetSheetCount(CWorkbook& workBook);
	// �ڹ������н�β�����һ����������������
	int AddSheetBack(CWorkbook& workBook);
	// ���ù�����ı�����
	void SetSheetName(CWorkbook& workBook, int nSheetIndex, const CString& name);
	// �ӹ�����1�и���ָ����������������ݵ�������2��
	bool CopySheetContents(CWorkbook& srcBook, int nSrcSheet, CWorkbook& destBook, int nDestSheet);
	// �������
	void ActiveSheet(CWorkbook& workBook, int nSheetIndex);
	// ���õ�Ԫ������
	void SetCellText(CWorkbook& workBook, int nSheetIndex, int row, int column, const CString& text);
	// ��úϲ���Ԫ��ĸ���
	int GetMergeCellsCount(CWorkbook& workBook, int nSheetIndex, int row, int column);
	// ��������ָ����Ԫ������Ϊ��С�������
	void SetShrinkToFit(CWorkbook& workBook, int nSheetIndex, int row, int column);

protected:	
	void XlsErrInfo(_com_error &e);
	void XlsSetHW(double Height,double Width);
	BOOL xlsAppIsInit();  //�ж�Excel�Ƿ��ʼ���ɹ�

	//����ַ������Ƿ�Ϸ�  ���������п� B:B
	BOOL XlsCheckStrIsValue(CString &str/*in*/, CString &strResult /*out*/);

	//EXCEL 2007
	CApplication	m_xlsAppLication;
	CWorkbooks		m_xlsWorkBooks;
	CWorksheets     m_xlsWorkSheets;   //���浱ǰ����Ĺ�������sheets
	BOOL            m_bIsXlsProcRedu; //ֻ��һ��Excel����
	///////////
private:
	CWorkbook       m_xlsWorkBook;  //��ǰ����Ĺ�����
};

#endif   //RLOPER_EXCEL_H
