
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

	//workBook操作
	//激活下一个工作簿 返回其索引ID
	long    XlsActivateNextWorkBook();
	CString XlsGetWorkBookName(int iIndex);
	int     XlsGetWorkBookIndex(const CString& strWorkName);
	int     XlsGetWorkBookCout();
	void    XlsDelWorkBook(long iIndex);
	void    XlsDelWorkBook(const CString& strWorkName);

	/*
	获取和释放必须成对出现
	*/
	CWorkbook XlsGetActiveBook();
	void    XlsReleaseActiveBook();

	int  XlsCloseBook(CWorkbook &workBook);			//close workbook
	int  XlsCloseBook();

	//插入一行
	BOOL XlsAddLine(const CString& v_hBegin, const CString& v_hEnd, XlInsertFormatOrigin xlsInsertFormat = xlFormatFromRightOrBelow);
	BOOL XmlGetRangeWidthHeight(RangePtr v_hRangePtr,double &v_dWidth,double &v_dHeight);
	
	void XlsRunMacro(const CString &macro);//运行宏
	
	BOOL  XlsGetMacroName(CStringArray &arrMarcoName);//得到当前Excel中存在的宏名字
	
	void XlsAddMacroFromFlie(CString &strMacroFilePath);//从文件中加载VBA宏
	
	enum XlsMouseCur
	{
		xlIBeamCur = 3,
		xlDefaultCur = -4143,
		xlNorthwestArrowCur = 1,
		xlWaitCur = 2
	};
   void XlsSetCursor(XlsMouseCur nflag);

	//创建XLs文件
	/*
	创建Excel文件有两种方式  
	1、从表格模板*.xlt中创建
	2、有程序自动创建
	二者是否可以同时初始化？？？
	*/

	/*----------------------------------------------------------------------------*
	* 描  述:  从表格模板创建XLs文件
	* 参  数:  	[in]  strXltFileName  xlt文件全路径
	*           [in]  bIsShowExcel    是否显示创建成功的Excel文件             
	* 返回值:   成功返回 -- TRUE      失败返回 -- FALSE
	* 历  史:  1.0 姜东 2013-10-31   create
	-----------------------------------------------------------------------------*/
	BOOL CreateXlsFromXltFile(const CString &strXltFileName, BOOL bIsShowExcel = TRUE);
	int CreateXls(BOOL bIsShowExcel = TRUE);
	int CreateXlsNoInfo();

    //打开和新建操作 
	int  XlsOpen(const CString &strOpenFileName, const CString &prjName);

	int  XlsName2Add(CString name,  CString &col,UINT &row);

	/*----------------------------------------------------------------------------*
	* 描  述:  向一组单元格中插入数据
	* 参  数:  	[in]  Row  第n行
	*           [in]  Col  第n列
	*           [in]  strValue  插入的字符串数据
	* 返回值:  
	* 历  史:  1.0 姜东 2013-10-31   create
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
	* 描  述:   绘制边框
	* 参  数:  	[in]  Row  第n行
	*           [in]  Col  第n列
	*           [in]  xlLineStyle  边框线型样式
	*           [in]  XlsWeight     边框的宽度
	*           [in]  coloIndex     颜色索引 最好设置在（0-100）之间
	* 返回值:  
	* 历  史:  1.0 姜东 2013-10-31   create
	-----------------------------------------------------------------------------*/
	//绘制内边框
	void XlsDrawInBorders(const CString &Row, const CString &Col,XlBordersLineStyle xlLineStyle, 
		XlBorderWeight XlsWeight,long coloIndex/*0-100*/);

	//绘制外边框
	void XlsDrawOutBorders(const CString &Row, const CString &Col,XlBordersLineStyle xlLineStyle, 
		XlBorderWeight XlsWeight,long coloIndex/*0-100*/);
		
    /*
	enum XlBordersIndex
	{
	xlInsideHorizontal = 12,
	xlInsideVertical = 11,
	xlDiagonalDown = 5, // 交叉
	xlDiagonalUp = 6,   //  '/'
	xlEdgeBottom = 9,
	xlEdgeLeft = 7,
	xlEdgeRight = 10,
	xlEdgeTop = 8
	};
	*/
	//绘制单个单元格边框样式
	void XlsDrawCellBorders(const CString &Cell, XlBordersLineStyle xlLineStyle, 
		XlBorderWeight XlsWeight,long coloIndex/*0-100*/, XlBordersIndex xlBordIndex);

	/*----------------------------------------------------------------------------*
	* 描  述:  向单个单元格中插入数据
	* 参  数:  [in]  Cell  第n个单元格
	*          [in]  strValue  插入的字符串数据
	* 返回值:  
	* 历  史:  1.0 姜东 2013-10-31   create
	-----------------------------------------------------------------------------*/
	void XlsSetCellData(const CString &cell, double dValue);
	void XlsSetCellData(const CString &cell, const CString& strValue);

	/*----------------------------------------------------------------------------*
	* 描  述:  对于给定单元格地址，double型的数值
	* 参  数:  [in]  addr        第n个单元格                     eg: addr = A1  其值是 100.0
	*          [out] dValue      返回double类型的数值                dValue =100.00; 
	* 返回值:  成功返回---TRUE,    失败返回--FALSE
	* 历  史:  1.0 姜东 2013-11-4   create
	-----------------------------------------------------------------------------*/
	BOOL  XlsGetCellValue(const CString &addr, double &dValue);
	BOOL  XlsGetCellValue(const CString &addr,  CString& strValue);
	
	//清除单元格内容
	void XlsClearRangeContens(const CString& strCell);
	
	/*----------------------------------------------------------------------------*
	* 描  述:  自动填充内容
	* 参  数:  [in]  strRow                          
	*          [in]  strCol      
	*          [in]  strValue        填充的数值 
	*          [in]  xlsAtuoFillType 填充样式
	* 历  史:  1.0 姜东 2013-11-7 16:36:12   create
	-----------------------------------------------------------------------------*/
	void XlsAutoFillCell(const CString &strRow, const CString &strCol,const CString& strValue,
		XlAutoFillType xlsAtuoFillType = xlFillDefault);

	/*----------------------------------------------------------------------------*
	* 描  述:  插入等差数列
	* 参  数:  [in]  strRow                          
	*          [in]  strCol      
	*          [in]  strFirstValue  等差数列第一项  eg ： p1
	*          [in]  strSecValue    等差数列第二项        p2
	*          [in]  xlsAtuoFillType 填充样式
	* 历  史:  1.0 姜东 2013-11-7 16:36:12   create
	-----------------------------------------------------------------------------*/
	void XlsAutoFillArithmeticSequenceCell(const CString &strRow, const CString &strCol,const CString& strFirstValue,const CString& strSecValue,
		XlAutoFillType xlsAtuoFillType = xlFillDefault);

	void XlsAutoFillNoArithSequenceCell(const CString &strFirstRow, const CString &strSecRow, const CString &strCol,double dFirtValue,double dSecValue,
		XlAutoFillType xlsAtuoFillType = xlFillDefault);

	//退出
	void XlsQuit1();
	void XlsQuit();

	//强制关闭EXCLE.exe进程
	void XlsTerminateExcelProcess(CString strProName = _T("EXCEL.EXE"));

	/*----------------------------------------------------------------------------*
	* 描  述:  合并单元格
	* 参  数:  [in]  iRow  第n行单元格
	*          [in]  iCol  第n列单元格
	*          [in]  iMergeCellRow  要合并的行数
	*          [in]  iMergeCellCol  要合并的列数
	* 返回值:  
	* 历  史:  1.0 姜东 2013-11-1   create
	-----------------------------------------------------------------------------*/
	void XlsMergeCells(int &iRow, int &iCol,int &iMergeCellRow,int &iMergeCellCol);
	
	void XlsGetLNum();                               //得到模板的数据行数
	void XlsGetLNum_dai();
	int  XlsGetStep(CStringArray &nameArray);	     //判断模板是否有合并行，有返回2，无返回1
	void XlsSetPage(int pageNum);				    //根据给定的页数，在EXCEL模板中复制页数
	
	//Sheet 操作函数
	void XlsAddSheet(const CString &strSheetName);       //增加sheet

	//设置sheet的名字(重命名sheet)
	void XlsSetSheetName(long iIndex/*in (iIndex 取值 >= 1)*/, const CString &strSheetName/*in*/);

	 //激活下一个sheet  返回该sheet索引 
	long XlsActivateNextSheet();    

	/*
	    注意：两个函数必须成对使用   XlsGetActiveWorkBookSheets(); 
	*                                XlsReleaseActiveBookSheets();
	*
	*/
	CWorksheets  XlsGetActiveWorkBookSheets();           //返回当前激活工作簿的sheets 接口  
	void  XlsReleaseActiveBookSheets();                  //释放当前激活工作簿的sheets 接口

	void XlsDelSheet(const CString &strSheetName);      //删除sheet
	void XlsDelSheet(int Pos);	
	long XlsGetSheetCount();							//得到sheet的数量
	CString XlsGetSheetName(int iIndex);               //得到sheet的名字

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
	
	//设置单元格对齐方式
	void XlsSetCellAlignment(int row, int colum, CellHorizonAlignment horizon, CellVerticalAlignment vertical);
	void XlsSetCellAlignment(const CString& cellRow, const CString& cellCol,  CellHorizonAlignment horizon, CellVerticalAlignment vertical);
	
	//插入图片
	void XlsInsertPictrue(const CString &strPicPath,float xPos, float yPos, float picWidth, float picHeight);
	 
	//Excle操作函数
	BOOL XlsSetVisible(BOOL bIsVisible);
	void XlsSaveAs(const CString &fileName,XlFileFormat NewFileFormat = xlExcel5);

	void XlsSetColumnWidth(CString strRow/*_variant_t("B:B")*/, double width);   //设置表格的列宽度
	void XlsSetRowHeight(int Row, double height);
	void XlsCloseAlert();								//关闭覆盖文件时的警告


	// 打开文件
	boost::optional<CWorkbook> OpenFile(const CString& fileName);
	// 新建文件
	boost::optional<CWorkbook> NewFile();
	// 保存文件
	bool SaveAs(CWorkbook& workBook, const CString& fileName);
	// 关闭CWorkbook
	void Close(CWorkbook& workBook);
	// 获得工作簿中工作表的个数
	int GetSheetCount(CWorkbook& workBook);
	// 在工作簿中结尾处添加一个工作表，返回索引
	int AddSheetBack(CWorkbook& workBook);
	// 设置工作表的表名称
	void SetSheetName(CWorkbook& workBook, int nSheetIndex, const CString& name);
	// 从工作簿1中复制指定索引工作表的内容到工作簿2中
	bool CopySheetContents(CWorkbook& srcBook, int nSrcSheet, CWorkbook& destBook, int nDestSheet);
	// 激活工作表
	void ActiveSheet(CWorkbook& workBook, int nSheetIndex);
	// 设置单元格内容
	void SetCellText(CWorkbook& workBook, int nSheetIndex, int row, int column, const CString& text);
	// 获得合并单元格的个数
	int GetMergeCellsCount(CWorkbook& workBook, int nSheetIndex, int row, int column);
	// 将工作表指定单元格设置为缩小字体填充
	void SetShrinkToFit(CWorkbook& workBook, int nSheetIndex, int row, int column);

protected:	
	void XlsErrInfo(_com_error &e);
	void XlsSetHW(double Height,double Width);
	BOOL xlsAppIsInit();  //判断Excel是否初始化成功

	//检查字符输入是否合法  用于设置列宽 B:B
	BOOL XlsCheckStrIsValue(CString &str/*in*/, CString &strResult /*out*/);

	//EXCEL 2007
	CApplication	m_xlsAppLication;
	CWorkbooks		m_xlsWorkBooks;
	CWorksheets     m_xlsWorkSheets;   //保存当前激活的工作簿的sheets
	BOOL            m_bIsXlsProcRedu; //只有一个Excel进程
	///////////
private:
	CWorkbook       m_xlsWorkBook;  //当前激活的工作簿
};

#endif   //RLOPER_EXCEL_H
