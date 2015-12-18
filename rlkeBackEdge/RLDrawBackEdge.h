/*!
 * \file RLDrawBackEdge.h
 * \date 2015/09/18 15:37
 *
 * \author Administrator
 * Contact: user@company.com
 *
 * \brief 
 *
 * TODO: long description
 *
 * \note
 \\by xzx 布置背楞
*/
#pragma once

#include <vector>
#include "RLOperExcel.h"
using namespace std;

typedef vector <vector <CString>>CTableDataArry;

struct ExcelTemplateInfo
{
	int nRow;       // 起始单元格行
	int nColumn;    // 起始单元格列
	int nRowCount;  // 行数
};

struct BackEdgeTypeData
{
	BackEdgeTypeData()
	{
		m_dWith = 90;
		m_dLen = 0.0;
		m_dLen2 = 0.0;
		m_iColorIndex = 4;
		m_bBlue = TRUE;
	}
	virtual bool operator < (const BackEdgeTypeData& other) const
	{
		if (m_dLen < other.m_dLen)
			return true;
		else if (m_dLen > other.m_dLen)
			return false;
		else if (m_dLen2 < other.m_dLen2)
			return true;
		else if (m_dLen2 > other.m_dLen2)
			return true;
		else if (m_bBlue&&m_bBlue != other.m_bBlue)
			return true;
		return false;
	}
	virtual bool operator >(const BackEdgeTypeData& other) const
	{
		if (m_dLen > other.m_dLen)
			return true;
		if (m_dLen2 > other.m_dLen2)
			return true;
		return false;
	}

	bool operator == (const BackEdgeTypeData& other) const
	{
		if ((this->m_strBackEdge == other.m_strBackEdge) && (fabs(this->m_dWith - other.m_dWith) < 0.00001)
			|| (fabs(this->m_dLen - other.m_dLen) < 0.00001) && (fabs(this->m_dLen2 - other.m_dLen2) < 0.00001)
			|| (this->m_iColorIndex == other.m_iColorIndex) && (this->m_bBlue == other.m_bBlue))
			return true;
		else
			return false;
	}

	CString m_strBackEdge;
	double m_dWith;
	double m_dLen;
	double m_dLen2;
	int m_iColorIndex;
	BOOL m_bBlue;
};

class CRLDrawBackEdge
{

private:
	struct backPtData
	{
		backPtData()
		{
			bStartEnt = FALSE;
			bEndNextPt = FALSE;
			bEndNextPt2 = FALSE;
		}
		AcGePoint3d ptStart;
		AcGePoint3d ptEnd;
		AcGePoint3d ptNextEndPt;
		BOOL bStartEnt; //起点处是否有背楞
		BOOL bEndNextPt;  //终点是否有拐点
		BOOL bEndNextPt2;
	};
	struct straigthInsertData  //直线背楞 插入数据
	{
		BackEdgeTypeData backData;
		AcGeMatrix3d mat;
		AcDbObjectId entBlockId;
	};

	struct breakInsertData  //拐点背楞插入类型
	{
		breakInsertData()
		{
			dStartInLen = dNextInLen = dStartLen = dNextLen = 0.0;
		}
		BackEdgeTypeData backData;
		AcDbObjectId entBlockId;
		AcGePoint3d ptZero;
		double dStartLen;
		double dNextLen;

		double dStartInLen;
		double dNextInLen;
		AcGeMatrix3d mat;
	};

	typedef std::map<int, std::vector<straigthInsertData>> CMapIndexStargthData;

public:
	CRLDrawBackEdge();
	~CRLDrawBackEdge();

public:

	//初始化背楞对象编组
	static BOOL InitBackEdge();

	//自动布置背楞
	void drawBackEdge();

	//识别选择模板
	void AutodrawBackEdge();

	//识别选择墙
	void AutodrawBackEdgeWall();

	//生成EXCEL清单
	void CreatExcel();

	void rllock();

private:

	BOOL getSSFirstIdArry(AcDbObjectIdArray& entIdArry, BOOL bTip = FALSE);

	//获取背楞清单数据
	BOOL getBackEdgeExcelData(CTableDataArry& tableArry);

	//获取表格模板文件路径
	BOOL getExcelTempPath(CString& strPath);

	//获取保存的表格文件路径
	BOOL getExcelSavePath(CString& strPath);

	BOOL getBackEdgeExcelData(std::map<CString, int>&mapBackEdgeNos, const std::map<AcDbObjectId, CString>& mapIdToBackType,
		AcArray<CString>& strBackTypeArry,AcDbEntity* pEnt);

	BOOL getPolyPathLine(std::set<AcGePolyline2d *>& polygons,AcDbObjectIdArray& entDwgIdArry);

	void updataWallPts(std::set<AcGePolyline2d *>& polygons);

	void updataWallPts(AcArray<AcGePoint3dArray>& pathPtArrys);

	//释放内存
	void releaseVoidPtr(AcGeVoidPointerArray& ptIines);

	//炸开多段线
	AcDbObjectIdArray exPlodePline(AcDbObjectIdArray& entDwgIds);

	//选取布置背楞的多段线
	bool selectAutoLines(AcDbObjectIdArray& entIdArry,AcDbObjectIdArray& entDwgIds);

	void selectAutoLines(AcDbObjectIdArray& entIdArry, AcDbEntity* pEnt);

	//合并二维多段线
	std::set<AcGePolyline2d *> mergePolyline(std::set<AcGePolyline2d *>& polygons);

	//绘制轮廓包围线
	AcDbObjectIdArray drawPathLines(const std::set<AcGePolyline2d *>& polygons);

	//选取过滤的线段
	AcArray<AcGePoint3dArray> delBackLines(const AcDbObjectIdArray&entCountIds);

	//选取多段线
	void getBackPathLines(AcArray<AcGePoint3dArray>& pathPtArrys, const AcArray<AcGePoint3dArray>&pathDelPtArrys, AcGePolyline2d *polygon);

	//获取文字样式表
	static AcDbObjectId getTextStyleId(const ACHAR* name);

	//初始化化背楞类型
	static void InitBackEdgeData(const BackEdgeTypeData& data, AcDbObjectId& entId, AcDbBlockTable*& pBlockTb);

	//选取布置背楞的多段线
	bool selectPLines(AcDbObjectIdArray& entIdArry);

	//选取背楞
	bool selectBackEdges(AcDbObjectIdArray& entIdArry);

	//处理不是垂直的多段线
	void dealWithErrorPt(AcArray<AcGePoint3dArray>& pathPtArrys);

	//布置背楞
	void drawBackEdge(const AcArray<AcGePoint3dArray>& pathPtArrys);

	// 	获取有拐点的背楞 以及连接的直线背楞
	BOOL getBreakBackDatas(std::vector<straigthInsertData>& vecStraigthData, std::vector<breakInsertData>& vecBreakData,
		const AcGePoint3dArray& ptArry, BOOL bLoop);

	//获取回路的第一个拐点背楞
	BOOL getLoopFirstBreakBackData(std::vector<breakInsertData>& vecBreakData, const AcGePoint3dArray& ptArry);

	//获取回路的第一个拐点背楞
	BOOL getLoopFirstBreakBackData(breakInsertData&breakData , std::map<BackEdgeTypeData, AcDbObjectId>::iterator& iter, const AcGePoint3dArray& ptArry);

	//获取没有回路的第一个拐点背楞
	BOOL getNoLoopFirstBreakBackData(breakInsertData&breakData , std::map<BackEdgeTypeData, AcDbObjectId>::iterator& iter ,const AcGePoint3dArray& ptArry);

	//获取没有回路的第一个拐点背楞
	BOOL getNoLoopFirstBreakBackData(std::vector<breakInsertData>& vecBreakData, const AcGePoint3dArray& ptArry);

	//获取第N个拐点背楞
	BOOL getBreakBackData(CMapIndexStargthData& mapStarData,std::vector<breakInsertData>& vecBreakData, 
		const AcGePoint3dArray& ptArry, int iIndex,BOOL bLoop);

	//获取第N个拐点背楞
	BOOL getBreakBackData(breakInsertData&breakData, CMapIndexStargthData& mapStarData, std::map<BackEdgeTypeData, AcDbObjectId>::iterator& iter,
		const std::vector<breakInsertData>& vecBreakData, const AcGePoint3dArray& ptArry, int iIndex, BOOL bLoop);

	BOOL getBreakBackData(breakInsertData&breakData, CMapIndexStargthData& mapStarData, const std::vector<breakInsertData>& vecBreakData, 
		const AcGePoint3dArray& ptArry, int iIndex, BOOL bLoop);

	//获取拐点数据
	void getBreakData(breakInsertData&breakData, const AcGePoint3dArray& ptArry,
		const std::map<BackEdgeTypeData, AcDbObjectId>::iterator& iter,
		int iStartIndex, int iIndex, int iEndIndex, BOOL bAdjust, BOOL bNegeit);

	//处理没有回路的最后一个拐点 返回点的索引值 索引为负数的结束
	int delTheLastPtNoLoop(std::vector<breakInsertData>& vecBreakData, std::vector<straigthInsertData>& vecStraigthData,const AcGePoint3dArray& ptArry);

	//处理非最后一个拐点信息
	BOOL getBreakBackDataHasEnd(breakInsertData&breakData, CMapIndexStargthData& mapStarData, std::vector<breakInsertData>& vecBreakData, 
		BOOL& bGetData, const AcGePoint3dArray& ptArry, int iIndex, BOOL bLoop);

	BOOL getBreakBackDataNoMaterEnd(breakInsertData&breakData, CMapIndexStargthData& mapStarData, std::vector<breakInsertData>& vecBreakData,
		BOOL& bGetData, const AcGePoint3dArray& ptArry, int iIndex);

	//获取直线背楞
	BOOL getStraightBackDatas(std::vector<straigthInsertData>& vecStraigthData, const AcGePoint3dArray& ptArry);

	//获取直线背楞
	BOOL getHasStartStraightBackDatas(std::vector<straigthInsertData>& vecStraigthData, const backPtData& pathPtData,
		AcGePoint3d& ptInsert, double dInLen, BOOL bBlue, BOOL bBeginStart = TRUE);

	//获取直线连接蓝色背楞
	BOOL getBlockIdConBlueStraigth(AcDbObjectId& blockId, BackEdgeTypeData& backEdgedata, AcGePoint3d& ptInsert,
		AcDbObjectId& blockId2, BackEdgeTypeData& backEdgedata2, AcGePoint3d& ptInsert2, BOOL& bNext, const backPtData& pathPtData, double dInLen);

	//获取直线连接绿色背楞
	BOOL getBlockIdConGreenStraigth(AcDbObjectId& blockId, BackEdgeTypeData& backEdgedata, AcGePoint3d& ptInsert,
		AcDbObjectId& blockId2, BackEdgeTypeData& backEdgedata2, AcGePoint3d& ptInsert2, BOOL& bNext, const backPtData& pathPtData, double dInLen);


	//获取没有前端也没有末端的背楞ID
	AcDbObjectId getBlockIdNoStartNoEnd(BackEdgeTypeData& backEdgedata, AcGePoint3d& ptInsert, BOOL& bNext, const AcGePoint3d& ptStart, const AcGePoint3d&ptEnd);

	//获取要求绿色-蓝色-绿色背楞ID
	BOOL getBlockIdConnect(AcDbObjectId& blockId, BackEdgeTypeData& backEdgedata1, AcGePoint3d& ptInsert, AcDbObjectId& blockId2, BackEdgeTypeData& backEdgedata2, AcGePoint3d& ptInsert2,
		AcDbObjectId& blockId3, BackEdgeTypeData& backEdgedata3, AcGePoint3d& ptInsert3, const AcGeVector3d& vecInesrt, const AcGePoint3d& ptStar, const AcGePoint3d&ptEnd,
		double dInLen1, double dInlen2);

	//获取要求蓝色-绿色背楞ID
	BOOL getBlockIdConnect(AcDbObjectId& blockId, BackEdgeTypeData& backEdgedata1, AcGePoint3d& ptInsert, AcDbObjectId& blockId2, BackEdgeTypeData& backEdgedata2, AcGePoint3d& ptInsert2,
		const AcGeVector3d& vecInesrt, const AcGePoint3d& ptStar, const AcGePoint3d&ptEnd,double dInLen1, double dInlen2);

	//添加快参照到数据库
	BOOL addBlockResBlockReference(const straigthInsertData& insertData);

	BOOL addBlockResBlockReference(const breakInsertData& insertData);
	//获取直段背楞迭代器
	void getMapStraightIter(std::vector<std::map<BackEdgeTypeData, AcDbObjectId>::iterator>& vecIter);

	//连接两端背楞
	BOOL connectBackEdge(std::vector<straigthInsertData>& vecStraigthData, const AcGeVector3d& vecInesrt, const AcGePoint3d& ptStart, 
		const AcGePoint3d&ptEnd, double dInLen1, double dInlen2, BOOL bBlue);

	//获取两个连接段插入的深度
	BOOL getInlenBetweenTwoBackEdge(double& dInResLen1, double& dInReslen2, double dTotalLen, double dInLen1, double dInlen2, double dInLen3);

	//获取三个连接段插入的深度
	BOOL getInlenBetweenThreeBackEdge(double& dInResLen1, double& dInReslen2, double& dInReslen3, double dTotalLen, double dInLen1, double dInlen2, double dInLen3,double dInLen4);

	// 从当前页表格模板中获得表格信息（起始单元格、行数）
	bool GetExcelTemplateInfo(CWorkbook& workBook, int nSheetIndex, vector<ExcelTemplateInfo>& vecExcelTableInfos);

	//添加到数据库
	bool AddToModelSpace(AcDbEntity* pEnt, AcDbObjectId& objId, bool bClose = true, AcDbDatabase* pDb = NULL);

	AcGeVector3d getVec3D(const AcGeVector2d& vec);

	AcGeVector2d getVec2D(const AcGeVector3d& vec);

	AcGePoint3d getPoint3D(const AcGePoint2d& vec);

	AcGePoint2d getPoint2D(const AcGePoint3d& vec);

	//创建图层
	AcDbObjectId createLayer(CString strLayerName, int layercolorIndex, bool bOffLayer = false, AcDbDatabase *pDb = NULL);

   //删除IDs
	void earseEnts(const AcDbObjectIdArray& entIdArry,BOOL bEarse = TRUE);

	//显示隐藏
	void ShowEnt(const AcDbObjectIdArray& arrIds, BOOL bIsShow = TRUE);

	//获取图层颜色索引
	int getColorIndexByLayer(CString lyname, AcDbDatabase *pDb = NULL);

	//获取实体颜色索引
	int getEntiColorIndex(AcDbEntity* pEnt);
};

