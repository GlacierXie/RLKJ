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
 \\by xzx ���ñ���
*/
#pragma once

#include <vector>
#include "RLOperExcel.h"
using namespace std;

typedef vector <vector <CString>>CTableDataArry;

struct ExcelTemplateInfo
{
	int nRow;       // ��ʼ��Ԫ����
	int nColumn;    // ��ʼ��Ԫ����
	int nRowCount;  // ����
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
		BOOL bStartEnt; //��㴦�Ƿ��б���
		BOOL bEndNextPt;  //�յ��Ƿ��йյ�
		BOOL bEndNextPt2;
	};
	struct straigthInsertData  //ֱ�߱��� ��������
	{
		BackEdgeTypeData backData;
		AcGeMatrix3d mat;
		AcDbObjectId entBlockId;
	};

	struct breakInsertData  //�յ㱳���������
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

	//��ʼ������������
	static BOOL InitBackEdge();

	//�Զ����ñ���
	void drawBackEdge();

	//ʶ��ѡ��ģ��
	void AutodrawBackEdge();

	//ʶ��ѡ��ǽ
	void AutodrawBackEdgeWall();

	//����EXCEL�嵥
	void CreatExcel();

	void rllock();

private:

	BOOL getSSFirstIdArry(AcDbObjectIdArray& entIdArry, BOOL bTip = FALSE);

	//��ȡ�����嵥����
	BOOL getBackEdgeExcelData(CTableDataArry& tableArry);

	//��ȡ���ģ���ļ�·��
	BOOL getExcelTempPath(CString& strPath);

	//��ȡ����ı���ļ�·��
	BOOL getExcelSavePath(CString& strPath);

	BOOL getBackEdgeExcelData(std::map<CString, int>&mapBackEdgeNos, const std::map<AcDbObjectId, CString>& mapIdToBackType,
		AcArray<CString>& strBackTypeArry,AcDbEntity* pEnt);

	BOOL getPolyPathLine(std::set<AcGePolyline2d *>& polygons,AcDbObjectIdArray& entDwgIdArry);

	void updataWallPts(std::set<AcGePolyline2d *>& polygons);

	void updataWallPts(AcArray<AcGePoint3dArray>& pathPtArrys);

	//�ͷ��ڴ�
	void releaseVoidPtr(AcGeVoidPointerArray& ptIines);

	//ը�������
	AcDbObjectIdArray exPlodePline(AcDbObjectIdArray& entDwgIds);

	//ѡȡ���ñ���Ķ����
	bool selectAutoLines(AcDbObjectIdArray& entIdArry,AcDbObjectIdArray& entDwgIds);

	void selectAutoLines(AcDbObjectIdArray& entIdArry, AcDbEntity* pEnt);

	//�ϲ���ά�����
	std::set<AcGePolyline2d *> mergePolyline(std::set<AcGePolyline2d *>& polygons);

	//����������Χ��
	AcDbObjectIdArray drawPathLines(const std::set<AcGePolyline2d *>& polygons);

	//ѡȡ���˵��߶�
	AcArray<AcGePoint3dArray> delBackLines(const AcDbObjectIdArray&entCountIds);

	//ѡȡ�����
	void getBackPathLines(AcArray<AcGePoint3dArray>& pathPtArrys, const AcArray<AcGePoint3dArray>&pathDelPtArrys, AcGePolyline2d *polygon);

	//��ȡ������ʽ��
	static AcDbObjectId getTextStyleId(const ACHAR* name);

	//��ʼ������������
	static void InitBackEdgeData(const BackEdgeTypeData& data, AcDbObjectId& entId, AcDbBlockTable*& pBlockTb);

	//ѡȡ���ñ���Ķ����
	bool selectPLines(AcDbObjectIdArray& entIdArry);

	//ѡȡ����
	bool selectBackEdges(AcDbObjectIdArray& entIdArry);

	//�����Ǵ�ֱ�Ķ����
	void dealWithErrorPt(AcArray<AcGePoint3dArray>& pathPtArrys);

	//���ñ���
	void drawBackEdge(const AcArray<AcGePoint3dArray>& pathPtArrys);

	// 	��ȡ�йյ�ı��� �Լ����ӵ�ֱ�߱���
	BOOL getBreakBackDatas(std::vector<straigthInsertData>& vecStraigthData, std::vector<breakInsertData>& vecBreakData,
		const AcGePoint3dArray& ptArry, BOOL bLoop);

	//��ȡ��·�ĵ�һ���յ㱳��
	BOOL getLoopFirstBreakBackData(std::vector<breakInsertData>& vecBreakData, const AcGePoint3dArray& ptArry);

	//��ȡ��·�ĵ�һ���յ㱳��
	BOOL getLoopFirstBreakBackData(breakInsertData&breakData , std::map<BackEdgeTypeData, AcDbObjectId>::iterator& iter, const AcGePoint3dArray& ptArry);

	//��ȡû�л�·�ĵ�һ���յ㱳��
	BOOL getNoLoopFirstBreakBackData(breakInsertData&breakData , std::map<BackEdgeTypeData, AcDbObjectId>::iterator& iter ,const AcGePoint3dArray& ptArry);

	//��ȡû�л�·�ĵ�һ���յ㱳��
	BOOL getNoLoopFirstBreakBackData(std::vector<breakInsertData>& vecBreakData, const AcGePoint3dArray& ptArry);

	//��ȡ��N���յ㱳��
	BOOL getBreakBackData(CMapIndexStargthData& mapStarData,std::vector<breakInsertData>& vecBreakData, 
		const AcGePoint3dArray& ptArry, int iIndex,BOOL bLoop);

	//��ȡ��N���յ㱳��
	BOOL getBreakBackData(breakInsertData&breakData, CMapIndexStargthData& mapStarData, std::map<BackEdgeTypeData, AcDbObjectId>::iterator& iter,
		const std::vector<breakInsertData>& vecBreakData, const AcGePoint3dArray& ptArry, int iIndex, BOOL bLoop);

	BOOL getBreakBackData(breakInsertData&breakData, CMapIndexStargthData& mapStarData, const std::vector<breakInsertData>& vecBreakData, 
		const AcGePoint3dArray& ptArry, int iIndex, BOOL bLoop);

	//��ȡ�յ�����
	void getBreakData(breakInsertData&breakData, const AcGePoint3dArray& ptArry,
		const std::map<BackEdgeTypeData, AcDbObjectId>::iterator& iter,
		int iStartIndex, int iIndex, int iEndIndex, BOOL bAdjust, BOOL bNegeit);

	//����û�л�·�����һ���յ� ���ص������ֵ ����Ϊ�����Ľ���
	int delTheLastPtNoLoop(std::vector<breakInsertData>& vecBreakData, std::vector<straigthInsertData>& vecStraigthData,const AcGePoint3dArray& ptArry);

	//��������һ���յ���Ϣ
	BOOL getBreakBackDataHasEnd(breakInsertData&breakData, CMapIndexStargthData& mapStarData, std::vector<breakInsertData>& vecBreakData, 
		BOOL& bGetData, const AcGePoint3dArray& ptArry, int iIndex, BOOL bLoop);

	BOOL getBreakBackDataNoMaterEnd(breakInsertData&breakData, CMapIndexStargthData& mapStarData, std::vector<breakInsertData>& vecBreakData,
		BOOL& bGetData, const AcGePoint3dArray& ptArry, int iIndex);

	//��ȡֱ�߱���
	BOOL getStraightBackDatas(std::vector<straigthInsertData>& vecStraigthData, const AcGePoint3dArray& ptArry);

	//��ȡֱ�߱���
	BOOL getHasStartStraightBackDatas(std::vector<straigthInsertData>& vecStraigthData, const backPtData& pathPtData,
		AcGePoint3d& ptInsert, double dInLen, BOOL bBlue, BOOL bBeginStart = TRUE);

	//��ȡֱ��������ɫ����
	BOOL getBlockIdConBlueStraigth(AcDbObjectId& blockId, BackEdgeTypeData& backEdgedata, AcGePoint3d& ptInsert,
		AcDbObjectId& blockId2, BackEdgeTypeData& backEdgedata2, AcGePoint3d& ptInsert2, BOOL& bNext, const backPtData& pathPtData, double dInLen);

	//��ȡֱ��������ɫ����
	BOOL getBlockIdConGreenStraigth(AcDbObjectId& blockId, BackEdgeTypeData& backEdgedata, AcGePoint3d& ptInsert,
		AcDbObjectId& blockId2, BackEdgeTypeData& backEdgedata2, AcGePoint3d& ptInsert2, BOOL& bNext, const backPtData& pathPtData, double dInLen);


	//��ȡû��ǰ��Ҳû��ĩ�˵ı���ID
	AcDbObjectId getBlockIdNoStartNoEnd(BackEdgeTypeData& backEdgedata, AcGePoint3d& ptInsert, BOOL& bNext, const AcGePoint3d& ptStart, const AcGePoint3d&ptEnd);

	//��ȡҪ����ɫ-��ɫ-��ɫ����ID
	BOOL getBlockIdConnect(AcDbObjectId& blockId, BackEdgeTypeData& backEdgedata1, AcGePoint3d& ptInsert, AcDbObjectId& blockId2, BackEdgeTypeData& backEdgedata2, AcGePoint3d& ptInsert2,
		AcDbObjectId& blockId3, BackEdgeTypeData& backEdgedata3, AcGePoint3d& ptInsert3, const AcGeVector3d& vecInesrt, const AcGePoint3d& ptStar, const AcGePoint3d&ptEnd,
		double dInLen1, double dInlen2);

	//��ȡҪ����ɫ-��ɫ����ID
	BOOL getBlockIdConnect(AcDbObjectId& blockId, BackEdgeTypeData& backEdgedata1, AcGePoint3d& ptInsert, AcDbObjectId& blockId2, BackEdgeTypeData& backEdgedata2, AcGePoint3d& ptInsert2,
		const AcGeVector3d& vecInesrt, const AcGePoint3d& ptStar, const AcGePoint3d&ptEnd,double dInLen1, double dInlen2);

	//��ӿ���յ����ݿ�
	BOOL addBlockResBlockReference(const straigthInsertData& insertData);

	BOOL addBlockResBlockReference(const breakInsertData& insertData);
	//��ȡֱ�α��������
	void getMapStraightIter(std::vector<std::map<BackEdgeTypeData, AcDbObjectId>::iterator>& vecIter);

	//�������˱���
	BOOL connectBackEdge(std::vector<straigthInsertData>& vecStraigthData, const AcGeVector3d& vecInesrt, const AcGePoint3d& ptStart, 
		const AcGePoint3d&ptEnd, double dInLen1, double dInlen2, BOOL bBlue);

	//��ȡ�������Ӷβ�������
	BOOL getInlenBetweenTwoBackEdge(double& dInResLen1, double& dInReslen2, double dTotalLen, double dInLen1, double dInlen2, double dInLen3);

	//��ȡ�������Ӷβ�������
	BOOL getInlenBetweenThreeBackEdge(double& dInResLen1, double& dInReslen2, double& dInReslen3, double dTotalLen, double dInLen1, double dInlen2, double dInLen3,double dInLen4);

	// �ӵ�ǰҳ���ģ���л�ñ����Ϣ����ʼ��Ԫ��������
	bool GetExcelTemplateInfo(CWorkbook& workBook, int nSheetIndex, vector<ExcelTemplateInfo>& vecExcelTableInfos);

	//��ӵ����ݿ�
	bool AddToModelSpace(AcDbEntity* pEnt, AcDbObjectId& objId, bool bClose = true, AcDbDatabase* pDb = NULL);

	AcGeVector3d getVec3D(const AcGeVector2d& vec);

	AcGeVector2d getVec2D(const AcGeVector3d& vec);

	AcGePoint3d getPoint3D(const AcGePoint2d& vec);

	AcGePoint2d getPoint2D(const AcGePoint3d& vec);

	//����ͼ��
	AcDbObjectId createLayer(CString strLayerName, int layercolorIndex, bool bOffLayer = false, AcDbDatabase *pDb = NULL);

   //ɾ��IDs
	void earseEnts(const AcDbObjectIdArray& entIdArry,BOOL bEarse = TRUE);

	//��ʾ����
	void ShowEnt(const AcDbObjectIdArray& arrIds, BOOL bIsShow = TRUE);

	//��ȡͼ����ɫ����
	int getColorIndexByLayer(CString lyname, AcDbDatabase *pDb = NULL);

	//��ȡʵ����ɫ����
	int getEntiColorIndex(AcDbEntity* pEnt);
};

