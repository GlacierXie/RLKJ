#include "StdAfx.h"
#include "RLDrawBackEdge.h"
#include "frgExtractTopologicalEntsFromLinesAlgm.h"
#include "frgPolygon2dRecognitionAlgm.h"

double g_dBlueWith = 90.0;
double g_dGreenWith = 80.0;
CString g_strBackEdge1 = _T("2500A");
CString g_strBackEdge2 = _T("1000A");
CString g_strBackEdge3 = _T("2000B");
CString g_strBackEdge4 = _T("1500B");
CString g_strBackEdge5 = _T("1000B");
CString g_strBackEdge6 = _T("500B");
CString g_strBackEdge7 = _T("500X350");
CString g_strBackEdge8 = _T("1000X500");
CString g_strBackEdge9 = _T("1000X1000");

CString g_strBackLayer = _T("背楞");
double g_singleMinLenout = 80.0; //单个突出最小距离距离
double g_groupMinLenout = 80.0;//组合突出最小距离
double g_minLenIn = 150.0;//连接插入最小距离
double g_standlLenout = 100.0;//突出标准距离
double g_sideMinLen= 100.0;//离背楞边界最短距离
double g_dBreakOutLen = 300.0; //L型突出允许的界限
double g_dAngleEqual = 2.0;
double g_dWallOutLen = 65.0;
#define PI 3.14159265358979323846

int g_bIsWall = FALSE;

#define  EPRES 0.001

std::map<BackEdgeTypeData, AcDbObjectId> g_mapStraightBackData;
std::map<BackEdgeTypeData, AcDbObjectId> g_mapTurnBackData;
std::map<AcDbObjectId, CString> g_mapIDToBlockText;

CRLDrawBackEdge::CRLDrawBackEdge()
{
	
}


CRLDrawBackEdge::~CRLDrawBackEdge()
{
	
}

void CRLDrawBackEdge::rllock()
{
	AcDbObjectIdArray entIdArry;
	if (!getSSFirstIdArry(entIdArry,TRUE))
	{
		ads_name ssText;
		acutPrintf(_T("请选择关闭实体"));
		int iRet = acedSSGet(NULL, NULL, NULL, NULL, ssText);
		if (RTNORM != iRet)
		{
			acedSSFree(ssText);
			return;
		}

		long ssLen;
		acedSSLength(ssText, &ssLen);
		for (int i = 0; i < ssLen; i++)
		{
			ads_name ent;
			acedSSName(ssText, i, ent);
			AcDbObjectId eld;
			acdbGetObjectId(eld, ent);
			entIdArry.append(eld);
			acedSSFree(ent);
		}
		acedSSFree(ssText);
	}
	AcDbObjectIdArray entLayerIdArry;
	int nLen = entIdArry.length();
	for (int i = 0; i < nLen; i++)
	{
		AcDbEntityPointer pEnt(entIdArry[i], AcDb::kForRead);
		if (pEnt.openStatus() != Acad::eOk || entIdArry.contains(pEnt->layerId()))
		   continue;
		entIdArry.append(pEnt->layerId());
		AcDbLayerTableRecord *pLayerTableRec = NULL;
		Acad::ErrorStatus es = acdbOpenObject(pLayerTableRec, pEnt->layerId(), AcDb::kForWrite);
		if (Acad::eOk == es)
		{
			pLayerTableRec->setIsOff(1);
			pLayerTableRec->close();
		}
	}
}

BOOL CRLDrawBackEdge::getSSFirstIdArry(AcDbObjectIdArray& entIdArry, BOOL bTip)
{
	entIdArry.removeAll();
	struct resbuf* rb;
	if (RTNORM != acedSSGetFirst(NULL, &rb))
		return FALSE;
	ads_name entname;
	ads_name_set(rb->resval.rlname, entname);
	long idLen = 0;
	acedSSLength(entname, &idLen);
	for (long i = 0; i < idLen; ++i)
	{
		ads_name name;
		acedSSName(entname, i, name);
		AcDbObjectId entId = AcDbObjectId::kNull;
		acdbGetObjectId(entId, name);
		if (entId.isValid() && !entIdArry.contains(entId))
			entIdArry.append(entId);
		acedSSFree(name);
	}
	acedSSFree(entname);
	acutRelRb(rb);
	if (entIdArry.isEmpty())
		return FALSE;
	if (bTip)
	{
		CString str;
		str.Format(_T("\n已选中%d个实体！"), entIdArry.length());
		acutPrintf(str);
	}
	return TRUE;
}

BOOL CRLDrawBackEdge::InitBackEdge()
{
	g_mapStraightBackData.clear();
	g_mapTurnBackData.clear();
	g_mapIDToBlockText.clear();
	AcDbDatabase* pDb = acdbCurDwg();
	if (pDb == NULL)
		return FALSE;

	AcDbBlockTable* pBlockTb = NULL;
	if (pDb->getBlockTable(pBlockTb, AcDb::kForWrite) != Acad::eOk)
		return FALSE;
	ASSERT(pBlockTb != NULL);

	AcDbObjectId entBackId = AcDbObjectId::kNull;

	BackEdgeTypeData backData;
	backData.m_strBackEdge = _T("2500A");
	backData.m_dLen = 2500;
	InitBackEdgeData(backData, entBackId, pBlockTb);
	g_mapStraightBackData[backData] = entBackId;
	g_mapIDToBlockText[entBackId] = backData.m_strBackEdge;

	backData.m_strBackEdge = _T("1000A");
	backData.m_dLen = 1000;
	InitBackEdgeData(backData, entBackId, pBlockTb);
	g_mapStraightBackData[backData] = entBackId;
	g_mapIDToBlockText[entBackId] = backData.m_strBackEdge;

	backData.m_strBackEdge = _T("2000B");
	backData.m_dLen = 2000;
	backData.m_bBlue = FALSE;
	backData.m_dWith = 80.0;
	backData.m_iColorIndex = 3;
	InitBackEdgeData(backData, entBackId, pBlockTb);
	g_mapStraightBackData[backData] = entBackId;
	g_mapIDToBlockText[entBackId] = backData.m_strBackEdge;

	backData.m_strBackEdge = _T("1500B");
	backData.m_dLen = 1500;
	backData.m_bBlue = FALSE;
	backData.m_dWith = 80.0;
	backData.m_iColorIndex = 3;
	InitBackEdgeData(backData, entBackId, pBlockTb);
	g_mapStraightBackData[backData] = entBackId;
	g_mapIDToBlockText[entBackId] = backData.m_strBackEdge;

	backData.m_strBackEdge = _T("1000B");
	backData.m_dLen = 1000;
	backData.m_bBlue = FALSE;
	backData.m_dWith = 80.0;
	backData.m_iColorIndex = 3;
	InitBackEdgeData(backData, entBackId, pBlockTb);
	g_mapStraightBackData[backData] = entBackId;
	g_mapIDToBlockText[entBackId] = backData.m_strBackEdge;

	backData.m_strBackEdge = _T("500B");
	backData.m_dLen = 500;
	backData.m_bBlue = FALSE;
	backData.m_dWith = 80.0;
	backData.m_iColorIndex = 3;
	InitBackEdgeData(backData, entBackId, pBlockTb);
	g_mapStraightBackData[backData] = entBackId;
	g_mapIDToBlockText[entBackId] = backData.m_strBackEdge;

	backData.m_strBackEdge = _T("500X350");
	backData.m_dLen = 500;
	backData.m_dLen2 = 350;
	backData.m_bBlue = TRUE;
	backData.m_dWith = 90.0;
	backData.m_iColorIndex = 4;
	InitBackEdgeData(backData, entBackId, pBlockTb);
	g_mapTurnBackData[backData] = entBackId;
	g_mapIDToBlockText[entBackId] = backData.m_strBackEdge;

	backData.m_strBackEdge = _T("1000X500");
	backData.m_dLen = 1000;
	backData.m_dLen2 = 500;
	InitBackEdgeData(backData, entBackId, pBlockTb);
	g_mapTurnBackData[backData] = entBackId;
	g_mapIDToBlockText[entBackId] = backData.m_strBackEdge;

	backData.m_strBackEdge = _T("1000X1000");
	backData.m_dLen = 1000;
	backData.m_dLen2 = 1000;
	InitBackEdgeData(backData, entBackId, pBlockTb);
	g_mapTurnBackData[backData] = entBackId;
	g_mapIDToBlockText[entBackId] = backData.m_strBackEdge;

	pBlockTb->close();
	return !g_mapTurnBackData.empty()&&!g_mapStraightBackData.empty();
}

void CRLDrawBackEdge::InitBackEdgeData(const BackEdgeTypeData& data, AcDbObjectId& entId, AcDbBlockTable*& pBlockTb)
{
	double dWith = data.m_dWith;
	if (pBlockTb->has(data.m_strBackEdge))
	{
		pBlockTb->getAt(data.m_strBackEdge, entId);
		if (entId.isValid())
			return;
		AcDbBlockTableRecord* pBlockTbRec = NULL;
		pBlockTb->getAt(data.m_strBackEdge, pBlockTbRec, AcDb::kForWrite);
		if (pBlockTbRec)
		{
			pBlockTbRec->erase(true);
			pBlockTbRec->close();
		}
	}
	AcDbBlockTableRecord* pBlockTbRec = new AcDbBlockTableRecord;
	pBlockTbRec->setName(data.m_strBackEdge);
	pBlockTb->add(entId, pBlockTbRec);
	if (data.m_dLen2 < EPRES)
	{
		double dLen = data.m_dLen;
		AcDbPolyline* pPline = new AcDbPolyline();
		pPline->setClosed(true);
		pPline->addVertexAt(0, AcGePoint2d(0, 0));
		pPline->addVertexAt(1, AcGePoint2d(0, dWith));
		pPline->addVertexAt(2, AcGePoint2d(dLen, dWith));
		pPline->addVertexAt(3, AcGePoint2d(dLen, 0));
		pPline->setColorIndex(data.m_iColorIndex);

// 		AcDbText* pText = new AcDbText();
// 		pText->setTextStyle(getTextStyleId(_T("Standard")));
// 		pText->setTextString(data.m_strBackEdge);
// 		pText->setColorIndex(data.m_iColorIndex);
// 		pText->setHeight(50.0);
// 		pText->setHorizontalMode(AcDb::kTextLeft);
// 		pText->setVerticalMode(AcDb::kTextBottom);
// 		pText->setAlignmentPoint(AcGePoint3d(dLen / 2 - 50, 15, 0));
// 		pText->setPosition(AcGePoint3d(dLen / 2 - 50, 15, 0));

		AcDbObjectId entTmpId = AcDbObjectId::kNull;
//		pBlockTbRec->appendAcDbEntity(entTmpId, pText);
		pBlockTbRec->appendAcDbEntity(entTmpId, pPline);

		pPline->close();
//		pText->close();
		pPline = NULL;
//		pText = NULL;

	}
	else
	{
		double dLen1 = data.m_dLen2;
		double dLen2 = data.m_dLen;
		AcGePoint3d ptInsert(dWith - 15, -15, 0);
		if (dLen2 > 999)
			ptInsert = AcGePoint3d(dLen2 / 2.0 - 2 * g_minLenIn, -15, 0);
		
		AcDbPolyline* pPline = new AcDbPolyline();
		pPline->setClosed(true);
		pPline->addVertexAt(0, AcGePoint2d(0, -dLen1));
		pPline->addVertexAt(1, AcGePoint2d(0, 0));
		pPline->addVertexAt(2, AcGePoint2d(dLen2, 0));
		pPline->addVertexAt(3, AcGePoint2d(dLen2, -dWith));
		pPline->addVertexAt(4, AcGePoint2d(dWith, -dWith));
		pPline->addVertexAt(5, AcGePoint2d(dWith, -dLen1));
		pPline->setColorIndex(data.m_iColorIndex);

		AcDbText* pText = new AcDbText();
		pText->setTextStyle(getTextStyleId(_T("Standard")));
		pText->setTextString(data.m_strBackEdge);
		pText->setColorIndex(data.m_iColorIndex);
		pText->setHeight(50.0);
		pText->setHorizontalMode(AcDb::kTextLeft);
		pText->setVerticalMode(AcDb::kTextTop);
		pText->setAlignmentPoint(ptInsert);
		pText->setPosition(ptInsert);

		AcDbObjectId entTmpId = AcDbObjectId::kNull;
		pBlockTbRec->appendAcDbEntity(entTmpId, pText);
		pBlockTbRec->appendAcDbEntity(entTmpId, pPline);

		pPline->close();
		pText->close();
		pPline = NULL;
		pText = NULL;
	}
	pBlockTbRec->close();
	pBlockTbRec = NULL;
}

BOOL CRLDrawBackEdge::getBackEdgeExcelData(CTableDataArry& tableArry)
{
	std::map<AcDbObjectId, CString> mapIdToBackType;
	std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapStraightBackData.begin();
	for (iter; iter != g_mapStraightBackData.end(); iter++)
	{
		mapIdToBackType[iter->second] = iter->first.m_strBackEdge;
	}
	iter = g_mapTurnBackData.begin();
	for (iter; iter != g_mapTurnBackData.end(); iter++)
	{
		mapIdToBackType[iter->second] = iter->first.m_strBackEdge;
	}
	if (mapIdToBackType.empty())
	{
		acutPrintf(_T("\n获取背楞类型失败！"));
		return FALSE;
	}
	AcArray<CString> strBackTypeArry;
	strBackTypeArry.append(g_strBackEdge1);
	strBackTypeArry.append(g_strBackEdge2);
	strBackTypeArry.append(g_strBackEdge3);
	strBackTypeArry.append(g_strBackEdge4);
	strBackTypeArry.append(g_strBackEdge5);
	strBackTypeArry.append(g_strBackEdge6);
	strBackTypeArry.append(g_strBackEdge7);
	strBackTypeArry.append(g_strBackEdge8);
	strBackTypeArry.append(g_strBackEdge9);
	std::map<CString, int> mapBackEdgeNos;
	AcDbObjectIdArray entIdArry;
	if (!selectBackEdges(entIdArry))
		return FALSE;
	int nLen = entIdArry.length();
	for (int i = 0; i < nLen; i++)
	{
		AcDbObjectPointer<AcDbEntity> pEnt(entIdArry[i], AcDb::kForRead);
		if (pEnt.openStatus() != Acad::eOk)
			continue;
		getBackEdgeExcelData(mapBackEdgeNos, mapIdToBackType, strBackTypeArry,pEnt);
	}
	if (mapBackEdgeNos.empty())
		return FALSE;
	std::map<CString, int>::iterator iterNos = mapBackEdgeNos.begin();
	int nNos = 1;
	for (iterNos; iterNos != mapBackEdgeNos.end(); iterNos++, nNos++)
	{
		vector<CString> vecStr;
		CString strTmp;
		strTmp.Format(_T("%d"), nNos);
		vecStr.push_back(strTmp);
		vecStr.push_back(iterNos->first);
		strTmp.Format(_T("%d"), iterNos->second);
		vecStr.push_back(strTmp);
		tableArry.push_back(vecStr);
	}
	return !tableArry.empty();
}

BOOL CRLDrawBackEdge::getBackEdgeExcelData(std::map<CString, int>&mapBackEdgeNos, const std::map<AcDbObjectId, CString>& mapIdToBackType,
	AcArray<CString>& strBackTypeArry, AcDbEntity* pEnt)
{
	if (pEnt == NULL)
		return FALSE;
	if (!pEnt->isKindOf(AcDbBlockReference::desc()) && !pEnt->isKindOf(AcDbText::desc())
		&&!pEnt->isKindOf(AcDbMText::desc()))
	{
		return FALSE;
	}
	CString strText;
	if (pEnt->isKindOf(AcDbBlockReference::desc()))
	{
		AcDbObjectId entBlockId = AcDbBlockReference::cast(pEnt)->blockTableRecord();
		std::map<AcDbObjectId, CString>::const_iterator iterBlock = mapIdToBackType.find(entBlockId);
		if (iterBlock != mapIdToBackType.end())
			strText = iterBlock->second;
		
		else
		{
			AcGeVoidPointerArray ptIines;
			if (pEnt->explode(ptIines) != Acad::eOk)
				releaseVoidPtr(ptIines);
			int nSize = ptIines.length();
			for (int j = 0; j < nSize; j++)
			{
				getBackEdgeExcelData(mapBackEdgeNos, mapIdToBackType, strBackTypeArry, static_cast<AcDbEntity*>(ptIines[j]));
				AcDbObject *pObj = static_cast<AcDbObject*>(ptIines[j]);
				if (pObj)
				{
					delete pObj;
					pObj = NULL;
				}
			}
		}
	}
	else if (pEnt->isKindOf(AcDbText::desc()))
	{
		CString strTextTmp = AcDbText::cast(pEnt)->textString();
		if (strBackTypeArry.contains(strTextTmp))
			strText = strTextTmp;
	}
	else if (pEnt->isKindOf(AcDbMText::desc()))
	{
		CString strTextTmp = AcDbMText::cast(pEnt)->text();
		if (strBackTypeArry.contains(strTextTmp))
			strText = strTextTmp;
	}
	if (!strText.IsEmpty())
	{
		int nNumber = 0;
		std::map<CString, int>::iterator iterNos = mapBackEdgeNos.find(strText);
		if (iterNos != mapBackEdgeNos.end())
			nNumber = iterNos->second;
		nNumber++;
		mapBackEdgeNos[strText] = nNumber;
	}
	return TRUE;
}

BOOL CRLDrawBackEdge::getExcelTempPath(CString& strPath)
{
	//CString StrFilter = _T("Excel 模板(*.xlt、*.xltx、*.xls)|*.xlt;*.xltx;*.xls||");
	// 	CFileDialog  dlg(TRUE, _T("*.xltx"), NULL, OFN_FILEMUSTEXIST, StrFilter, NULL);
	// 	if (IDOK == dlg.DoModal())
	// 		strPath = dlg.GetPathName();
	CString strModulePath = _T("");
	CString strArxFullPath;
	TCHAR lpszFilePath[_MAX_PATH];
	if (::GetModuleFileName(_hdllInstance, lpszFilePath, _MAX_PATH) == 0)
		return FALSE;
	strArxFullPath = lpszFilePath;
	int nLength = strArxFullPath.ReverseFind('\\');
	strModulePath = strArxFullPath.Left(nLength);
	strPath = strModulePath + _T("\\清单模板.xls");
	return _taccess((TCHAR *)LPCTSTR(strPath), 4) != -1;
}

BOOL CRLDrawBackEdge::getExcelSavePath(CString& strPath)
{
	CString StrFilter = _T("Excel 工作簿|*.xlsx;*.xls||");
	CFileDialog  dlg2(FALSE, NULL, NULL, OFN_HIDEREADONLY, StrFilter, NULL);
	dlg2.m_ofn.lpstrTitle = _T("保存EXCEL文件");
	if (IDOK == dlg2.DoModal())
		strPath = dlg2.GetPathName();
	if (strPath.IsEmpty())
		return FALSE;
	std::ofstream ofile(strPath);
	if (!ofile)
		return FALSE;
	ofile.close();
	::DeleteFile(strPath);
	return TRUE;
}

//生成EXCEL清单
void CRLDrawBackEdge::CreatExcel()
{
	CTableDataArry tableArry;
	if (!getBackEdgeExcelData(tableArry))
	{
		acutPrintf(_T("\n没有选择背楞！"));
		return;
	}
	
	CString strExcelTemplatePath;
	CString strExcelPath;
	while (!getExcelTempPath(strExcelTemplatePath))
	{
		if (strExcelTemplatePath.IsEmpty())
			return;
		AfxMessageBox(_T("选择的模板文件不可读，请重新选择!\n"));
		continue;
	}

	while (!getExcelSavePath(strExcelPath))
	{
		if (strExcelPath.IsEmpty())
			return;
		AfxMessageBox(_T("存储路径不合法或该文件正在使用，无法覆盖!\n"));
		continue;
	}

	acedSetStatusBarProgressMeter(_T("正在生成Excel表格..."), 0, tableArry.size());
	
	try
	{
		CRLOperExcel operExcel;
		operExcel.CreateXls(FALSE);
		boost::optional<CWorkbook> templateFile = operExcel.OpenFile(strExcelTemplatePath);
		if (!templateFile)
		{
			acutPrintf(_T("\n打开Excel表格模板文件失败!"));
			throw CString();
		}
		// 生成表格文件
		boost::optional<CWorkbook> tableFile = operExcel.NewFile();
		if (!tableFile)
			throw CString();
		int nCurProgress(0);
		int nCurPage = 1;   // 当前页

		int nSrcPage = nCurPage;
		int nDestPage = nCurPage;

		int nTempSheetCount = operExcel.GetSheetCount(templateFile.get());
		if (nSrcPage > nTempSheetCount)
			nSrcPage = nTempSheetCount;

		// 如果需要则在目标表格中添加工作表
		int nTableSheetCount = operExcel.GetSheetCount(tableFile.get());
		if (nDestPage > nTableSheetCount)
		{
			int nIndex = operExcel.AddSheetBack(tableFile.get());
			assert(nDestPage == nIndex);
		}

		CString sSheetName;
		sSheetName.Format(_T("第%d页"), nDestPage);
		operExcel.SetSheetName(tableFile.get(), nDestPage, sSheetName);

		// 从当前页表格模板中复制表格
		if (!operExcel.CopySheetContents(templateFile.get(), nSrcPage, tableFile.get(), nDestPage))
			throw CString();

		vector<ExcelTemplateInfo> vecExcelTableInfos;
		if (!GetExcelTemplateInfo(templateFile.get(), nSrcPage, vecExcelTableInfos))
			throw CString();

		if (vecExcelTableInfos.size() != 1)
			throw CString(_T("Excel表格模板不匹配!"));

		int i;
		for (i = 0; i < tableArry.size(); ++i)
		{
			// 进度条
			nCurProgress += 1;
			acedSetStatusBarProgressMeterPos(nCurProgress);

			int nRow, nCol(0);
			for (int j = 0; j < tableArry[i].size(); ++j)
			{
				nRow = vecExcelTableInfos[0].nRow + i;
				if (nCol == 0)
					nCol = vecExcelTableInfos[0].nColumn + j;

				operExcel.SetCellText(tableFile.get(), nDestPage,
					nRow, nCol, tableArry[i][j]);

				// 将单元格格式设置为缩小字体填充
				operExcel.SetShrinkToFit(tableFile.get(), nDestPage, nRow, nCol);

				// 如果当前行列所在单元格是合并单元格，列数加上合并单元格的个数
				nCol += operExcel.GetMergeCellsCount(tableFile.get(), nDestPage, nRow, nCol);
			}
		}

		// 关闭模板文件
		operExcel.Close(templateFile.get());

		// 保存并关闭表格文件
		operExcel.ActiveSheet(tableFile.get(), 1);
		operExcel.SaveAs(tableFile.get(), strExcelPath);
		operExcel.Close(tableFile.get());
	}

	catch (CString& e)
	{
		if (!e.IsEmpty())
			acutPrintf(_T("\n生成Excel表格失败: %s"), e);
		else
			acutPrintf(_T("\n生成Excel表格失败!"));
		acedRestoreStatusBar();
		return ;
	}
	catch (...)
	{
		acedRestoreStatusBar();
		acutPrintf(_T("\n生成Excel表格失败!"));
		return;
	}

	acedSetStatusBarProgressMeterPos(0);
	acedRestoreStatusBar();
	acutPrintf(_T("\n生成Excel表格成功!"));
}

bool CRLDrawBackEdge::GetExcelTemplateInfo(CWorkbook& workBook, int nSheetIndex, vector<ExcelTemplateInfo>& vecExcelTableInfos)
{
	assert(vecExcelTableInfos.empty());

	CWorksheets sheets(workBook.get_Worksheets());
	if (sheets.m_lpDispatch == NULL)
		return false;

	CWorksheet sheet(sheets.get_Item(_variant_t(nSheetIndex)));
	if (sheet.m_lpDispatch == NULL)
		return false;

	CRange allCells(sheet.get_Cells());
	if (allCells.m_lpDispatch == NULL)
		return false;

	CString sFirstMatch;
	CRange  lastPos;
	while (true)
	{
		VARIANT after = lastPos.m_lpDispatch ? _variant_t(lastPos) : vtMissing;
		CRange targetCell(allCells.Find(_variant_t(_T("startRow")), after, vtMissing,
			vtMissing, vtMissing, xlNext, vtMissing, vtMissing, vtMissing));
		if (targetCell.m_lpDispatch == NULL)
			break;

		lastPos = targetCell;

		if (sFirstMatch.IsEmpty())
			sFirstMatch = targetCell.get_Address(vtMissing, vtMissing, xlA1, vtMissing, vtMissing);
		else if (sFirstMatch == targetCell.get_Address(vtMissing, vtMissing, xlA1, vtMissing, vtMissing))
			break;

		CString sText = (LPCTSTR)(_bstr_t)_variant_t(targetCell.get_Text());
		sText.MakeLower();
		int nPos1 = sText.Find(_T("rownumber"));
		if (nPos1 == -1)
			return false;

		nPos1 += 9;
		if (nPos1 >= sText.GetLength())
			return false;

		int nPos2 = sText.Find(';', nPos1);
		if (nPos2 == -1)
			nPos2 = sText.GetLength();

		CString sRowCount = sText.Mid(nPos1, nPos2 - nPos1);
		sRowCount.TrimLeft(_T(" ="));  // 去掉开头的空格和等号

		ExcelTemplateInfo excelTableInfo;
		excelTableInfo.nRowCount = _tstoi(sRowCount);
		excelTableInfo.nColumn = (int)targetCell.get_Column();
		excelTableInfo.nRow = (int)targetCell.get_Row();
		vecExcelTableInfos.push_back(excelTableInfo);
	}

	if (vecExcelTableInfos.empty())
		return false;

	return true;
}

BOOL CRLDrawBackEdge::getPolyPathLine(std::set<AcGePolyline2d *>& polygons, AcDbObjectIdArray& entDwgIdArry)
{
	if (g_mapStraightBackData.empty() || g_mapTurnBackData.empty())
	{
		if (!InitBackEdge())
		{
			acutPrintf(_T("\n背楞类型初始化失败！"));
			return FALSE;
		}
	}
	createLayer(_T("fmk_lines"), 1);
	AcDbObjectIdArray entResLineIds;
	if (!selectAutoLines(entResLineIds, entDwgIdArry))
	{
		earseEnts(entResLineIds);
		return FALSE;
	}
	frgExtractTopologicalEntsFromLinesAlgm frgLinesAlgm(entResLineIds);
	rlTopologicalEnt2dSet * pEntSet = frgLinesAlgm.start();
	if (pEntSet == NULL)
	{
		earseEnts(entResLineIds);
		return FALSE;
	}
	frgPolygon2dRecognitionAlgm frgRecognitionAlgm(pEntSet);
	polygons = frgRecognitionAlgm.start();
	if (pEntSet)
	{
		delete pEntSet;
		pEntSet = NULL;
	}
	if (polygons.empty())
	{
		earseEnts(entResLineIds);
		return FALSE;
	}
	ShowEnt(entDwgIdArry, FALSE);
	earseEnts(entResLineIds);

	polygons = mergePolyline(polygons);

	return !polygons.empty();
}
void CRLDrawBackEdge::updataWallPts(AcArray<AcGePoint3dArray>& pathPtArrys)
{
	AcArray<AcGePoint3dArray> pathResArrys = pathPtArrys;
	pathPtArrys.removeAll();
	int nLen = pathResArrys.length();
	for (int i = 0; i < nLen; i++)
	{
		AcGePoint3dArray ptResArry = pathResArrys[i];
		AcGePoint3dArray ptArry;
		int size = ptResArry.length();
		if (size < 2)
			continue;
		BOOL bIsLoop = ptResArry.first().isEqualTo(ptResArry.last(), frgGlobals::Gtol);
		for (int i = 0; i < size;i++)
		{
			int iStartIndex = i - 1;
			int iEndIndex = i + 1;
			if (bIsLoop)
			{
				if (i == 0)
					iStartIndex = size - 2;
				if (i == size - 2)
					iEndIndex = 0;
				if (iStartIndex == iEndIndex || i == size -1)
					continue;
			}
			else
			{
				if (i == 0)
					iStartIndex = i;
				if (i == size - 1)
					iEndIndex = i;
			}
			AcGePoint3d pt = ptResArry[i];
			AcGePoint3d ptStart = ptResArry[iStartIndex];
			AcGePoint3d ptEnd = ptResArry[iEndIndex];
			AcGeVector3d vecX = pt - ptStart;
			AcGeVector3d vecY = ptEnd - pt;
			vecX = vecX.perpVector().normal();
			vecY = vecY.perpVector().normal();
			if (vecX.isZeroLength())
				vecX = (pt - ptEnd).normal();
			if (vecY.isZeroLength())
				vecY = (pt - ptStart).normal();
			pt = pt + vecX*g_dWallOutLen + vecY*g_dWallOutLen;
			ptArry.append(pt);
		}
		if (bIsLoop)
			ptArry.append(ptArry.first());
		if (!ptArry.isEmpty())
			pathPtArrys.append(ptArry);
	}
}

void CRLDrawBackEdge::updataWallPts(std::set<AcGePolyline2d *>& polygons)
{
	std::set<AcGePolyline2d *> pResPoline = polygons;
	polygons.clear();
	std::set<AcGePolyline2d *>::const_iterator it = pResPoline.begin();
	for (int m = 0; it != pResPoline.end(); ++it, m++)
	{
		AcGePolyline2d *polygon = *it;
		if (polygon == NULL)
			continue;
		int size = polygon->numFitPoints();
		if (size < 2)
			continue;
		AcGePoint2dArray vertices;
		for (int i = 0; i < size - 1; i++)
	    {
			int iStartIndex = i -1;
			int iEndIndex = i+1;
			if (i == 0)
				iStartIndex = size - 2;
			if (i == size - 2)
				iEndIndex = 0;
			if (iStartIndex == iEndIndex)
			    continue;
			AcGePoint2d pt = polygon->fitPointAt(i);
			AcGePoint2d ptStart = polygon->fitPointAt(iStartIndex);
			AcGePoint2d ptEnd = polygon->fitPointAt(iEndIndex);
			AcGeVector2d vecX = ptStart - pt;
			AcGeVector2d vecY = pt - ptEnd;
			vecX = vecX.perpVector().normal();
			vecY = vecY.perpVector().normal();
// 			double dLen = vecY.length();
// 			AcGeVector2d vecXPre = vecX;
// 			vecXPre.rotateBy(PI / 2);
// 			if (!(pt + vecXPre.normal()*dLen).isEqualTo(ptEnd,frgGlobals::Gtol))
// 			{
// 				vecX = -vecX;
// 				vecY = -vecY;
// 			}
// 			vecX.normalize();
// 			vecY.normalize();
			pt = pt + vecX*g_dWallOutLen + vecY*g_dWallOutLen;
			vertices.append(pt);
	    }
		vertices.append(vertices.first());
		delete polygon;
		if (!vertices.isEmpty())
		{
			polygon = new AcGePolyline2d(vertices);
			polygons.insert(polygon);
		}
	}
}

void CRLDrawBackEdge::AutodrawBackEdgeWall()
{
	AcDbObjectIdArray entDwgIds;
	std::set<AcGePolyline2d *> polygons;
	if (!getPolyPathLine(polygons, entDwgIds))
		return;

	updataWallPts(polygons);

	AcDbObjectIdArray entResLineIds = drawPathLines(polygons);
	assert(!entResLineIds.isEmpty());

	AcArray<AcGePoint3dArray> pathDelPtArrys = delBackLines(entResLineIds);
	earseEnts(entResLineIds);

	AcArray<AcGePoint3dArray> pathPtArrys;
	std::set<AcGePolyline2d *>::const_iterator it = polygons.begin();
	acedSetStatusBarProgressMeter(_T("正在计算背楞布置路径..."), 0, (int)polygons.size());
	for (int m = 0; it != polygons.end(); ++it, m++)
	{
		AcGePolyline2d *polygon = *it;
		if (polygon == NULL)
			continue;
		getBackPathLines(pathPtArrys, pathDelPtArrys, polygon);
		acedSetStatusBarProgressMeterPos(m);
		delete polygon;
		polygon = NULL;
	}
	acedRestoreStatusBar();
	polygons.clear();


	dealWithErrorPt(pathPtArrys);
	drawBackEdge(pathPtArrys);
	ShowEnt(entDwgIds, TRUE);
}

void CRLDrawBackEdge::AutodrawBackEdge()
{
	AcDbObjectIdArray entDwgIds;
	std::set<AcGePolyline2d *> polygons;
	if (!getPolyPathLine(polygons, entDwgIds))
		return;

	AcDbObjectIdArray entResLineIds = drawPathLines(polygons);
	assert(!entResLineIds.isEmpty());

	AcArray<AcGePoint3dArray> pathDelPtArrys = delBackLines(entResLineIds);
	earseEnts(entResLineIds);

	AcArray<AcGePoint3dArray> pathPtArrys;
	std::set<AcGePolyline2d *>::const_iterator it = polygons.begin();
	acedSetStatusBarProgressMeter(_T("正在计算背楞布置路径..."), 0, (int)polygons.size());
	for (int m = 0; it != polygons.end(); ++it, m++)
	{
		AcGePolyline2d *polygon = *it;
		if (polygon == NULL)
			continue;
		getBackPathLines(pathPtArrys, pathDelPtArrys, polygon);
		acedSetStatusBarProgressMeterPos(m);
		delete polygon;
		polygon = NULL;
	}
	acedRestoreStatusBar();
	polygons.clear();
	
	
	dealWithErrorPt(pathPtArrys);
	drawBackEdge(pathPtArrys);
	ShowEnt(entDwgIds, TRUE);
}

std::set<AcGePolyline2d *> CRLDrawBackEdge::mergePolyline(std::set<AcGePolyline2d *>& polygons)
{
	std::set<AcGePolyline2d *> pResPoline;
	std::set<AcGePolyline2d *>::const_iterator it = polygons.begin();
	for (int m = 0; it != polygons.end(); ++it, m++)
	{
		AcGePolyline2d *polygon = *it;
		if (polygon == NULL)
			continue;
		int size = polygon->numFitPoints();
		if (size <= 2)
		    continue;
		AcGePoint2dArray vertices;
		AcGePoint2d pt1 = polygon->fitPointAt(size-1);
		AcGePoint2d pt2 = polygon->fitPointAt(size-2);
		AcGeVector2d vecNormal = (pt2 - pt1).normal();
		int iNdex0 = 0;
		while (iNdex0 < size)
		{
			pt2 = polygon->fitPointAt(iNdex0);
			if (pt2.isEqualTo(pt1,frgGlobals::Gtol))
			{
				iNdex0++;
				continue;
			}
			AcGeVector2d vecNormalTmp = (pt2 - pt1).normal();
			if (vecNormalTmp.isParallelTo(vecNormal, frgGlobals::Gtol))
			{
				iNdex0++;
				continue;
			}
			iNdex0--;
			break;
		}
		if (iNdex0 < 0)
			iNdex0 = size + iNdex0;
		pt1 = polygon->fitPointAt(iNdex0);
		vertices.append(pt1);
		for (int i = size - 2; i >= iNdex0; i--)
		{
			pt2 = polygon->fitPointAt(i);
			if (i == iNdex0)
			{
				vertices.append(pt2);
				continue;
			}
			pt1 = polygon->fitPointAt(i-1);
			AcGeVector2d vecNormalTmp = (pt2 - pt1).normal();
			if (vecNormalTmp.isParallelTo(vecNormal,frgGlobals::Gtol))
				continue;
			vecNormal = vecNormalTmp;
			vertices.append(pt2);
		}
		delete polygon;
		if (!vertices.isEmpty())
		{
			polygon = new AcGePolyline2d(vertices);
			pResPoline.insert(polygon);
		}
	}
	return pResPoline;
}

AcDbObjectIdArray CRLDrawBackEdge::drawPathLines(const std::set<AcGePolyline2d *>& polygons)
{
	AcDbObjectIdArray entIdArry;
	std::set<AcGePolyline2d *>::const_iterator it = polygons.begin();
	for (int m = 0; it != polygons.end(); ++it, m++)
	{
		AcGePolyline2d *polygon = *it;
		if (polygon == NULL)
			continue;
		int size = polygon->numFitPoints();
		for (int i = 0; i <size-1; i++)
		{
			AcGePoint3d pt1 = getPoint3D(polygon->fitPointAt(i));
			AcGePoint3d pt2 = getPoint3D(polygon->fitPointAt(i+1));
			AcDbLine* pLine = new AcDbLine(pt1, pt2);
			pLine->setLayer(_T("fmk_lines"));
			pLine->setColorIndex(256);
			AcDbObjectId ret;
			if (!AddToModelSpace(pLine,ret))
			{
				delete pLine;
				pLine = NULL;
			}
			entIdArry.append(ret);
		}
	}
	return entIdArry;
}

void CRLDrawBackEdge::getBackPathLines(AcArray<AcGePoint3dArray>& pathPtArrys, const AcArray<AcGePoint3dArray>&pathDelPtArrys, AcGePolyline2d *polygon)
{
	AcArray<int> delIndex;
	int size = polygon->numFitPoints()-1;
	if (pathDelPtArrys.isEmpty())
	{
		if (size < 2)
			return;
		AcGePoint3dArray ptArry;
		for (int n = size-1; n >=0 ; n--)
		{
			AcGePoint3d pt = getPoint3D(polygon->fitPointAt(n));
			ptArry.append(pt);
		}
		ptArry.append(getPoint3D(polygon->fitPointAt(size - 1)));
		if (ptArry.length() >= 2 && !pathPtArrys.contains(ptArry))
			pathPtArrys.append(ptArry);
		return;
	}
	for (int n = 0; n < size; n++)
	{
		AcGePoint3d pt1, pt2;
		int iNdex = n;
		if (n == size - 1)
			iNdex = 0;
		else
			iNdex = n + 1;
		pt1 = getPoint3D(polygon->fitPointAt(iNdex));
		pt2 = getPoint3D(polygon->fitPointAt(n));
		AcGePoint3dArray ptArry;
		ptArry.append(pt1);
		ptArry.append(pt2);
		if (pathDelPtArrys.contains(ptArry))
			delIndex.append(iNdex);
	}
	if (delIndex.isEmpty())
	{
		if (size < 2)
			return;
		AcGePoint3dArray ptArry;
		for (int n = size - 1; n >= 0; n--)
		{
			AcGePoint3d pt = getPoint3D(polygon->fitPointAt(n));
			ptArry.append(pt);
		}
		ptArry.append(getPoint3D(polygon->fitPointAt(size - 1)));
		if (ptArry.length() >= 2 && !pathPtArrys.contains(ptArry))
			pathPtArrys.append(ptArry);
		return;
	}
	int nLen = delIndex.length();
	int nStartIndex = 0;
	int nEndIndex = delIndex[0];
	for (int i = 1; i < nLen; i++)
	{
		nStartIndex = delIndex[i];
		if (nStartIndex == 0)
			nStartIndex = size - 1;
		else
			nStartIndex = nStartIndex - 1;
		int nTotals = nStartIndex - nEndIndex;
		if (nTotals < 0)
			nTotals = size + nTotals;
		AcGePoint3dArray ptArry;
		for (int j = 0; j <= nTotals; j++)
		{
			if (nStartIndex < 0)
				nStartIndex = size + nStartIndex;
			AcGePoint3d pt = getPoint3D(polygon->fitPointAt(nStartIndex));
			ptArry.append(pt);
			nStartIndex--;
		}
		if (ptArry.length() >= 2&&!pathPtArrys.contains(ptArry))
			pathPtArrys.append(ptArry);
		nEndIndex = delIndex[i];
	}
	nStartIndex = delIndex[0];
	if (nStartIndex == 0)
		nStartIndex = size - 1;
	else
		nStartIndex = nStartIndex - 1;
	int nTotals = nStartIndex - nEndIndex;
	if (nTotals < 0)
		nTotals = size + nTotals;
	AcGePoint3dArray ptArry;
	for (int i = 0; i <= nTotals; i++)
	{
		if (nStartIndex < 0)
			nStartIndex = size + nStartIndex;
		AcGePoint3d pt = getPoint3D(polygon->fitPointAt(nStartIndex));
		ptArry.append(pt);
		nStartIndex--;
	}
	if (ptArry.length() >= 2&&!pathPtArrys.contains(ptArry))
		pathPtArrys.append(ptArry);
}

AcDbObjectIdArray CRLDrawBackEdge::exPlodePline(AcDbObjectIdArray& entDwgIds)
{
	struct resbuf rb;
	rb.restype = 0; //实体名
	CString strEnt = _T("LWPOLYLINE");
	rb.resval.rstring = strEnt.GetBuffer();
	strEnt.ReleaseBuffer();
	rb.rbnext = NULL; //无其他内容
	ads_name ssText;
	acutPrintf(_T("\n请选择背楞多段线路径:"));
	int iRet = acedSSGet(_T("X"), NULL, NULL, &rb, ssText);
	if (RTNORM != iRet)
	{
		acedSSFree(ssText);
		return entDwgIds;
	}
	long ssLen;
	acedSSLength(ssText, &ssLen);
	for (int i = 0; i < ssLen; i++)
	{
		ads_name ent;
		acedSSName(ssText, i, ent);
		AcDbObjectId eld;
		acdbGetObjectId(eld, ent);
		if (eld.isValid() && !entDwgIds.contains(eld))
			entDwgIds.append(eld);
		acedSSFree(ent);
	}
	acedSSFree(ssText);
	AcDbObjectIdArray entLineIdArry;
	int nLen = entDwgIds.length();
	for (int i = 0; i < nLen; i++)
	{
		AcDbObjectId entId = entDwgIds[i];
		AcDbObjectPointer<AcDbPolyline> pEnt(entId, AcDb::kForRead);
		if (pEnt.openStatus() != Acad::eOk)
			continue;
		int nNumber = pEnt->numVerts();
		if (nNumber < 2)
			continue;
		AcGeVoidPointerArray ptIines;
		if (pEnt->explode(ptIines) != Acad::eOk)
			releaseVoidPtr(ptIines);
		int nSize = ptIines.length();
		for (int j = 0; j < nSize; j++)
		{
			AcDbLine *pLineEnt = static_cast<AcDbLine*>(ptIines[j]);
			if (pLineEnt)
			{
				pLineEnt->setPropertiesFrom(pEnt);
				AcDbObjectId ret;
				if (!AddToModelSpace(pLineEnt, ret))
				{
					delete pEnt;
					pLineEnt = NULL;
				}
				entLineIdArry.append(ret);
			}
			else
			{
				AcDbObject *pObj = static_cast<AcDbObject*>(ptIines[j]);
				if (pObj)
				{
					delete pObj;
					pObj = NULL;
				}
			}
		}
	}
	actrTransactionManager->flushGraphics();
	acedUpdateDisplay();
	return entLineIdArry;
}

void CRLDrawBackEdge::releaseVoidPtr(AcGeVoidPointerArray& ptIines)
{
	int nLen = ptIines.length();
	for (int i = 0; i < nLen; i++)
	{
		AcDbObject *pObj = static_cast<AcDbObject*>(ptIines[i]);
		if (pObj)
		{
			delete pObj;
			pObj = NULL;
		}
	}
	ptIines.removeAll();
}

void CRLDrawBackEdge::drawBackEdge()
{
	if (g_mapStraightBackData.empty() || g_mapTurnBackData.empty())
	{
		if (!InitBackEdge())
		{
			acutPrintf(_T("\n背楞类型初始化失败！"));
			return;
		}
	}
	
	createLayer(g_strBackLayer, 7);
	AcDbObjectIdArray entIdArry;
	if (!selectPLines(entIdArry))
		return;
	int nLen = entIdArry.length();
	AcArray<AcGePoint3dArray> pathPtArrys;
	for (int i = 0; i < nLen; i++)
	{
		AcDbObjectId entId = entIdArry[i];
		AcDbObjectPointer<AcDbPolyline> pEnt(entId, AcDb::kForRead);
		if (pEnt.openStatus() != Acad::eOk)
		   continue;
		int nNumber = pEnt->numVerts();
		if (nNumber < 2)
		   continue;
		AcGePoint3dArray ptArry;
		AcGePoint3d ptTmp;
		for (int j = 0; j < nNumber; j++)
		{
			pEnt->getPointAt(j, ptTmp);
			if (!ptArry.contains(ptTmp))
				ptArry.append(ptTmp);
		}
		if (pEnt->isClosed())
		{
			ptArry.append(ptArry.first());
			if (ptArry.length() < 5)
				continue;
		}
		if (ptArry.length() >= 2 && !pathPtArrys.contains(ptArry))
			pathPtArrys.append(ptArry);
	}
	dealWithErrorPt(pathPtArrys);
	
	CString strWall = g_bIsWall ? _T("Y"):_T("N");
	int nRet2 = acedInitGet(0, _T("Y N"));
	if (nRet2 != RTNORM)
		return;
	strWall.Format(_T("\n是否墙背楞[是(Y)/否(N)]<%s>: "), strWall);
	ACHAR szInput2[32] = {};
	nRet2 = acedGetKword(strWall, szInput2);
	if (nRet2 == RTCAN)
		return;
	else if (nRet2 == RTNONE)
	{

	}
	else
	{
		if (_tcscmp(szInput2, _T("Y")) == 0)
			g_bIsWall = TRUE;
		else if (_tcscmp(szInput2, _T("N")) == 0)
			g_bIsWall = FALSE;
	}
	if (g_bIsWall)
		updataWallPts(pathPtArrys);

	drawBackEdge(pathPtArrys);
}

void CRLDrawBackEdge::dealWithErrorPt(AcArray<AcGePoint3dArray>& pathPtArrys)
{
	AcArray<AcGePoint3dArray> pathPtArryRse = pathPtArrys;
	pathPtArrys.removeAll();
	int nLen = pathPtArryRse.length();
	if (nLen == 0)
		return;
	for (int i = 0; i < nLen; i++)
	{
		AcGePoint3dArray ptArry = pathPtArryRse[i];
		int nSize = ptArry.length();
		if (nSize < 3)
		{
			pathPtArrys.append(ptArry);
			continue;
		}
		BOOL bTrueData = TRUE;
		for (int i = 2; i < nSize; i++)
		{
			AcGeVector3d vecNormlX = (ptArry[i - 2] - ptArry[i - 1]).normal();
			AcGeVector3d vecNormlY = (ptArry[i] - ptArry[i - 1]);
			AcGeVector3d vecResY = vecNormlX;
			double dAngle = vecNormlX.angleTo(vecNormlY,AcGeVector3d(0,0,1));
			if (fabs(dAngle - PI/2.0) < g_dAngleEqual*PI/180)
			{
				vecResY.rotateBy(PI / 2.0, AcGeVector3d(0, 0, 1));
				vecResY.normalize();
			}
			else if (fabs(dAngle - PI / 2.0*3.0) <g_dAngleEqual * PI / 180)
			{
				vecResY.rotateBy(PI / 2.0*3.0, AcGeVector3d(0, 0, 1));
				vecResY.normalize();
			}
			else 
			{
				bTrueData = FALSE;
				break;
			}
			AcGePoint3d pt1 = ptArry[i];
			double dLen = ptArry[i].distanceTo(ptArry[i - 1]);
			ptArry[i] = ptArry[i - 1] + vecResY*dLen;
			pt1 = ptArry[i];
		}
		if (bTrueData)
			pathPtArrys.append(ptArry);
	}
}

void CRLDrawBackEdge::drawBackEdge(const AcArray<AcGePoint3dArray>& pathPtArrys)
{
	int nLen = pathPtArrys.length();
	if (nLen == 0)
		return;
	
	for (int i = 0; i < nLen; i++)
	{
		AcGePoint3dArray ptArry = pathPtArrys[i];
		int nSize = ptArry.length();
		if (nSize < 2)
		   continue;

		BOOL bLoopBack = ptArry[0].isEqualTo(ptArry[nSize - 1], frgGlobals::Gtol);
		if (bLoopBack)
		{
			if (nSize < 5)
			    continue;
			ptArry.removeLast();
		}
		nSize = ptArry.length();
		
		AcGePoint3d ptInsert = ptArry[0];
		BOOL bNoNextData = nSize == 2;
		
		if (bNoNextData)
		{
			std::vector<straigthInsertData> vecBackInsetrtData;
			if (!getStraightBackDatas(vecBackInsetrtData,ptArry))
			   continue;
			nSize = vecBackInsetrtData.size();
			for (int j = 0; j < nSize; j++)
			{
				if (!addBlockResBlockReference(vecBackInsetrtData[j]))
					continue;
			}
		}
		else
		{
			std::vector<breakInsertData> vecBackInsetrtData;
			std::vector<straigthInsertData> vecStraigthData;
			if (!getBreakBackDatas(vecStraigthData, vecBackInsetrtData, ptArry, bLoopBack))
				continue;
			nSize = vecStraigthData.size();
			for (int j = 0; j < nSize; j++)
			{
				if (!addBlockResBlockReference(vecStraigthData[j]))
					continue;
			}
			nSize = vecBackInsetrtData.size();
			for (int j = 0; j < nSize; j++)
			{
				if (!addBlockResBlockReference(vecBackInsetrtData[j]))
					continue;
			}
		}
	}
}

BOOL CRLDrawBackEdge::getBreakBackDatas(std::vector<straigthInsertData>& vecStraigthData, std::vector<breakInsertData>& vecBreakData,
	                                    const AcGePoint3dArray& ptArry,BOOL bLoop)
{
	double dNextInlen = 0;
	int nSize = ptArry.length();
	if (nSize < 2)
		return FALSE;
	int iNdex = 0;
	if (bLoop)
	{
		if (!getLoopFirstBreakBackData(vecBreakData, ptArry))
			return FALSE;
	}
	else
	{
		if (!getNoLoopFirstBreakBackData(vecBreakData, ptArry))
			return FALSE;
		iNdex = 1;
	}
	std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter;
	iNdex = iNdex + 1;
	CMapIndexStargthData mapIndexStaiData;
	while (iNdex < nSize)
	{
		//处理不是回路的最后一个拐点
		if (iNdex == nSize - 1 && !bLoop)
		{
			int iTmp = delTheLastPtNoLoop(vecBreakData, vecStraigthData, ptArry);
			if (iTmp > 0)
	        {
				iNdex = iTmp;
				continue;
	        }
			break;
		}
		//处理拐点
		if (!getBreakBackData(mapIndexStaiData, vecBreakData, ptArry, iNdex, bLoop))
		{
			breakInsertData insertData = vecBreakData.back();
			iter = g_mapTurnBackData.find(insertData.backData);
			if (iter == g_mapTurnBackData.end())
				return FALSE;
			vecBreakData.pop_back();
			if (bLoop&&iNdex == 1)
			{
				if (!getLoopFirstBreakBackData(insertData, iter, ptArry))
					return FALSE;
				insertData.ptZero = ptArry[0];
				vecBreakData.push_back(insertData);
			}
			else if (!bLoop&&iNdex == 2)
			{
				if (!getNoLoopFirstBreakBackData(insertData, iter, ptArry))
					return FALSE;
				insertData.ptZero = ptArry[1];
				vecBreakData.push_back(insertData);
			}
			else
			{
				if (!getBreakBackData(insertData, mapIndexStaiData, iter, vecBreakData, ptArry, iNdex - 1,bLoop))
					return FALSE;
				insertData.ptZero = ptArry[iNdex - 1];
				vecBreakData.push_back(insertData);
			}
			continue;
		}
		if (iNdex == nSize -1&&bLoop)
		{
			AcGePoint3d ptTmp1 = ptArry.last() - vecBreakData.back().dNextLen * (ptArry.last() - ptArry[0]).normal();
			AcGePoint3d ptTmp2 = ptArry[0] + vecBreakData[0].dStartLen* (ptArry.last() - ptArry[0]).normal();
			if (!connectBackEdge(vecStraigthData, ptArry[0] - ptArry.last(), ptTmp1, ptTmp2, vecBreakData.back().dNextInLen - g_sideMinLen, vecBreakData[0].dStartInLen - g_sideMinLen, TRUE))
			{
// 				std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapTurnBackData.find(breakData.backData);
// 				if (iter == g_mapTurnBackData.end())
// 					return FALSE;
// 				return getBreakBackData(breakData, mapStarData, iter, vecBreakData, ptArry, iIndex);
			}
		}
		iNdex++;
	}
	CMapIndexStargthData::iterator iterStar = mapIndexStaiData.begin();
	for (iterStar; iterStar != mapIndexStaiData.end(); iterStar++)
	{
		vecStraigthData.insert(vecStraigthData.end(), iterStar->second.begin(), iterStar->second.end());
	}
	return TRUE;
}

int CRLDrawBackEdge::delTheLastPtNoLoop(std::vector<breakInsertData>& vecBreakData, std::vector<straigthInsertData>& vecStraigthData,const AcGePoint3dArray& ptArry)
{
	std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter;
	breakInsertData insertData = vecBreakData[0];
	backPtData pathPtDataTmp;
	pathPtDataTmp.ptStart = ptArry[1] - insertData.dStartLen * (ptArry[1] - ptArry[0]).normal();
	pathPtDataTmp.ptEnd = ptArry[0];
	pathPtDataTmp.bStartEnt = TRUE;
	pathPtDataTmp.bEndNextPt = FALSE;
	if (pathPtDataTmp.ptStart.isEqualTo(ptArry[0], frgGlobals::Gtol) 
		|| !((pathPtDataTmp.ptStart - ptArry[0]).normal()).isEqualTo((ptArry[1] - ptArry[0]).normal(), frgGlobals::Gtol))
	{
		if (pathPtDataTmp.ptStart.distanceTo(ptArry[0]) < g_groupMinLenout)
		{
			pathPtDataTmp.ptEnd = pathPtDataTmp.ptStart;
			pathPtDataTmp.ptStart = pathPtDataTmp.ptStart + (2 * EPRES)*(ptArry[1] - ptArry[0]).normal();
		}
	}
	if (((pathPtDataTmp.ptStart - pathPtDataTmp.ptEnd).normal()).isEqualTo((ptArry[1] - ptArry[0]).normal(), frgGlobals::Gtol))
	{
		AcGePoint3d ptInsertTmp = pathPtDataTmp.ptStart;
		if (!getHasStartStraightBackDatas(vecStraigthData, pathPtDataTmp, ptInsertTmp, insertData.dStartInLen - g_sideMinLen, TRUE, FALSE))
		{
			vecBreakData.clear();
			iter = g_mapTurnBackData.find(insertData.backData);
			if (iter == g_mapTurnBackData.end())
				return FALSE;
			if (!getNoLoopFirstBreakBackData(insertData, iter, ptArry))
				return FALSE;
			insertData.ptZero = ptArry[1];
			vecBreakData.push_back(insertData);
			return 2;
		}
	}

	insertData = vecBreakData.back();
	pathPtDataTmp.ptStart = ptArry[ptArry.length() - 2] + (ptArry.last() - ptArry[ptArry.length() - 2]).normal() * insertData.dNextLen;
	pathPtDataTmp.ptEnd = ptArry.last();
	if (pathPtDataTmp.ptEnd.isEqualTo(pathPtDataTmp.ptStart, frgGlobals::Gtol) ||
		!((pathPtDataTmp.ptEnd - pathPtDataTmp.ptStart).normal()).isEqualTo((ptArry.last() - ptArry[ptArry.length() - 2]).normal(), frgGlobals::Gtol))
	{
		if (pathPtDataTmp.ptStart.distanceTo(ptArry.last()) < g_groupMinLenout)
		{
			pathPtDataTmp.ptStart = pathPtDataTmp.ptStart;
			pathPtDataTmp.ptEnd = pathPtDataTmp.ptStart + (2 * EPRES)*(ptArry.last() - ptArry[ptArry.length() - 2]).normal();
		}
	}
	if (((pathPtDataTmp.ptEnd - pathPtDataTmp.ptStart).normal()).isEqualTo((ptArry.last() - ptArry[ptArry.length() - 2]).normal(), frgGlobals::Gtol))
	{
		AcGePoint3d ptInsertTmp = pathPtDataTmp.ptStart;
		if (!getHasStartStraightBackDatas(vecStraigthData, pathPtDataTmp, ptInsertTmp, insertData.dNextLen - g_sideMinLen, TRUE))
		{
			return ptArry.length() - 2;
		}
	}
	return -1;
}

BOOL CRLDrawBackEdge::getBreakBackData(CMapIndexStargthData& mapStarData, std::vector<breakInsertData>& vecBreakData, 
	const AcGePoint3dArray& ptArry, int iIndex, BOOL bLoop)
{
	int nPtLen = ptArry.length();
	BOOL bEndBreak = nPtLen - iIndex > 2;
	if (bLoop)
		bEndBreak = TRUE;
	breakInsertData breakData;
	BOOL bGetData = FALSE;
	if (bEndBreak)
	{
		if (!getBreakBackDataHasEnd(breakData, mapStarData, vecBreakData, bGetData, ptArry, iIndex, bLoop))
			return FALSE;
	}
	if (!bGetData)
	{
		if (!getBreakBackDataNoMaterEnd(breakData, mapStarData, vecBreakData, bGetData, ptArry, iIndex))
			return FALSE;
	}

	if (bGetData)
	{
		if (getBreakBackData(breakData, mapStarData, vecBreakData, ptArry, iIndex, bLoop))
		{
			breakData.ptZero = ptArry[iIndex];
			vecBreakData.push_back(breakData);
			return TRUE;
		}
	}
	return FALSE;
}

void CRLDrawBackEdge::getBreakData(breakInsertData&breakData, const AcGePoint3dArray& ptArry,
	const std::map<BackEdgeTypeData, AcDbObjectId>::iterator& iter,
	int iStartIndex, int iIndex, int iEndIndex, BOOL bAdjust, BOOL bNegeit)
{
	breakData.entBlockId = iter->second;
	breakData.backData = iter->first;
	AcGePoint3d ptInsert = ptArry[iIndex];
	AcGeVector3d vecInsertX, vecInsertY;
	if (bNegeit)
	{
		breakData.dStartInLen = breakData.dStartLen = iter->first.m_dLen2;
		breakData.dNextInLen = breakData.dNextLen = iter->first.m_dLen;

		vecInsertX = (ptArry[iEndIndex] - ptArry[iIndex]).normal();
		vecInsertY = (ptArry[iIndex] - ptArry[iStartIndex]).normal();
	}
	else
	{
		breakData.dStartInLen = breakData.dStartLen = iter->first.m_dLen;
		breakData.dNextInLen = breakData.dNextLen = iter->first.m_dLen2;
		vecInsertX = (ptArry[iStartIndex] - ptArry[iIndex]).normal();
		vecInsertY = (ptArry[iIndex] - ptArry[iEndIndex]).normal();
	}
	if (bAdjust)
	{
		ptInsert = ptInsert - vecInsertX * breakData.backData.m_dWith + vecInsertY * breakData.backData.m_dWith;
		breakData.dStartLen -= breakData.backData.m_dWith;
		breakData.dNextLen -= breakData.backData.m_dWith;
	}
	breakData.mat.setCoordSystem(ptInsert, vecInsertX, vecInsertY, vecInsertX.crossProduct(vecInsertY).normal());
}

BOOL CRLDrawBackEdge::getBreakBackDataNoMaterEnd(breakInsertData&breakData, CMapIndexStargthData& mapStarData, std::vector<breakInsertData>& vecBreakData,
	BOOL& bGetData, const AcGePoint3dArray& ptArry, int iIndex)
{
	std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapTurnBackData.begin();
	int nPtLen = ptArry.length();
	assert(iIndex > 0 && iIndex < nPtLen);
	int iStartIndex = iIndex - 1;
	int iEndIndex = iIndex + 1;
	if (iEndIndex >= nPtLen)
		iEndIndex = 0;

	AcGeVector3d vecXPer = ptArry[iEndIndex] - ptArry[iIndex];
	AcGeVector3d vecYPer = ptArry[iIndex] - ptArry[iStartIndex];
	AcGeVector3d vecXPer1 = vecXPer;
	vecXPer1.rotateBy(PI / 2.0, AcGeVector3d(0, 0, 1));
	vecXPer1.normalize();
	vecYPer.normalize();
	AcGePoint3d ptInsert = ptArry[iIndex];
	BOOL bAdjust = vecXPer1.isEqualTo(vecYPer, frgGlobals::Gtol);

	AcGeVector3d vecInsertX, vecInsertY;
	double dMinLen1 = 350;
	double dMinLen2 = 500;
	double dMinLen3 = 1000;

	double dStarTotalLen = ptArry[iStartIndex].distanceTo(ptArry[iIndex]);
	double dEndTotalLen = ptArry[iEndIndex].distanceTo(ptArry[iIndex]);
	if (!vecBreakData.empty())
		dStarTotalLen -= vecBreakData.back().dNextLen;

	if (bAdjust)
	{
		dStarTotalLen += 90.0;
		dEndTotalLen += 90.0;
	}
	if (dStarTotalLen < dMinLen1)
		return FALSE;

	if (dStarTotalLen < dMinLen2)
	{
		getBreakData(breakData, ptArry, iter, iStartIndex, iIndex, iEndIndex, bAdjust, TRUE);
		bGetData = TRUE;
	}
	if (dStarTotalLen < dMinLen3)
	{
		if (dEndTotalLen - dMinLen1 + g_groupMinLenout < EPRES/*dNextLen < 350 - g_groupMinLenout*/)
		{
			getBreakData(breakData, ptArry, iter, iStartIndex, iIndex, iEndIndex, bAdjust, FALSE);
			bGetData = TRUE;
		}
		else if (dEndTotalLen - dMinLen2 + g_groupMinLenout < EPRES
			|| dEndTotalLen - dMinLen3 + g_dBreakOutLen < EPRES/*dNextLen < 350 - g_groupMinLenout*/)
		{
			getBreakData(breakData, ptArry, iter, iStartIndex, iIndex, iEndIndex, bAdjust, TRUE);
			bGetData = TRUE;
		}
		else
		{
			iter++;
			getBreakData(breakData, ptArry, iter, iStartIndex, iIndex, iEndIndex, bAdjust, TRUE);
			bGetData = TRUE;
		}
	}
	else /*if (dStarTotalLen - dMinLen3 + g_groupMinLenout < -EPRES/ *dLen < 1000 - g_groupMinLenout* /)*/
	{
		if (dEndTotalLen - dMinLen2 + g_groupMinLenout < EPRES
			|| dEndTotalLen - dMinLen3 + g_dBreakOutLen < EPRES/*dNextLen < 500 - g_groupMinLenout*/)
		{
			iter++;
			getBreakData(breakData, ptArry, iter, iStartIndex, iIndex, iEndIndex, bAdjust, FALSE);
			bGetData = TRUE;
		}
		else
		{
			iter++;
			iter++;
			getBreakData(breakData, ptArry, iter, iStartIndex, iIndex, iEndIndex, bAdjust, TRUE);
			bGetData = TRUE;
		}
	}
	return TRUE;
}

BOOL CRLDrawBackEdge::getBreakBackDataHasEnd(breakInsertData&breakData, CMapIndexStargthData& mapStarData, 
	std::vector<breakInsertData>& vecBreakData, BOOL& bGetData, const AcGePoint3dArray& ptArry, int iIndex, BOOL bLoop)
{
	std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapTurnBackData.begin();
	int nPtLen = ptArry.length();
	assert(iIndex > 0 && iIndex < nPtLen);
	int iStartIndex = iIndex - 1;
	int iEndIndex = iIndex + 1;
	if (iEndIndex >= nPtLen)
		iEndIndex = 0;

	AcGeVector3d vecXPer = ptArry[iEndIndex] - ptArry[iIndex];
	AcGeVector3d vecYPer = ptArry[iIndex] - ptArry[iStartIndex];
	AcGeVector3d vecXPer1 = vecXPer;
	vecXPer1.rotateBy(PI / 2.0, AcGeVector3d(0, 0, 1));
	vecXPer1.normalize();
	vecYPer.normalize();
	AcGePoint3d ptInsert = ptArry[iIndex];
	BOOL bAdjust = vecXPer1.isEqualTo(vecYPer, frgGlobals::Gtol);

	AcGeVector3d vecInsertX, vecInsertY;
	double dMinLen1 = 350;
	double dMinLen2 = 500;
	double dMinLen3 = 1000;

	double dStarTotalLen = ptArry[iStartIndex].distanceTo(ptArry[iIndex]);
	double dEndTotalLen = ptArry[iEndIndex].distanceTo(ptArry[iIndex]);
	if (!vecBreakData.empty())
		dStarTotalLen -= vecBreakData.back().dNextLen;
	if (bLoop&&iEndIndex == 0)
		dEndTotalLen -= vecBreakData[0].dStartLen;
	if (bAdjust)
	{
		dStarTotalLen += 90.0;
		dEndTotalLen += 90.0;
	}
	if (dStarTotalLen < dMinLen1)
		return FALSE;

	if ((!bLoop&&dEndTotalLen < 2 * dMinLen1) || (bLoop&&dEndTotalLen < dMinLen1))
		return FALSE;
	else  if ((!bLoop&&dEndTotalLen < dMinLen1 + dMinLen2) || (bLoop&&dEndTotalLen < dMinLen2))
	{
		if (dStarTotalLen < dMinLen2)
			return FALSE;
		getBreakData(breakData, ptArry, iter, iStartIndex, iIndex, iEndIndex, bAdjust, FALSE);
		bGetData = TRUE;
	}
	else  if ((!bLoop&&dEndTotalLen < dMinLen3 + dMinLen1) || (bLoop&&dEndTotalLen < dMinLen3))
	{
		if (dStarTotalLen < dMinLen2)
		{
			getBreakData(breakData, ptArry, iter, iStartIndex, iIndex, iEndIndex, bAdjust, TRUE);
			bGetData = TRUE;
		}
		else if (dStarTotalLen < dMinLen3)
		{
			getBreakData(breakData, ptArry, iter, iStartIndex, iIndex, iEndIndex, bAdjust, FALSE);
			bGetData = TRUE;
		}
		else
		{
			iter++;
			getBreakData(breakData, ptArry, iter, iStartIndex, iIndex, iEndIndex, bAdjust, FALSE);
			bGetData = TRUE;
		}
	}
	return TRUE;
}

BOOL CRLDrawBackEdge::getBreakBackData(breakInsertData&breakData, CMapIndexStargthData& mapStarData, 
	const std::vector<breakInsertData>& vecBreakData, const AcGePoint3dArray& ptArry, int iIndex, BOOL bLoop)
{
	int nPtLen = ptArry.length();
	int iStartIndex = iIndex - 1;
	int iEndIndex = iIndex + 1;
	if (iEndIndex >= nPtLen)
		iEndIndex = 0;
	AcGePoint3d ptInsert = ptArry[iIndex];
	if (vecBreakData.empty())
	{
		backPtData pathPtDataTmp;
		pathPtDataTmp.ptStart = ptInsert - breakData.dStartLen * (ptArry[iIndex] - ptArry[iStartIndex]).normal();
		pathPtDataTmp.ptEnd = ptArry[0];
		pathPtDataTmp.bStartEnt = TRUE;
		pathPtDataTmp.bEndNextPt = FALSE;
		if (!((pathPtDataTmp.ptStart - ptArry[iStartIndex]).normal()).isEqualTo((ptArry[iIndex] - ptArry[iStartIndex]).normal(), frgGlobals::Gtol))
		{
			if (pathPtDataTmp.ptStart.distanceTo(ptArry[iStartIndex]) < g_groupMinLenout)
			{
				pathPtDataTmp.ptEnd = pathPtDataTmp.ptStart;
				pathPtDataTmp.ptStart = pathPtDataTmp.ptStart + (2 * EPRES)*(ptArry[iIndex] - ptArry[iStartIndex]).normal();
			}
		}
		if (!((pathPtDataTmp.ptStart - pathPtDataTmp.ptEnd).normal()).isEqualTo((ptArry[iIndex] - ptArry[iStartIndex]).normal(), frgGlobals::Gtol))
			return FALSE;
		std::vector<straigthInsertData> vecStraigthData;
		if (!getHasStartStraightBackDatas(vecStraigthData, pathPtDataTmp, ptInsert, breakData.dStartInLen - g_sideMinLen, TRUE, FALSE))
		{
			std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapTurnBackData.find(breakData.backData);
			if (iter == g_mapTurnBackData.end())
				return FALSE;
			return getBreakBackData(breakData, mapStarData, iter, vecBreakData, ptArry, iIndex, bLoop);
		}
	}
	else
	{
		AcGePoint3d ptTmp1 = ptInsert - breakData.dStartLen * (ptArry[iIndex] - ptArry[iStartIndex]).normal();
		AcGePoint3d ptTmp2 = ptArry[iIndex - 1] + vecBreakData.back().dNextLen* (ptArry[iIndex] - ptArry[iStartIndex]).normal();
		std::vector<straigthInsertData> vecStraigthData;
		if (!connectBackEdge(vecStraigthData, ptArry[iIndex] - ptArry[iStartIndex], ptTmp2, ptTmp1, vecBreakData.back().dNextInLen - g_sideMinLen, breakData.dStartInLen - g_sideMinLen, TRUE))
		{
			std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapTurnBackData.find(breakData.backData);
			if (iter == g_mapTurnBackData.end())
				return FALSE;
			return getBreakBackData(breakData, mapStarData, iter, vecBreakData, ptArry, iIndex, bLoop);
		}
		else
		{
			mapStarData[iIndex] = vecStraigthData;
		}
	}
	return TRUE;
}

BOOL CRLDrawBackEdge::getBreakBackData(breakInsertData&breakData, CMapIndexStargthData& mapStarData, std::map<BackEdgeTypeData, AcDbObjectId>::iterator& iter,
	const std::vector<breakInsertData>& vecBreakData, const AcGePoint3dArray& ptArry, int iIndex, BOOL bLoop)
{
	int nPtLen = ptArry.length();
	assert(iIndex > 0 && iIndex < nPtLen);
	BOOL bEndBreak = nPtLen - iIndex > 2;
	if (bLoop)
		bEndBreak = TRUE;
	int iStartIndex = iIndex - 1;
	int iEndIndex = iIndex + 1;
	if (iEndIndex >= nPtLen)
		iEndIndex = 0;
	double outLen = vecBreakData.empty() ? g_groupMinLenout : 0.0;
	AcGeVector3d vecXPer = ptArry[iEndIndex] - ptArry[iIndex];
	AcGeVector3d vecYPer = ptArry[iIndex] - ptArry[iStartIndex];
	AcGeVector3d vecXPer1 = vecXPer;
	vecXPer1.rotateBy(PI / 2.0, AcGeVector3d(0, 0, 1));
	vecXPer1.normalize();
	vecYPer.normalize();
	AcGePoint3d ptInsert = ptArry[iIndex];
	BOOL bAdjust = vecXPer1.isEqualTo(vecYPer, frgGlobals::Gtol);

	AcGeVector3d vecInsertX, vecInsertY;
	double dMinLen1 = 350;
	double dMinLen2 = 500;
	double dMinLen3 = 1000;

	double dStarTotalLen = ptArry[iStartIndex].distanceTo(ptArry[iIndex]);
	double dEndTotalLen = ptArry[iEndIndex].distanceTo(ptArry[iIndex]);
	if (!vecBreakData.empty())
		dStarTotalLen -= vecBreakData.back().dNextLen;
	if (bLoop&&iEndIndex == 0)
		dEndTotalLen -= vecBreakData[0].dStartLen;

	if (bAdjust)
	{
		dStarTotalLen += 90.0;
		dEndTotalLen += 90.0;
	}
	BOOL bGetData = FALSE;
	if (dStarTotalLen < dMinLen1)
		return FALSE;
	if (breakData.dNextLen > breakData.dStartLen)
	{
		if (breakData.dNextLen > dStarTotalLen)
			bGetData = FALSE;
		else
		{
			std::swap(breakData.dStartLen, breakData.dNextLen);
			std::swap(breakData.dStartInLen, breakData.dNextInLen);
			AcGeVector3d vecInsertX, vecInsertZ, vecInsertY;
			breakData.mat.getCoordSystem(ptInsert, vecInsertX, vecInsertY, vecInsertZ);
			breakData.mat.setCoordSystem(ptInsert, vecInsertY.normal(), vecInsertX.normal(), vecInsertY.crossProduct(vecInsertX).normal());
			if (getBreakBackData(breakData, mapStarData,vecBreakData, ptArry, iIndex,bLoop))
				return TRUE;
		}
	}
	iter--;
	if (iter == g_mapTurnBackData.end())
		return FALSE;
	if (iter->first.m_dLen - dStarTotalLen < EPRES)
	{
		if (bEndBreak&&dEndTotalLen > iter->first.m_dLen2)
			bGetData = FALSE;
		else
		{
			getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, FALSE);
			bGetData = TRUE;
		}
	}
	else if (iter->first.m_dLen2 - dStarTotalLen < EPRES)
	{
		if (bEndBreak&&dEndTotalLen > iter->first.m_dLen2)
			bGetData = FALSE;
		else
		{
			getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, TRUE);
			bGetData = TRUE;
		}
	}
	if (!bGetData)
		return FALSE;
	return getBreakBackData(breakData, mapStarData,vecBreakData, ptArry, iIndex,bLoop);
}

BOOL CRLDrawBackEdge::getNoLoopFirstBreakBackData(breakInsertData&breakData, std::map<BackEdgeTypeData, AcDbObjectId>::iterator& iter, const AcGePoint3dArray& ptArry)
{
	int nPtLen = ptArry.length();
	BOOL bEndBreak = nPtLen > 3;

	AcGeVector3d vecXPer = ptArry[2] - ptArry[1];
	AcGeVector3d vecYPer = ptArry[1] - ptArry[0];
	AcGeVector3d vecXPer1 = vecXPer;
	vecXPer1.rotateBy(PI / 2.0, AcGeVector3d(0, 0, 1));
	vecXPer1.normalize();
	vecYPer.normalize();
	AcGePoint3d ptInsert = ptArry[1];
	BOOL bAdjust = vecXPer1.isEqualTo(vecYPer, frgGlobals::Gtol);

	AcGeVector3d vecInsertX, vecInsertY;
	double dMinLen1 = 350;
	double dMinLen2 = 500;
	double dMinLen3 = 1000;

	double dStarTotalLen = ptArry[1].distanceTo(ptArry[0]);
	double dEndTotalLen = ptArry[1].distanceTo(ptArry[2]);
	if (bAdjust)
	{
		dStarTotalLen += 90.0;
		dEndTotalLen += 90.0;
	}
	BOOL bGetData = FALSE;
	if (bEndBreak&&dEndTotalLen < 2 * dMinLen1)
		return FALSE;
	if (breakData.dNextLen > breakData.dStartLen)
	{
		std::swap(breakData.dStartLen, breakData.dNextLen);
		std::swap(breakData.dStartInLen, breakData.dNextInLen);
		AcGeVector3d vecInsertX, vecInsertZ, vecInsertY;
		breakData.mat.getCoordSystem(ptInsert, vecInsertX, vecInsertY, vecInsertZ);
		breakData.mat.setCoordSystem(ptInsert, vecInsertY.normal(), vecInsertX.normal(), vecInsertY.crossProduct(vecInsertX).normal());
		return TRUE;
	}
	iter--;
	if (iter == g_mapTurnBackData.end())
		return FALSE;
	if (bEndBreak)
	{
		if (dEndTotalLen < iter->first.m_dLen2 + dMinLen1)
			return FALSE;
		if (dEndTotalLen - iter->first.m_dLen - dMinLen1 >-EPRES)
		{
			getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, TRUE);
			bGetData = TRUE;
		}
		else
		{
			getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, TRUE);
			bGetData = TRUE;
		}
	}
	else
	{
		getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, TRUE);
		bGetData = TRUE;
	}
	return bGetData;
}

BOOL CRLDrawBackEdge::getNoLoopFirstBreakBackData(std::vector<breakInsertData>& vecBreakData, const AcGePoint3dArray& ptArry)
{
	int nPtLen = ptArry.length();
	BOOL bEndBreak = nPtLen > 3;

	AcGeVector3d vecXPer = ptArry[2] - ptArry[1];
	AcGeVector3d vecYPer = ptArry[1] - ptArry[0];
	AcGeVector3d vecXPer1 = vecXPer;
	vecXPer1.rotateBy(PI / 2.0, AcGeVector3d(0, 0, 1));
	vecXPer1.normalize();
	vecYPer.normalize();
	AcGePoint3d ptInsert = ptArry[1];
	BOOL bAdjust = vecXPer1.isEqualTo(vecYPer, frgGlobals::Gtol);

	AcGeVector3d vecInsertX, vecInsertY;
	double dMinLen1 = 350;
	double dMinLen2 = 500;
	double dMinLen3 = 1000;

	double dStarTotalLen = ptArry[1].distanceTo(ptArry[0]);
	double dEndTotalLen = ptArry[1].distanceTo(ptArry[2]);
	if (bAdjust)
	{
		dStarTotalLen += 90.0;
		dEndTotalLen += 90.0;
	}
	breakInsertData breakData;
	std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapTurnBackData.begin();
	BOOL bGetData = FALSE;
	if (bEndBreak)
	{
		if (dEndTotalLen < 2 * dMinLen1)
			return FALSE;
		else if (dEndTotalLen < dMinLen1 + dMinLen2)
		{
			getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, FALSE);
			bGetData = TRUE;
		}
		else if (dEndTotalLen < dMinLen3 + dMinLen1)
		{
			if (dStarTotalLen - dMinLen1 + g_groupMinLenout < EPRES/*dLen < 350 - g_groupMinLenout*/)
			{
				getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, TRUE);
				bGetData = TRUE;
			}
			else if (dStarTotalLen - dMinLen2 + g_groupMinLenout < EPRES/*dLen < 500 - g_groupMinLenout*/
				|| dStarTotalLen - dMinLen3 + g_dBreakOutLen < EPRES)
			{
				getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, FALSE);
				bGetData = TRUE;
			}
			else 
			{
				iter++;
				getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, FALSE);
				bGetData = TRUE;
			}
		}
	}
	if (!bGetData)
	{
		if (dStarTotalLen - dMinLen1 + g_groupMinLenout < EPRES/*dLen < 350 - g_groupMinLenout*/)
		{
			getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, TRUE);
			bGetData = TRUE;
		}
		else if (dStarTotalLen - dMinLen2 + g_groupMinLenout < EPRES
			|| dStarTotalLen - dMinLen3 + g_dBreakOutLen < EPRES/*dLen < 500 - g_groupMinLenout*/)
		{
			if (dEndTotalLen - dMinLen1 + g_groupMinLenout < EPRES/*dNextLen < 350 - g_groupMinLenout*/)
			{
				getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, FALSE);
				bGetData = TRUE;
			}
			else if (dEndTotalLen - dMinLen3 + g_dBreakOutLen < EPRES)
			{
				getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, TRUE);
				bGetData = TRUE;
			}
			else
			{
				iter++;
				getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, TRUE);
				bGetData = TRUE;
			}
		}
		else /*if (dStarTotalLen - dMinLen3 + g_groupMinLenout < -EPRES/ *dLen < 1000 - g_groupMinLenout* /)*/
		{
			if (dEndTotalLen - dMinLen2 + g_groupMinLenout < EPRES 
				|| dEndTotalLen - dMinLen3 + g_dBreakOutLen < EPRES/*dNextLen < 500 - g_groupMinLenout*/)
			{
				iter++;
				getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, FALSE);
				bGetData = TRUE;
			}
			else
			{
				iter++;
				iter++;
				getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, TRUE);
				bGetData = TRUE;
			}
		}
	}
	if (bGetData)
	{
		breakData.ptZero = ptArry[1];
		vecBreakData.push_back(breakData);
	}
	return bGetData;
}

BOOL CRLDrawBackEdge::getLoopFirstBreakBackData(breakInsertData&breakData, std::map<BackEdgeTypeData, AcDbObjectId>::iterator& iter, const AcGePoint3dArray& ptArry)
{
	int nPtLen = ptArry.length();
	AcGeVector3d vecXPer = ptArry[1] - ptArry[0];
	AcGeVector3d vecYPer = ptArry[0] - ptArry[nPtLen-1];
	AcGeVector3d vecXPer1 = vecXPer;
	vecXPer1.rotateBy(PI / 2.0, AcGeVector3d(0, 0, 1));
	vecXPer1.normalize();
	vecYPer.normalize();
	AcGePoint3d ptInsert = ptArry[0];
	BOOL bAdjust = vecXPer1.isEqualTo(vecYPer, frgGlobals::Gtol);
	AcGeVector3d vecInsertX, vecInsertY;
	double dMinLen1 = 350;
	double dMinLen2 = 500;
	double dMinLen3 = 1000;

	double dStarTotalLen = ptArry[nPtLen - 1].distanceTo(ptArry[0]);
	double dEndTotalLen = ptArry[1].distanceTo(ptArry[0]);
	if (bAdjust)
	{
		dStarTotalLen += 90.0;
		dEndTotalLen += 90.0;
	}
	if (dStarTotalLen < 2 * dMinLen1 || dEndTotalLen < 2 * dMinLen1)
		return FALSE;
	else if (dStarTotalLen < dMinLen1 + dMinLen2&&dEndTotalLen < dMinLen1 + dMinLen2)
		return FALSE;
	BOOL bGetData = FALSE;
	if (breakData.dNextLen > breakData.dStartLen)
	{
		if (dStarTotalLen >breakData.dNextLen + dMinLen1)
			bGetData = FALSE;
		else
		{
			std::swap(breakData.dStartLen, breakData.dNextLen);
			std::swap(breakData.dStartInLen, breakData.dNextInLen);
			AcGeVector3d vecInsertX, vecInsertZ, vecInsertY;
			breakData.mat.getCoordSystem(ptInsert, vecInsertX, vecInsertY, vecInsertZ);
			breakData.mat.setCoordSystem(ptInsert, vecInsertY.normal(), vecInsertX.normal(), vecInsertY.crossProduct(vecInsertX).normal());
			return TRUE;
		}
	}
	iter--;
	if (iter == g_mapTurnBackData.end())
		return FALSE;
	if (dStarTotalLen < iter->first.m_dLen + dMinLen1&&dStarTotalLen < iter->first.m_dLen2 + dMinLen1)
		return FALSE;
	else if (dStarTotalLen > iter->first.m_dLen +dMinLen1)
	{
		getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, FALSE);
		bGetData = TRUE;
	}
	else
	{
		getBreakData(breakData, ptArry, iter, 0, 1, 2, bAdjust, TRUE);
		bGetData = TRUE;
	}
	return bGetData;
}

BOOL CRLDrawBackEdge::getLoopFirstBreakBackData(std::vector<breakInsertData>& vecBreakData, const AcGePoint3dArray& ptArry)
{
	int nPtLen = ptArry.length();
	AcGeVector3d vecXPer = ptArry[1] - ptArry[0];
	AcGeVector3d vecYPer = ptArry[0] - ptArry[nPtLen-1];
	AcGeVector3d vecXPer1 = vecXPer;
	vecXPer1.rotateBy(PI / 2.0, AcGeVector3d(0, 0, 1));
	vecXPer1.normalize();
	vecYPer.normalize();
	AcGePoint3d ptInsert = ptArry[0];
	BOOL bAdjust = vecXPer1.isEqualTo(vecYPer, frgGlobals::Gtol);
	AcGeVector3d vecInsertX, vecInsertY;
	double dMinLen1 = 350;
	double dMinLen2 = 500;
	double dMinLen3 = 1000;

	double dStarTotalLen = ptArry[nPtLen - 1].distanceTo(ptArry[0]);
	double dEndTotalLen = ptArry[1].distanceTo(ptArry[0]);
	if (bAdjust)
	{
		dStarTotalLen += 90.0;
		dEndTotalLen += 90.0;
	}
	breakInsertData breakData;
	if (dStarTotalLen < 2 * dMinLen1 || dEndTotalLen < 2 * dMinLen1)
		return FALSE;
	else if (dStarTotalLen < dMinLen1 + dMinLen2&&dEndTotalLen < dMinLen1 + dMinLen2)
		return FALSE;

	if (dStarTotalLen < dMinLen1 + dMinLen2)
	{
		std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapTurnBackData.begin();
		getBreakData(breakData, ptArry, iter, nPtLen - 1, 0, 1, bAdjust, TRUE);
	}
	else if (dStarTotalLen < dMinLen1 + dMinLen3)
	{
		std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapTurnBackData.begin();
		if (dEndTotalLen < dMinLen1 + dMinLen2)
		{
			getBreakData(breakData, ptArry, iter, nPtLen - 1, 0, 1, bAdjust, FALSE);
		}
		else if (dEndTotalLen < dMinLen1+dMinLen3)
		{
			getBreakData(breakData, ptArry, iter, nPtLen - 1, 0, 1, bAdjust, TRUE);
		}
		else
		{
			iter++;
			getBreakData(breakData, ptArry, iter, nPtLen - 1, 0, 1, bAdjust, TRUE);
		}
	}
	else
	{
		std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapTurnBackData.begin();
		if (dEndTotalLen < dMinLen1 + dMinLen2)
		{
			getBreakData(breakData, ptArry, iter, nPtLen - 1, 0, 1, bAdjust, FALSE);
		}
		else if (dEndTotalLen < dMinLen1 + dMinLen3)
		{
			iter++;
			getBreakData(breakData, ptArry, iter, nPtLen - 1, 0, 1, bAdjust, FALSE);
		}
		else
		{
			iter++;
			iter++;
			getBreakData(breakData, ptArry, iter, nPtLen - 1, 0, 1, bAdjust, FALSE);
		}
	}
	breakData.ptZero = ptArry[0];
	vecBreakData.push_back(breakData);
	return TRUE;
}

BOOL CRLDrawBackEdge::getStraightBackDatas(std::vector<straigthInsertData>& vecStraigthData, const AcGePoint3dArray& ptArry)
{
	if (ptArry.length() != 2)
		return FALSE;

	double dLen = ptArry[0].distanceTo(ptArry[1]);
	if (dLen < EPRES)
		return FALSE;
	AcGeVector3d vecInsertX = ptArry[1] - ptArry[0];
	vecInsertX.normalize();
	AcGeVector3d vecInsertY = vecInsertX;
	vecInsertY.rotateBy(PI / 2.0, AcGeVector3d(0, 0, 1));
	vecInsertY.normalize();
	BOOL bNext = FALSE;
	AcGePoint3d ptInsert = ptArry[0];
	BackEdgeTypeData backEdgedata;
	AcDbObjectId blockId = getBlockIdNoStartNoEnd(backEdgedata, ptInsert, bNext, ptArry[0], ptArry[1]);
	if (!blockId.isValid())
		return FALSE;
	if (!backEdgedata.m_bBlue)
		ptInsert = ptInsert + vecInsertY*((g_dBlueWith - g_dGreenWith) / 2.0);
	AcGeMatrix3d mat;
	mat.setCoordSystem(ptInsert, vecInsertX.normal(), vecInsertY.normal(), vecInsertX.crossProduct(vecInsertY).normal());
	straigthInsertData insertData;
	insertData.backData = backEdgedata;
	insertData.mat = mat;
	insertData.entBlockId = blockId;
	vecStraigthData.push_back(insertData);

	if (!backEdgedata.m_bBlue)
		ptInsert = ptInsert - vecInsertY*((g_dBlueWith - g_dGreenWith) / 2.0);
	if (bNext)
	{
		backPtData ptDataTmp;
		ptDataTmp.ptStart = ptInsert + vecInsertX * backEdgedata.m_dLen;
		ptDataTmp.ptEnd = ptArry[1];
		if (!((ptDataTmp.ptEnd - ptDataTmp.ptStart).normal()).isEqualTo(vecInsertX.normal(), frgGlobals::Gtol))
			return FALSE;
		if (!getHasStartStraightBackDatas(vecStraigthData, ptDataTmp, ptInsert, backEdgedata.m_dLen - g_sideMinLen, backEdgedata.m_bBlue))
			return FALSE;
	}
	return TRUE;
}

BOOL CRLDrawBackEdge::getHasStartStraightBackDatas(std::vector<straigthInsertData>& vecStraigthData, const backPtData& pathPtData,
	                                              AcGePoint3d& ptInsert, double dInLen, BOOL bBlue, BOOL bBeginStart)
{
	if (dInLen < g_minLenIn)
		return FALSE;
	double dLen = pathPtData.ptStart.distanceTo(pathPtData.ptEnd);
	if (dLen < EPRES)
	   return FALSE;
	AcGeVector3d vecInsertX = pathPtData.ptEnd - pathPtData.ptStart;
	vecInsertX.normalize();
	BOOL bNext = FALSE;
	AcDbObjectId blockId = AcDbObjectId::kNull;
	AcDbObjectId blockId2 = AcDbObjectId::kNull;
	AcGePoint3d ptInsert2;
	BackEdgeTypeData backEdgedata;
	BackEdgeTypeData backEdgedata2;
	if (bBlue)
	{
		if (!getBlockIdConBlueStraigth(blockId, backEdgedata, ptInsert, blockId2, backEdgedata2,ptInsert2, bNext, pathPtData, dInLen))
			return FALSE;
	}
	else
	{
		if (!getBlockIdConGreenStraigth(blockId, backEdgedata, ptInsert, blockId2, backEdgedata2,ptInsert2, bNext, pathPtData, dInLen))
			return FALSE;
	}
	if (!blockId.isValid())
		return FALSE;

	AcGeVector3d vecInsertY = vecInsertX;
	if (!bBeginStart)
		vecInsertY = pathPtData.ptStart - pathPtData.ptEnd;
	vecInsertY.rotateBy(PI / 2.0, AcGeVector3d(0, 0, 1));
	vecInsertY.normalize();
	AcGeMatrix3d mat;
	AcGeMatrix3d mat2;
	if (!backEdgedata.m_bBlue)
		ptInsert = ptInsert + vecInsertY*((g_dBlueWith - g_dGreenWith) / 2.0);
	mat.setCoordSystem(ptInsert, vecInsertX.normal(), vecInsertY.normal(), vecInsertX.crossProduct(vecInsertY).normal());

	straigthInsertData insertData;
	if (blockId2.isValid())
	{
		if (!backEdgedata2.m_bBlue)
			ptInsert2 = ptInsert2 + vecInsertY*((g_dBlueWith - g_dGreenWith) / 2.0);
		mat2.setCoordSystem(ptInsert2, vecInsertX.normal(), vecInsertY.normal(), vecInsertX.crossProduct(vecInsertY).normal());

		insertData.backData = backEdgedata2;
		insertData.mat = mat2;
		insertData.entBlockId = blockId2;
		vecStraigthData.push_back(insertData);
	}
	
	insertData.backData = backEdgedata;
	insertData.mat = mat;
	insertData.entBlockId = blockId;
	vecStraigthData.push_back(insertData);

	
	if (!backEdgedata.m_bBlue)
		ptInsert = ptInsert - vecInsertY*((g_dBlueWith - g_dGreenWith) / 2.0);
	if (bNext)
	{
		backPtData ptDataTmp;
		double dLenTmp = 2000;
		if (backEdgedata.m_bBlue)
			dLenTmp = 2500;
		ptDataTmp.ptStart = ptInsert + vecInsertX * dLenTmp;
		ptDataTmp.ptEnd = pathPtData.ptEnd;
		if (!getHasStartStraightBackDatas(vecStraigthData, ptDataTmp, ptInsert, dLenTmp - g_sideMinLen, backEdgedata.m_bBlue, bBeginStart))
			return FALSE;
	}
	return TRUE;
}

BOOL CRLDrawBackEdge::getBlockIdConBlueStraigth(AcDbObjectId& blockId, BackEdgeTypeData& backEdgedata, AcGePoint3d& ptInsert,
	AcDbObjectId& blockId2, BackEdgeTypeData& backEdgedata2, AcGePoint3d& ptInsert2, BOOL& bNext, const backPtData& pathPtData, double dInLen)
{
	if (dInLen < g_minLenIn)
		return FALSE;

	AcGeVector3d vecInsertX = pathPtData.ptEnd - pathPtData.ptStart;
	vecInsertX.normalize();
	double dLen = pathPtData.ptEnd.distanceTo(pathPtData.ptStart);
	std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapStraightBackData.begin();
	bNext = FALSE;
	for (iter; iter != g_mapStraightBackData.end(); iter++)
	{
		backEdgedata = iter->first;
		double dBackLen = iter->first.m_dLen;
		if (backEdgedata.m_bBlue)
			continue;
		if (dLen - dBackLen + g_groupMinLenout + g_minLenIn < EPRES/*dLen < dBackLen - g_groupMinLenout - g_minLenIn*/)
		{
			double dTmpLen = dBackLen - dLen - g_standlLenout;
			blockId = iter->second;
			if (dTmpLen - g_minLenIn < EPRES/*dTmpLen < g_minLenIn*/)
				ptInsert = pathPtData.ptStart - vecInsertX*g_minLenIn;
			else if (dTmpLen < dInLen)
				ptInsert = pathPtData.ptStart - vecInsertX*dTmpLen;
			else
				ptInsert = pathPtData.ptStart - vecInsertX*dInLen;
			break;
		}
	}
	if (!blockId.isValid())
	{
		int nCount = 0;
		std::vector<std::map<BackEdgeTypeData, AcDbObjectId>::iterator> vecIter;
		getMapStraightIter(vecIter);
		assert(vecIter.size() == 6);
		while (nCount < 8)
		{
			nCount++;
			std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter2 = g_mapStraightBackData.end();
			std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapStraightBackData.end();
			switch (nCount)
			{
			case  1:
				iter2 = vecIter[0];
				iter = vecIter[1];
				break; //(500B-1000A)
			case  2:
				iter2 = vecIter[2];
				iter = vecIter[1]; //(1000B-1000A)
				break;
			case  3:
				iter2 = vecIter[3];
				iter = vecIter[1]; //(1500B-1000A)
				break;
			case  4:
				iter2 = vecIter[0];
				iter = vecIter[5]; //(500B-2500A)
				break;
			case  5:
				iter2 = vecIter[4];
				iter = vecIter[1]; //(2000B-1000A)
				break;
			case  6:
				iter2 = vecIter[2];
				iter = vecIter[5]; //(1000B-2500A)
				break;
			case  7:
				iter2 = vecIter[3];
				iter = vecIter[5];  //(1500B-2500A)
				break;
			case  8:
				iter2 = vecIter[4];
				iter = vecIter[5];//(2000B-2500A)
				break;
			default:
				break;
			}
			if (iter2 == g_mapStraightBackData.end() || iter == g_mapStraightBackData.end()
				|| iter2->first.m_bBlue || !iter->first.m_bBlue)
				continue;
			backEdgedata = iter2->first;
			double dTotalLen = iter2->first.m_dLen + iter->first.m_dLen;
			if (dLen - dTotalLen + 2 * g_minLenIn + g_groupMinLenout < EPRES/*dLen < dTotalLen - 2 * g_minLenIn - g_groupMinLenout*/)
			{
				double dTmpInlen = dInLen;
				if (dTmpInlen > iter2->first.m_dLen - 2 * g_minLenIn)
					dTmpInlen = iter2->first.m_dLen - 2 * g_minLenIn;
				double dTmpLen = dTotalLen - dLen;
				double dTmp = dTmpLen - 2 * g_minLenIn;
				blockId = iter2->second;
				blockId2 = iter->second;
				backEdgedata2 = iter->first;
				if (dTmp - g_standlLenout < EPRES)
				{
					ptInsert = pathPtData.ptStart - vecInsertX*g_minLenIn;
					ptInsert2 = ptInsert + vecInsertX*(iter2->first.m_dLen - g_minLenIn);
				}
				else
				{
					dTmp = dTmpLen - g_standlLenout;
					dTmp /= 2;
					if (dTmp > dTmpInlen)
						dTmp = dTmpInlen;
					double dElseInLen = dTmpLen - g_standlLenout - dTmp;
					if (dElseInLen > iter->first.m_dLen - g_minLenIn)
						dElseInLen = iter->first.m_dLen - g_minLenIn;
					dTmp = dTmpLen - g_standlLenout - dElseInLen;
					if (dTmp > dTmpInlen)
						return FALSE;
					ptInsert = pathPtData.ptStart - vecInsertX*dTmp;
					ptInsert2 = ptInsert + vecInsertX*(iter2->first.m_dLen - dElseInLen);
					if (!(ptInsert2 - pathPtData.ptStart).normal().isEqualTo(vecInsertX, frgGlobals::Gtol))
					{
// 						AcGePoint3d ptEndTmp = ptInsert2 + vecInsertX*iter->first.m_dLen;
// 						if (ptEndTmp.distanceTo(pathPtData.ptEnd) > g_dBreakOutLen)
							continue;
						ptInsert = pathPtData.ptStart - vecInsertX*g_minLenIn;
						ptInsert2 = ptInsert + vecInsertX*(iter2->first.m_dLen - dTmpInlen);
					}
				}
				break;
			}
		}
		if (!blockId.isValid())
		{
			double dTmpLen = g_minLenIn;
			if (dTmpLen > dInLen)
				dTmpLen = dInLen;
			ptInsert = pathPtData.ptStart - vecInsertX*dTmpLen;
			std::map<BackEdgeTypeData, AcDbObjectId>::reverse_iterator iter2 = g_mapStraightBackData.rbegin();
			iter2++;
			backEdgedata = iter2->first;
			blockId = iter2->second;
			bNext = TRUE;
		}
	}
	return TRUE;
}

BOOL CRLDrawBackEdge::getBlockIdConGreenStraigth(AcDbObjectId& blockId, BackEdgeTypeData& backEdgedata, AcGePoint3d& ptInsert,
	AcDbObjectId& blockId2, BackEdgeTypeData& backEdgedata2, AcGePoint3d& ptInsert2, BOOL& bNext, const backPtData& pathPtData, double dInLen)
{
	if (dInLen < g_minLenIn)
		return FALSE;
	AcGeVector3d vecInsertX = pathPtData.ptEnd - pathPtData.ptStart;
	vecInsertX.normalize();
	double dLen = pathPtData.ptEnd.distanceTo(pathPtData.ptStart);
	std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapStraightBackData.begin();
	for (iter; iter != g_mapStraightBackData.end(); iter++)
	{
		backEdgedata = iter->first;
		double dBackLen = iter->first.m_dLen;
		if (!backEdgedata.m_bBlue)
			continue;
		if (dLen - dBackLen + g_groupMinLenout + g_minLenIn < EPRES/*dLen < dBackLen - g_groupMinLenout - g_minLenIn*/)
		{
			double dTmpLen = dBackLen - dLen - g_standlLenout;
			blockId = iter->second;
			if (dTmpLen - g_minLenIn < EPRES/*dTmpLen < g_minLenIn*/)
				ptInsert = pathPtData.ptStart - vecInsertX*g_minLenIn;
			else if (dTmpLen < dInLen)
				ptInsert = pathPtData.ptStart - vecInsertX*dTmpLen;
			else
				ptInsert = pathPtData.ptStart - vecInsertX*dInLen;
			break;
		}
	}
	if (!blockId.isValid())
	{
		int nCount = 0;
		std::vector<std::map<BackEdgeTypeData, AcDbObjectId>::iterator> vecIter;
		getMapStraightIter(vecIter);
		assert(vecIter.size() == 6);
		while (nCount < 8)
		{
			nCount++;
			std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter2 = g_mapStraightBackData.end();
			std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapStraightBackData.end();
			switch (nCount)
			{
			case  1:
				iter2 = vecIter[1];
				iter = vecIter[0];
				break; //(500B-1000A)
			case  2:
				iter2 = vecIter[1];
				iter = vecIter[2]; //(1000B-1000A)
				break;
			case  3:
				iter2 = vecIter[1];
				iter = vecIter[3]; //(1500B-1000A)
				break;
			case  4:
				iter2 = vecIter[5];
				iter = vecIter[0]; //(500B-2500A)
				break;
			case  5:
				iter2 = vecIter[1];
				iter = vecIter[4]; //(2000B-1000A)
				break;
			case  6:
				iter2 = vecIter[5];
				iter = vecIter[2];  //(1000B-2500A)
				break;
			case  7:
				iter2 = vecIter[5];
				iter = vecIter[3]; //(1500B-2500A)
				break;
			case  8:
				iter2 = vecIter[5];
				iter = vecIter[4]; //(2000B-2500A)
				break;
			default:
				break;
			}
			if (iter2 == g_mapStraightBackData.end() || iter == g_mapStraightBackData.end()
				|| iter->first.m_bBlue || !iter2->first.m_bBlue)
				continue;
			backEdgedata = iter2->first;
			double dBackLen = backEdgedata.m_dLen;
			double dTotalLen = iter2->first.m_dLen + iter->first.m_dLen;
			if (dLen - dTotalLen + 2 * g_minLenIn + g_groupMinLenout < EPRES/*dLen < dTotalLen - 2 * g_minLenIn - g_groupMinLenout*/)
			{
				double dTmpInlen = dInLen;
				if (dTmpInlen > iter2->first.m_dLen - 2* g_minLenIn)
					dTmpInlen = iter2->first.m_dLen - 2 * g_minLenIn;
				double dTmpLen = dTotalLen - dLen;
				double dTmp = dTmpLen - 2 * g_minLenIn;
				blockId = iter2->second;
				blockId2 = iter->second;
				backEdgedata2 = iter->first;
				if (dTmp - g_standlLenout < EPRES)
				{
					ptInsert = pathPtData.ptStart - vecInsertX*g_minLenIn;
					ptInsert2 = ptInsert + vecInsertX*(iter2->first.m_dLen - g_minLenIn);
				}
				else
				{
					dTmp = dTmpLen - g_standlLenout;
					dTmp /= 2;
					if (dTmp > dTmpInlen)
						dTmp = dTmpInlen;
					double dElseInLen = dTmpLen - g_standlLenout - dTmp;
					if (dElseInLen > iter->first.m_dLen - g_minLenIn)
						dElseInLen = iter->first.m_dLen - g_minLenIn;
					dTmp = dTmpLen - g_standlLenout - dElseInLen;
					if (dTmp > dTmpInlen)
						return FALSE;

					ptInsert = pathPtData.ptStart - vecInsertX*dTmp;
					ptInsert2 = ptInsert + vecInsertX*(iter2->first.m_dLen - dElseInLen);
					if (!(ptInsert2 - pathPtData.ptStart).normal().isEqualTo(vecInsertX, frgGlobals::Gtol))
					{
// 						AcGePoint3d ptEndTmp = ptInsert2 + vecInsertX*iter->first.m_dLen;
// 						if (ptEndTmp.distanceTo(pathPtData.ptEnd) > g_dBreakOutLen)
							continue;
						ptInsert = pathPtData.ptStart - vecInsertX*g_minLenIn;
						ptInsert2 = ptInsert + vecInsertX*(iter2->first.m_dLen - dTmpInlen);
					}
				}
				break;
			}
		}
	}
	if (!blockId.isValid())
	{
		double dTmpLen = g_minLenIn;
		if (dTmpLen > dInLen)
			return FALSE;
		ptInsert = pathPtData.ptStart - vecInsertX*dTmpLen;
		std::map<BackEdgeTypeData, AcDbObjectId>::reverse_iterator iter2 = g_mapStraightBackData.rbegin();
		backEdgedata = iter2->first;
		blockId = iter2->second;
bNext = TRUE;
	}
	return TRUE;
}

AcDbObjectId CRLDrawBackEdge::getBlockIdNoStartNoEnd(BackEdgeTypeData& backEdgedata, AcGePoint3d& ptInsert, BOOL& bNext, const AcGePoint3d& ptStart, const AcGePoint3d&ptEnd)
{
	bNext = FALSE;
	AcGeVector3d vecInsertX = ptEnd - ptStart;
	vecInsertX.normalize();
	AcDbObjectId blockId = AcDbObjectId::kNull;
	double dLen = ptEnd.distanceTo(ptStart);
	if (dLen < EPRES)
		return blockId;
	std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapStraightBackData.begin();
	for (iter; iter != g_mapStraightBackData.end(); iter++)
	{
		backEdgedata = iter->first;
		double dBackLen = iter->first.m_dLen;
		if (dLen > dBackLen - 2 * g_singleMinLenout)
			continue;

		double dTmpLen = (dBackLen - dLen) / 2.0;
		ptInsert = ptStart - vecInsertX*dTmpLen;
		blockId = iter->second;
		break;
	}
	if (!blockId.isValid())
	{
		double dTmpLen = g_standlLenout;
		ptInsert = ptStart - vecInsertX*dTmpLen;
		std::map<BackEdgeTypeData, AcDbObjectId>::reverse_iterator iter2 = g_mapStraightBackData.rbegin();
		if (!((ptEnd - (ptInsert + vecInsertX * iter2->first.m_dLen)).normal()).isEqualTo(vecInsertX.normal(), frgGlobals::Gtol))
			iter2++;
		blockId = iter2->second;
		backEdgedata = iter2->first;
		bNext = TRUE;
	}
	return blockId;
}

BOOL CRLDrawBackEdge::addBlockResBlockReference(const breakInsertData& insertData)
{
	BackEdgeTypeData backData = insertData.backData;
	AcDbObjectId ret;
	AcDbBlockReference* pBlockRef = new AcDbBlockReference(AcGePoint3d::kOrigin, insertData.entBlockId);
	pBlockRef->transformBy(insertData.mat);
	pBlockRef->setLayer(g_strBackLayer);
	pBlockRef->setColorIndex(backData.m_iColorIndex);
	if (!AddToModelSpace(pBlockRef, ret))
	{
		delete pBlockRef;
		return FALSE;
	}
// 	if (g_mapIDToBlockText.find(insertData.entBlockId) == g_mapIDToBlockText.end())
// 		return TRUE;
// 	CString strBlockText = g_mapIDToBlockText[insertData.entBlockId];
// 	if (strBlockText.IsEmpty())
// 		return TRUE;
// 	AcGePoint3d origin;
// 	AcGeVector3d xAxis;
// 	AcGeVector3d yAxis;
// 	AcGeVector3d zAxis;
// 	insertData.mat.getCoordSystem(origin, xAxis, yAxis, zAxis);
// 
// 	BOOL bUpdataVec = xAxis.angleTo(AcGeVector3d::kXAxis) > PI / 2.0;
// 	AcGeVector3d xResAxis = xAxis.normal();
// 	if (bUpdataVec)
// 		xResAxis = -xAxis.normal();
// 
// 	AcGeVector3d yResAxis = xResAxis.perpVector().normal();
// 	AcGeVector3d zResAxis = xResAxis.crossProduct(yResAxis).normal();
// 	AcGeMatrix3d matText;
// 
// 	double dWith = backData.m_dWith;
// 	AcGeVector3d vecMove = -yAxis * 15;
// 	AcDbText* pText = new AcDbText();
// 	pText->setTextStyle(getTextStyleId(_T("Standard")));
// 	pText->setTextString(backData.m_strBackEdge);
// 	pText->setColorIndex(backData.m_iColorIndex);
// 	pText->setHeight(50.0);
// 	AcGePoint3d ptInsert;
// 	if (bUpdataVec)
// 	{
// 		ptInsert = origin + vecMove + xAxis*(backData.m_dLen - 50);
// 		if (backData.m_dLen > 999)
// 			ptInsert = origin + vecMove + xAxis*(backData.m_dLen / 2.0 + 2 * g_minLenIn);
// 	}
// 	else
// 	{
// 		ptInsert = origin + vecMove + xAxis*(dWith - 15);
// 		if (backData.m_dLen > 999)
// 			ptInsert = origin + vecMove + xAxis*(backData.m_dLen / 2.0 - 2 * g_minLenIn);
// 	}
// 	pText->setHorizontalMode(AcDb::kTextLeft);
// 	pText->setVerticalMode(AcDb::kTextBottom);
// 	if (!origin.isEqualTo(insertData.ptZero) && xAxis.perpVector().normal().isEqualTo(yAxis, frgGlobals::Gtol))
// 	{
// 		pText->setHorizontalMode(AcDb::kTextLeft);
// 		pText->setVerticalMode(AcDb::kTextTop);
// 	}
// 	else if (origin.isEqualTo(insertData.ptZero) && !xAxis.perpVector().normal().isEqualTo(yAxis, frgGlobals::Gtol))
// 	{
// 		pText->setHorizontalMode(AcDb::kTextLeft);
// 		pText->setVerticalMode(AcDb::kTextTop);
// 	}
// 	matText.setCoordSystem(ptInsert, xResAxis, yResAxis, zResAxis);
// 	pText->transformBy(matText);
// 	if (!AddToModelSpace(pText, ret))
// 	{
// 		delete pText;
// 		return FALSE;
// 	}
	return TRUE;
}

BOOL CRLDrawBackEdge::addBlockResBlockReference(const straigthInsertData& insetData)
{
	BackEdgeTypeData backData = insetData.backData;
	AcDbObjectId ret;
	AcDbBlockReference* pBlockRef = new AcDbBlockReference(AcGePoint3d::kOrigin, insetData.entBlockId);
	pBlockRef->transformBy(insetData.mat);
	pBlockRef->setLayer(g_strBackLayer);
	pBlockRef->setColorIndex(backData.m_iColorIndex);
	if (!AddToModelSpace(pBlockRef, ret))
	{
		delete pBlockRef;
		return FALSE;
	}
	if (g_mapIDToBlockText.find(insetData.entBlockId) == g_mapIDToBlockText.end())
		return TRUE;
	CString strBlockText = g_mapIDToBlockText[insetData.entBlockId];
	if (strBlockText.IsEmpty())
		return TRUE;
	AcGePoint3d origin;
	AcGeVector3d xAxis;
	AcGeVector3d yAxis;
	AcGeVector3d zAxis;
	insetData.mat.getCoordSystem(origin, xAxis, yAxis, zAxis);

	BOOL bUpdataVec = xAxis.angleTo(AcGeVector3d::kXAxis) > PI / 2.0;
	AcGeVector3d xResAxis = xAxis.normal();
	if (bUpdataVec)
		xResAxis = -xAxis.normal();

	AcGeVector3d yResAxis = xResAxis.perpVector().normal();
	AcGeVector3d zResAxis = xResAxis.crossProduct(yResAxis).normal();
	AcGeMatrix3d matText;

	double dWith = backData.m_dWith;
	AcGeVector3d vecMove = yAxis * 15;
	if (!backData.m_bBlue)
		vecMove = yAxis*((g_dBlueWith - g_dGreenWith) / 2.0);

	double dLen = backData.m_dLen;
	AcDbText* pText = new AcDbText();
	pText->setTextStyle(getTextStyleId(_T("Standard")));
	pText->setTextString(backData.m_strBackEdge + _T(" "));
	pText->setColorIndex(backData.m_iColorIndex);
	pText->setHeight(50.0);
	AcGePoint3d ptInsert;
	if ((bUpdataVec&&xAxis.perpVector().normal().isEqualTo(yAxis, frgGlobals::Gtol))
		|| (!xAxis.perpVector().normal().isEqualTo(yAxis, frgGlobals::Gtol) && !bUpdataVec))
	{
		pText->setHorizontalMode(AcDb::kTextLeft);
		pText->setVerticalMode(AcDb::kTextTop);
		ptInsert = origin + vecMove + xAxis*(dLen / 2 + 50);
	}
	else
	{
		pText->setHorizontalMode(AcDb::kTextLeft);
		pText->setVerticalMode(AcDb::kTextBottom);
		ptInsert = origin + vecMove + xAxis*(dLen / 2 - 50);
	}

	matText.setCoordSystem(ptInsert, xResAxis, yResAxis, zResAxis);
	pText->transformBy(matText);
	if (!AddToModelSpace(pText, ret))
	{
		delete pText;
		return FALSE;
	}
	return TRUE;
}


void CRLDrawBackEdge::getMapStraightIter(std::vector<std::map<BackEdgeTypeData, AcDbObjectId>::iterator>& vecIter)
{
	std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapStraightBackData.begin();
	for (iter; iter != g_mapStraightBackData.end();iter++)
		vecIter.push_back(iter);
}

BOOL CRLDrawBackEdge::getBlockIdConnect(AcDbObjectId& blockId, BackEdgeTypeData& backEdgedata1, AcGePoint3d& ptInsert, AcDbObjectId& blockId2, 
	BackEdgeTypeData& backEdgedata2, AcGePoint3d& ptInsert2, AcDbObjectId& blockId3, BackEdgeTypeData& backEdgedata3, AcGePoint3d& ptInsert3,
	 const AcGeVector3d& vecInesrt, const AcGePoint3d& ptStar, const AcGePoint3d&ptEnd,double dInLen1, double dInlen2)
{
	if (dInLen1 < g_minLenIn || dInlen2 < g_minLenIn)
		return FALSE;
	double dLen = ptStar.distanceTo(ptEnd);
	AcGeVector3d vecInsertX = vecInesrt;
	vecInsertX.normalize();
	ptInsert = ptStar;
	std::vector<std::map<BackEdgeTypeData, AcDbObjectId>::iterator> vecIter;
	getMapStraightIter(vecIter);
	assert(vecIter.size() == 6);
	int nCount = 0;
	while (nCount < 14)
	{
		nCount++;
		std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter1 = g_mapStraightBackData.end();
		std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter2 = g_mapStraightBackData.end();
		std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter3 = g_mapStraightBackData.end();
		switch (nCount)
		{
		case  1:
			iter1 = vecIter[0];
			iter2 = vecIter[1];
			iter3 = vecIter[0];
			break; //(500B-1000A-500B)
		case  2:
			iter1 = vecIter[0];
			iter2 = vecIter[1];
			iter3 = vecIter[2];
			break; //(500B-1000A-1000B)
		case  3:
			iter1 = vecIter[0];
			iter2 = vecIter[1];
			iter3 = vecIter[3];
			break; //(500B-1000A-1500B)
		case  4:
			iter1 = vecIter[0];
			iter2 = vecIter[5];
			iter3 = vecIter[0];
			break; //(500B-2500A-500B)
		case  5:
			iter1 = vecIter[0];
			iter2 = vecIter[1];
			iter3 = vecIter[4];
			break; //(500B-1000A-2000B)
		case  6:
			iter1 = vecIter[0];
			iter2 = vecIter[5];
			iter3 = vecIter[2];
			break; //(500B-2500A-1000B)
		case  7:
			iter1 = vecIter[2];
			iter2 = vecIter[1];
			iter3 = vecIter[4];
			break; //(1000B-1000A-2000B)
		case  8:
			iter1 = vecIter[0];
			iter2 = vecIter[5];
			iter3 = vecIter[3];
			break; //(500B-2500A-1500B)
		case  9:
			iter1 = vecIter[3];
			iter2 = vecIter[1];
			iter3 = vecIter[4];
			break; //(1500B-1000A-2000B)
		case  10:
			iter1 = vecIter[0];
			iter2 = vecIter[5];
			iter3 = vecIter[4];
			break; //(500B-2500A-2000B)
		case  11:
			iter1 = vecIter[4];
			iter2 = vecIter[1];
			iter3 = vecIter[4];
			break; //(2000B-1000A-2000B)
		case  12:
			iter1 = vecIter[2];
			iter2 = vecIter[5];
			iter3 = vecIter[4];
			break; //(1000B-2500A-2000B)
		case  13:
			iter1 = vecIter[3];
			iter2 = vecIter[5];
			iter3 = vecIter[4];
			break; //(1500B-2500A-2000B)
		case  14:
			iter1 = vecIter[4];
			iter2 = vecIter[5];
			iter3 = vecIter[4];
			break; //(1000B-2500A-2000B)
		default:
			break;
		}
		if (iter2 == g_mapStraightBackData.end() || iter1 == g_mapStraightBackData.end() || iter3 == g_mapStraightBackData.end())
			continue;
		if (iter1->first.m_bBlue || !iter2->first.m_bBlue || iter3->first.m_bBlue)
		   continue;
		if (iter2->first.m_dLen > dLen)
		   continue;
		double dTotalLen = iter1->first.m_dLen + iter2->first.m_dLen + iter3->first.m_dLen;
		if (dLen - dTotalLen +4*g_minLenIn < EPRES/*dLen < dTotalLen - 4*g_minLenIn*/)
		{
			double dTmpLen = dTotalLen - dLen;
			blockId = iter1->second;
			blockId2 = iter2->second;
			blockId3 = iter3->second;
			backEdgedata1 = iter1->first;
			backEdgedata2 = iter2->first;
			backEdgedata3 = iter3->first;
			double dInLenTmp1 = dTmpLen / 4.0;
			double dInLenTmp2 = dInLenTmp1;
			double dInLenTmp3 = dInLenTmp1;
			if (!getInlenBetweenThreeBackEdge(dInLenTmp1, dInLenTmp2, dInLenTmp3, dTmpLen, dInLen1, iter1->first.m_dLen - g_sideMinLen, iter2->first.m_dLen - g_sideMinLen, dInlen2))
				return FALSE;
			ptInsert = ptInsert - vecInsertX*dInLenTmp1;
			ptInsert2 = ptInsert + vecInsertX*iter1->first.m_dLen - vecInsertX*dInLenTmp2;
			ptInsert3 = ptInsert2 + vecInsertX*iter2->first.m_dLen - vecInsertX*dInLenTmp3;
			break;
		}
	}
	return TRUE;
}

BOOL CRLDrawBackEdge::getInlenBetweenThreeBackEdge(double& dInResLen1, double& dInReslen2, double& dInReslen3, double dTotalLen,
	                         double dInLen1, double dInlen2, double dInLen3, double dInLen4)
{
	if (dInLen1 < g_minLenIn || dInlen2 < g_minLenIn || dInLen3 < g_minLenIn || dInLen4 < g_minLenIn)
		return FALSE;

	dInResLen1 = dTotalLen / 4.0;
	if (dInResLen1 - dInLen1 < EPRES &&dInResLen1 - dInlen2 < EPRES&&dInResLen1 - dInLen3 < EPRES
		&&dInResLen1 - dInLen4 < EPRES)
	{
		dInReslen2 = dInReslen3 = dInResLen1;
		return TRUE;
	}
	else if (dInResLen1 > dInLen1)
	{
		dInResLen1 = dInLen1;
		if (!getInlenBetweenTwoBackEdge(dInReslen2, dInReslen3, dTotalLen - dInResLen1, dInlen2, dInLen3, dInLen4))
			return FALSE;
	}
	else if (dInResLen1 > dInlen2)
	{
		dInReslen2 = dInlen2;
		if (!getInlenBetweenTwoBackEdge(dInResLen1, dInReslen3, dTotalLen - dInReslen2, dInLen1, dInLen3, dInLen4))
			return FALSE;
	}
	else if (dInResLen1 > dInLen3)
	{
		dInReslen3 = dInlen2;
		if (!getInlenBetweenTwoBackEdge(dInResLen1, dInReslen2, dTotalLen - dInReslen3, dInLen1, dInlen2, dInLen4))
			return FALSE;
	}
	else if (dInResLen1 > dInLen4)
	{
		double dTmpLen = dInLen4;
		dInResLen1 = (dTotalLen - dTmpLen)/3.0;
		if (dInResLen1 - dInLen1 < EPRES &&dInResLen1 - dInlen2 < EPRES&&dInResLen1 - dInLen3 < EPRES)
		{
			dInReslen3 = dInReslen2 = dInResLen1;
			return TRUE;
		}
		else if (dInResLen1 > dInLen1)
		{
			dInResLen1 = dInLen1;
			dInReslen2 = (dTotalLen - dInResLen1) / 2.0;
			if (dInReslen2 > dInlen2)
				dInReslen2 = dInlen2;
			dInReslen3 = dTotalLen - dInResLen1 - dInReslen2;
			if (dInReslen3 > dInLen3)
				dInReslen3 = dInLen3;
			dInReslen2 = dTotalLen - dInReslen3 - dInResLen1;
			if (dInReslen2 > dInlen2)
				return FALSE;
		}
		else if (dInResLen1 > dInlen2)
		{
			dInReslen2 = dInlen2;
			dInResLen1 = (dTotalLen - dInReslen2) / 2.0;
			if (dInResLen1 > dInLen1)
				dInResLen1 = dInLen1;
			dInReslen3 = dTotalLen - dInResLen1 - dInReslen2;
			if (dInReslen3 > dInLen3)
				dInReslen3 = dInLen3;
			dInResLen1 = dTotalLen - dInReslen3 - dInReslen2;
			if (dInResLen1 > dInLen1)
				return FALSE;
		}
		else if (dInResLen1 > dInLen3)
		{
			dInReslen3 = dInLen3;
			dInResLen1 = (dTotalLen - dInReslen3) / 2.0;
			if (dInResLen1 > dInLen1)
				dInResLen1 = dInLen1;
			dInReslen2 = dTotalLen - dInReslen3 - dInResLen1;
			if (dInReslen2 > dInlen2)
				dInReslen2 = dInlen2;
			dInResLen1 = dTotalLen - dInReslen3 - dInReslen2;
			if (dInResLen1 > dInLen1)
				return FALSE;
		}
	}
	return TRUE;
}

BOOL CRLDrawBackEdge::getBlockIdConnect(AcDbObjectId& blockId, BackEdgeTypeData& backEdgedata1, AcGePoint3d& ptInsert, AcDbObjectId& blockId2, 
	BackEdgeTypeData& backEdgedata2, AcGePoint3d& ptInsert2,const AcGeVector3d& vecInesrt, 
	const AcGePoint3d& ptStar, const AcGePoint3d&ptEnd, double dInLen1, double dInlen2)
{
	if (dInLen1 < g_minLenIn || dInlen2 < g_minLenIn )
		return FALSE;

	double dLen = ptStar.distanceTo(ptEnd);
	AcGeVector3d vecInsertX = vecInesrt;
	vecInsertX.normalize();
	ptInsert = ptStar;
	std::vector<std::map<BackEdgeTypeData, AcDbObjectId>::iterator> vecIter;
	getMapStraightIter(vecIter);
	assert(vecIter.size() == 6);

	int nCount = 0;
	while (nCount < 8)
	{
		nCount++;
		std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter1 = g_mapStraightBackData.end();
		std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter2 = g_mapStraightBackData.end();
		switch (nCount)
		{
		case  1:
			iter1 = vecIter[1];
			iter2 = vecIter[0];
			break; //(500B-1000A)
		case  2:
			iter1 = vecIter[1];
			iter2 = vecIter[2]; //(1000B-1000A)
			break;
		case  3:
			iter1 = vecIter[1];
			iter2 = vecIter[3]; //(1500B-1000A)
			break;
		case  4:
			iter1 = vecIter[5];
			iter2 = vecIter[0]; //(500B-2500A)
			break;
		case  5:
			iter1 = vecIter[1];
			iter2 = vecIter[4]; //(2000B-1000A)
			break;
		case  6:
			iter1 = vecIter[5];
			iter2 = vecIter[2];  //(1000B-2500A)
			break;
		case  7:
			iter1 = vecIter[5];
			iter2 = vecIter[3]; //(1500B-2500A)
			break;
		case  8:
			iter1 = vecIter[5];
			iter2 = vecIter[4]; //(2000B-2500A)
			break;
		default:
			break;
		}
		if (iter2 == g_mapStraightBackData.end() || iter1 == g_mapStraightBackData.end()
			|| iter2->first.m_bBlue || !iter1->first.m_bBlue)
			continue;
		double dTotalLen = iter2->first.m_dLen + iter1->first.m_dLen;
		if (dLen - dTotalLen + 3 * g_minLenIn< EPRES/*dLen < dTotalLen - 2 * g_minLenIn - g_groupMinLenout*/)
		{
			double dTmpLen = dTotalLen - dLen;
			blockId = iter1->second;
			blockId2 = iter2->second;
			backEdgedata1 = iter1->first;
			backEdgedata2 = iter2->first;
			double dInLenTmp1 = dTmpLen / 3.0;
			double dInLenTmp2 = dTmpLen / 3.0;
			if (!getInlenBetweenTwoBackEdge(dInLenTmp1, dInLenTmp2, dTmpLen, dInLen1, iter1->first.m_dLen2 - g_minLenIn, dInlen2))
				return FALSE;
			ptInsert = ptInsert - vecInsertX*dInLenTmp1;
			ptInsert2 = ptInsert + vecInsertX*iter1->first.m_dLen - vecInsertX*dInLenTmp2;
			break;
		}
	}
	return TRUE;
}

BOOL CRLDrawBackEdge::getInlenBetweenTwoBackEdge(double& dInResLen1, double& dInReslen2, double dTotalLen, double dInLen1, double dInlen2, double dInLen3)
{
	if (dInLen1 < g_minLenIn || dInlen2 < g_minLenIn || dInLen3 < g_minLenIn)
		return FALSE;

	dInResLen1 = dTotalLen / 3.0;
	if (dInResLen1 - dInLen1 < EPRES &&dInResLen1 - dInlen2 <EPRES&&dInResLen1 - dInLen3 < EPRES)
	{
		dInReslen2 = dInResLen1;
		return TRUE;
	}
	else if (dInResLen1 > dInLen1)
	{
		dInResLen1 = dInLen1;
		dInReslen2 = (dTotalLen - dInResLen1) / 2.0;
		if (dInReslen2 > dInlen2)
			dInReslen2 = dInlen2;
		double dInLenTmp3 = dTotalLen - dInResLen1 - dInReslen2;
		if (dInLenTmp3 > dInLen3)
			dInLenTmp3 = dInLen3;
		dInReslen2 = dTotalLen - dInLenTmp3 - dInResLen1;
		if (dInReslen2 > dInlen2)
			return FALSE;
	}
	else if (dInResLen1 > dInlen2)
	{
		dInReslen2 = dInlen2;
		dInResLen1 = (dTotalLen - dInReslen2) / 2.0;
		if (dInResLen1 > dInLen1)
			dInResLen1 = dInLen1;
		double dInLenTmp3 = dTotalLen - dInResLen1 - dInReslen2;
		if (dInLenTmp3 > dInLen3)
			dInLenTmp3 = dInLen3;
		dInResLen1 = dTotalLen - dInLenTmp3 - dInReslen2;
		if (dInResLen1 > dInLen1)
			return FALSE;
	}
	else if (dInResLen1 > dInLen3)
	{
		double dInLenTmp3 = dInLen3;
		dInResLen1 = (dTotalLen - dInLenTmp3) / 2.0;
		if (dInResLen1 > dInLen1)
			dInResLen1 = dInLen1;
		dInReslen2 = dTotalLen - dInLenTmp3 - dInResLen1;
		if (dInReslen2 > dInlen2)
			dInReslen2 = dInlen2;
		dInResLen1 = dTotalLen - dInLenTmp3 - dInReslen2;
		if (dInResLen1 > dInLen1)
			return FALSE;
	}
	return TRUE;
}

BOOL CRLDrawBackEdge::connectBackEdge(std::vector<straigthInsertData>& vecStraigthData, const AcGeVector3d& vecInesrt, const AcGePoint3d& ptStar, const AcGePoint3d&ptEnd,
	 double dInLen1, double dInlen2, BOOL bBlue)
{
	if (dInLen1 < g_minLenIn || dInlen2 < g_minLenIn)
		return FALSE;
	//两端全是BLUE端
	BackEdgeTypeData backEdgedata;
	double dLen = ptStar.distanceTo(ptEnd);
	AcGeVector3d vecInsertX = vecInesrt;
	vecInsertX.normalize();
	AcGeVector3d vecInsertY = vecInsertX;
	vecInsertY.rotateBy(PI / 2.0, AcGeVector3d(0, 0, 1));
	vecInsertY.normalize();
	double bNext = FALSE;
	AcGePoint3d ptInsert = ptStar;
	if (bBlue)
	{
		AcDbObjectId blockId = AcDbObjectId::kNull;
		AcDbObjectId blockId2;
		AcGePoint3d ptInsert2;
		AcDbObjectId blockId3;
		AcGePoint3d ptInsert3;
		BackEdgeTypeData backEdgedata2;
		BackEdgeTypeData backEdgedata3;
		std::map<BackEdgeTypeData, AcDbObjectId>::iterator iter = g_mapStraightBackData.begin();
		for (iter; iter != g_mapStraightBackData.end(); iter++)
		{
			backEdgedata = iter->first;
			double dBackLen = iter->first.m_dLen;
			if (backEdgedata.m_bBlue)
				continue;
			if (dLen-dBackLen +2*g_minLenIn < EPRES/*dBackLen - 2 * g_minLenIn > dLen*/)
			{
				double dTmpLen = dBackLen - dLen;
				blockId = iter->second;
				dTmpLen /= 2;
				if (dTmpLen > dInLen1)
					dTmpLen = dInLen1;
				double dElsInlen = dBackLen - dLen - dTmpLen;
				if (dElsInlen > dInlen2)
					dElsInlen = dInlen2;
				dTmpLen = dBackLen - dLen - dElsInlen;
				if (dTmpLen > dInLen1)
					return FALSE;
				ptInsert = ptStar - vecInsertX*dTmpLen;
				break;
			}
		}
		if (!blockId.isValid())
		{
			getBlockIdConnect(blockId, backEdgedata,ptInsert, blockId2, backEdgedata2, ptInsert2, blockId3, backEdgedata3, ptInsert3, vecInesrt, ptStar, ptEnd, dInLen1, dInlen2);
		}
		if (!blockId.isValid())
		{
			std::map<BackEdgeTypeData, AcDbObjectId>::reverse_iterator iter2 = g_mapStraightBackData.rbegin();
			ptInsert = ptStar - vecInsertX*g_minLenIn;
			iter2++;
			backEdgedata = iter2->first;
			blockId = iter2->second;
			bNext = TRUE;
		}
		if (!blockId.isValid())
			return FALSE;
		ptInsert = ptInsert + vecInsertY*((g_dBlueWith - g_dGreenWith) / 2.0);
		AcGeMatrix3d mat;
		mat.setCoordSystem(ptInsert, vecInsertX.normal(), vecInsertY.normal(), vecInsertX.crossProduct(vecInsertY).normal());
		//if (!addBlockResBlockReference(blockId, 3, mat))
			//return FALSE;
		straigthInsertData insertData;
		insertData.backData = backEdgedata;
		insertData.mat = mat;
		insertData.entBlockId = blockId;
		vecStraigthData.push_back(insertData);

		ptInsert = ptInsert - vecInsertY*((g_dBlueWith - g_dGreenWith) / 2.0);
		if (blockId2.isValid() && blockId3.isValid())
		{
			mat.setCoordSystem(ptInsert2, vecInsertX.normal(), vecInsertY.normal(), vecInsertX.crossProduct(vecInsertY).normal());
// 			if (!addBlockResBlockReference(blockId2, 3, mat))
// 				return FALSE;
			insertData.backData = backEdgedata2;
			insertData.mat = mat;
			insertData.entBlockId = blockId2;
			vecStraigthData.push_back(insertData);

			ptInsert3 = ptInsert3 + vecInsertY*((g_dBlueWith - g_dGreenWith) / 2.0);
			mat.setCoordSystem(ptInsert3, vecInsertX.normal(), vecInsertY.normal(), vecInsertX.crossProduct(vecInsertY).normal());
			//if (!addBlockResBlockReference(blockId3, 3, mat))
			//	return FALSE;
			insertData.backData = backEdgedata3;
			insertData.mat = mat;
			insertData.entBlockId = blockId3;
			vecStraigthData.push_back(insertData);
		}
	}
	else
	{
		AcDbObjectId blockId = AcDbObjectId::kNull;
		AcDbObjectId blockId2;
		AcGePoint3d ptInsert2;
		BackEdgeTypeData backEdgedata2;
		getBlockIdConnect(blockId, backEdgedata,ptInsert, blockId2, backEdgedata2, ptInsert2, vecInesrt, ptStar, ptEnd, dInLen1, dInlen2);
		if (!blockId.isValid())
		{
			std::map<BackEdgeTypeData, AcDbObjectId>::reverse_iterator iter2 = g_mapStraightBackData.rbegin();
			ptInsert = ptStar - vecInsertX*g_minLenIn;
			backEdgedata = iter2->first;
			blockId = iter2->second;
			bNext = TRUE;
		}
		if (!blockId.isValid())
			return FALSE;
		AcGeMatrix3d mat;
		mat.setCoordSystem(ptInsert, vecInsertX.normal(), vecInsertY.normal(), vecInsertX.crossProduct(vecInsertY).normal());
		//if (!addBlockResBlockReference(blockId, 3, mat))
		//	return FALSE;
		straigthInsertData insertData;
		insertData.backData = backEdgedata;
		insertData.mat = mat;
		insertData.entBlockId = blockId;
		vecStraigthData.push_back(insertData);

		if (blockId2.isValid())
		{
			ptInsert2 = ptInsert2 + vecInsertY*((g_dBlueWith - g_dGreenWith) / 2.0);
			mat.setCoordSystem(ptInsert2, vecInsertX.normal(), vecInsertY.normal(), vecInsertX.crossProduct(vecInsertY).normal());
			straigthInsertData insertData;
			insertData.backData = backEdgedata2;
			insertData.mat = mat;
			insertData.entBlockId = blockId2;
			vecStraigthData.push_back(insertData);
			//if (!addBlockResBlockReference(blockId2, 3, mat))
			//	return FALSE;
		}
	}
	if (bNext)  
	{
		bBlue = !bBlue;
		double dTmpIn = backEdgedata.m_dLen - g_sideMinLen;
		AcGePoint3d ptTmp = ptInsert + vecInsertX*backEdgedata.m_dLen;
		if (!connectBackEdge(vecStraigthData,vecInesrt, ptTmp, ptEnd, dTmpIn, dInlen2, bBlue))
			return FALSE;
	}
	return TRUE;
}

AcArray<AcGePoint3dArray> CRLDrawBackEdge::delBackLines(const AcDbObjectIdArray&entCountIds)
{
	AcArray<AcGePoint3dArray> pathDelPtArrys;
	AcDbObjectIdArray entIdArry;
	struct resbuf* rb = acutBuildList(-4, _T("<AND"), 8, _T("fmk_lines"), RTDXF0, _T("LINE"), -4, _T("AND>"), RTNONE);
	ads_name ssText;
	acutPrintf(_T("\n请选择去掉的路径线:"));
	int iRet = acedSSGet(NULL, NULL, NULL, rb, ssText);
	if (RTNORM != iRet)
	{
		acutRelRb(rb);
		acedSSFree(ssText);
		return pathDelPtArrys;
	}
	long ssLen;
	acedSSLength(ssText, &ssLen);
	for (int i = 0; i < ssLen; i++)
	{
		ads_name ent;
		acedSSName(ssText, i, ent);
		AcDbObjectId eld;
		acdbGetObjectId(eld, ent);
		acedSSFree(ent);
		if (!entCountIds.contains(eld))
		    continue;
		AcDbObjectPointer<AcDbLine> pLine(eld, AcDb::kForWrite);
		if (pLine.openStatus() != Acad::eOk)
			continue;
		AcGePoint3dArray ptArry;
		ptArry.append(pLine->startPoint());
		ptArry.append(pLine->endPoint());
		pathDelPtArrys.append(ptArry);
		ptArry.removeAll();
		ptArry.append(pLine->endPoint());
		ptArry.append(pLine->startPoint());
		pathDelPtArrys.append(ptArry);
		pLine->erase();
	}
	acutRelRb(rb);
	acedSSFree(ssText);
	actrTransactionManager->flushGraphics();
	acedUpdateDisplay();
	return pathDelPtArrys;
}

bool CRLDrawBackEdge::selectBackEdges(AcDbObjectIdArray& entIdArry)
{
	struct resbuf* rb = acutBuildList(-4, _T("<OR"), RTDXF0, _T("INSERT"), RTDXF0, _T("TEXT"), RTDXF0, _T("MTEXT"), -4, _T("OR>"), RTNONE);
	/*struct resbuf* rb = acutBuildList(0, "INSERT", AcDb::kDxfLayerName, "背楞",RTNONE);*/
	ads_name ssText;
	acutPrintf(_T("\n请选择背楞:"));
	int iRet = acedSSGet(NULL, NULL, NULL, rb, ssText);
	if (RTNORM != iRet)
	{
		acutRelRb(rb);
		acedSSFree(ssText);
		return false;
	}
	long ssLen;
	acedSSLength(ssText, &ssLen);
	for (int i = 0; i < ssLen; i++)
	{
		ads_name ent;
		acedSSName(ssText, i, ent);
		AcDbObjectId eld;
		acdbGetObjectId(eld, ent);
		if (eld.isValid() && !entIdArry.contains(eld))
			entIdArry.append(eld);
		acedSSFree(ent);
	}
	acutRelRb(rb);
	acedSSFree(ssText);
	return !entIdArry.isEmpty();
}

bool CRLDrawBackEdge::selectAutoLines(AcDbObjectIdArray& entIdArry, AcDbObjectIdArray& entDwgIds)
{
	struct resbuf* rb = acutBuildList(-4, _T("<OR"), RTDXF0, _T("INSERT"), RTDXF0, _T("REGION"), RTDXF0, _T("LWPOLYLINE"), RTDXF0, _T("LINE"), -4, _T("OR>"), RTNONE);
	ads_name ssText; 
	acutPrintf(_T("\n请选择背楞路径:"));
	int iRet = acedSSGet(NULL, NULL, NULL, rb, ssText);
	if (RTNORM != iRet)
	{
		acedSSFree(ssText);
		acutRelRb(rb);
		return false;
	}
	long ssLen;
	acedSSLength(ssText, &ssLen);
	for (int i = 0; i < ssLen; i++)
	{
		ads_name ent;
		acedSSName(ssText, i, ent);
		AcDbObjectId eld;
		acdbGetObjectId(eld, ent);
		acedSSFree(ent);
		AcDbObjectPointer<AcDbEntity> pEnt(eld, AcDb::kForWrite);
		if (pEnt.openStatus() != Acad::eOk)
			continue;
		entDwgIds.append(eld);
		if (pEnt->isKindOf(AcDbLine::desc()))
		{
			AcRxObject* pObj = pEnt->clone();
			if (pObj == NULL)
				continue;
			AcDbLine* pCloneLineEnt = AcDbLine::cast(pObj);
			if (pCloneLineEnt == NULL)
			{
				delete pObj;
				pObj = NULL;
				continue;
			}
			pCloneLineEnt->setLayer(_T("fmk_lines"));
			pCloneLineEnt->setColorIndex(getEntiColorIndex(pEnt));
			AcDbObjectId ret;
			if (!AddToModelSpace(pCloneLineEnt, ret))
			{
				delete pCloneLineEnt;
				pCloneLineEnt = NULL;
			}
			entIdArry.append(ret);
		}
		else
		{
			selectAutoLines(entIdArry, pEnt);
		}
	}
	acutRelRb(rb);
	acedSSFree(ssText);
	actrTransactionManager->flushGraphics();
	acedUpdateDisplay();
	return !entIdArry.isEmpty();
}

void CRLDrawBackEdge::selectAutoLines(AcDbObjectIdArray& entIdArry, AcDbEntity* pEnt)
{
	if (pEnt == NULL)
		return;

	if (!pEnt->isKindOf(AcDbRegion::desc()) && !pEnt->isKindOf(AcDbPolyline::desc())
		&& !pEnt->isKindOf(AcDbBlockReference::desc()))
		return;

	AcGeVoidPointerArray ptIines;
	if (pEnt->explode(ptIines) != Acad::eOk)
		releaseVoidPtr(ptIines);
	int nSize = ptIines.length();
	for (int j = 0; j < nSize; j++)
	{
		AcDbEntity* pExEnt = static_cast<AcDbEntity*>(ptIines[j]);
		if (pExEnt != NULL)
		{
			if (pExEnt->isKindOf(AcDbLine::desc()))
			{
				pExEnt->setPropertiesFrom(pEnt);
				pExEnt->setLayer(_T("fmk_lines"));
				pExEnt->setColorIndex(getEntiColorIndex(pEnt));
				AcDbObjectId ret;
				if (!AddToModelSpace(pExEnt, ret))
				{
					delete pExEnt;
					pExEnt = NULL;
				}
				entIdArry.append(ret);
			}
			else
			{
				selectAutoLines(entIdArry, pExEnt);
				delete pExEnt;
				pExEnt = NULL;
			}
		}
		else
		{
			AcDbObject *pObj = static_cast<AcDbObject*>(ptIines[j]);
			if (pObj)
			{
				delete pObj;
				pObj = NULL;
			}
		}
	}
}

bool CRLDrawBackEdge::selectPLines(AcDbObjectIdArray& entIdArry)
{
	struct resbuf rb;
	rb.restype = 0; //实体名
	CString strEnt = _T("LWPOLYLINE");
	rb.resval.rstring = strEnt.GetBuffer();
	strEnt.ReleaseBuffer();
	rb.rbnext = NULL; //无其他内容
	ads_name ssText;
	acutPrintf(_T("\n请选择背楞多段线路径:"));
	int iRet = acedSSGet(NULL, NULL, NULL, &rb, ssText);
	if (RTNORM != iRet)
	{
		acedSSFree(ssText);
		return false;
	}
	long ssLen;
	acedSSLength(ssText, &ssLen);
	for (int i = 0; i < ssLen; i++)
	{
		ads_name ent;
		acedSSName(ssText, i, ent);
		AcDbObjectId eld;
		acdbGetObjectId(eld, ent);
		if (eld.isValid() && !entIdArry.contains(eld))
			entIdArry.append(eld);
		acedSSFree(ent);
	}
	acedSSFree(ssText);
	return !entIdArry.isEmpty();
}

void  CRLDrawBackEdge::earseEnts(const AcDbObjectIdArray& entIdArry, BOOL bEarse)
{
	int nLen = entIdArry.length();
	for (int i = 0; i < nLen; i++)
	{
		AcDbEntityPointer pEnt(entIdArry[i], AcDb::kForWrite);
		if (pEnt.openStatus() != Acad::eOk)
		   continue;
		pEnt->erase(bEarse);
	}
	actrTransactionManager->flushGraphics();
	acedUpdateDisplay();
}

void CRLDrawBackEdge::ShowEnt(const AcDbObjectIdArray& arrIds, BOOL bIsShow)
{
	int nSize = arrIds.length();
	for (int i = 0; i < nSize; ++i)
	{
		AcDbEntity* pEnt = NULL;
		if (Acad::eOk == acdbOpenAcDbObject((AcDbObject*&)pEnt, arrIds[i], AcDb::kForWrite))
		{
			AcDb::Visibility entViS = AcDb::kInvisible;
			if (bIsShow)
				entViS = AcDb::kVisible;
			pEnt->setVisibility(entViS);
			pEnt->close();
		}
	}
	actrTransactionManager->flushGraphics();
	acedUpdateDisplay();
}

AcDbObjectId CRLDrawBackEdge::getTextStyleId(const ACHAR* name)
{
	AcDbObjectId id = AcDbObjectId::kNull;
	AcDbDatabase* pDb = acdbCurDwg();
	if (pDb == NULL)
		return id;
	AcDbTextStyleTable* pTb = NULL;
	if (pDb->getTextStyleTable(pTb, AcDb::kForRead) != Acad::eOk)
		return id;
	ASSERT(pTb != NULL);
	if (pTb->has(name))
		pTb->getAt(name, id);
	pTb->close();
	return id;
}

bool CRLDrawBackEdge::AddToModelSpace(AcDbEntity* pEnt, AcDbObjectId& objId, bool bClose, AcDbDatabase* pDb)
{
	bool bl = true;
	if (pEnt == NULL)
		return false;
	if (pDb == NULL) pDb = acdbHostApplicationServices()->workingDatabase();
	AcApDocument* pDoc = acDocManager->document(pDb);
	AcDbBlockTable * pBlkTable;
	if (pDb->getBlockTable(pBlkTable, AcDb::kForRead) != Acad::eOk)
		return false;

	AcDbBlockTableRecord * pRec;
	if (pBlkTable->getAt(ACDB_MODEL_SPACE, pRec, AcDb::kForWrite) != Acad::eOk)
	{
		pBlkTable->close();
		return false;
	}
	pBlkTable->close();

	if (pRec->appendAcDbEntity(objId, pEnt) != Acad::eOk)
	{
		pRec->close();
		return false;
	}
	if (bClose)
	{
		pEnt->close();
	}
	pEnt->draw();
	pRec->close();
	return bl;
}

AcDbObjectId CRLDrawBackEdge::createLayer(CString strLayerName, int layercolorIndex, bool bOffLayer, AcDbDatabase *pDb)
{
	if (pDb == NULL)
		pDb = acdbCurDwg();
	AcDbLayerTable *pLayerTbl;
	AcDbObjectId entLayerId = AcDbObjectId::kNull;
	Acad::ErrorStatus es = Acad::eOk;
	es = pDb->getLayerTable(pLayerTbl, AcDb::kForWrite);
	if (es != Acad::eOk)
		return entLayerId;
	AcDbLayerTableRecord *pLayerTblRcd;
	if (!pLayerTbl->has(strLayerName))
	{
		pLayerTblRcd = new AcDbLayerTableRecord();
		pLayerTblRcd->setName(strLayerName);
		AcCmColor color;
		color.setColorIndex(layercolorIndex);// set color
		pLayerTblRcd->setColor(color);
		pLayerTblRcd->setIsOff(bOffLayer);
		pLayerTbl->add(entLayerId, pLayerTblRcd);
		pLayerTblRcd->close();
	}
	pLayerTbl->getAt(strLayerName, (AcDbLayerTableRecord*&)pLayerTblRcd, AcDb::kForRead);
	entLayerId = pLayerTblRcd->id();
	pLayerTblRcd->close();
	pLayerTbl->close();
	return entLayerId;
}

int CRLDrawBackEdge::getEntiColorIndex(AcDbEntity* pEnt)
{
	if (pEnt == NULL)
		return 256;
	int iColorIndex = pEnt->colorIndex();
	if (iColorIndex == 256)
	{
		iColorIndex = getColorIndexByLayer(pEnt->layer());
	}
	return iColorIndex;
}

int CRLDrawBackEdge::getColorIndexByLayer(CString lyname, AcDbDatabase *pDb)
{
	int iColor = 7;
	if (pDb == NULL)
		pDb = acdbCurDwg();
	Acad::ErrorStatus es = Acad::eOk;
	AcDbLayerTable *pDbLy;
	es = pDb->getLayerTable(pDbLy, AcDb::kForRead);
	if (es != Acad::eOk)
		return iColor;
	Adesk::Boolean bHas = pDbLy->has(lyname);
	if (bHas)
	{
		AcDbLayerTableRecord *pLayerTblRcd;
		es = pDbLy->getAt(lyname, (AcDbLayerTableRecord*&)pLayerTblRcd, AcDb::kForRead);
		if (es == Acad::eOk)
		{
			AcCmColor color = pLayerTblRcd->color();
			iColor = color.colorIndex();
			pLayerTblRcd->close();
		}
	}
	pDbLy->close();
	return iColor;
}

AcGeVector3d CRLDrawBackEdge::getVec3D(const AcGeVector2d& vec)
{
	return AcGeVector3d(vec.x, vec.y, 0);
}

AcGeVector2d CRLDrawBackEdge::getVec2D(const AcGeVector3d& vec)
{
	return AcGeVector2d(vec.x, vec.y);
}

AcGePoint3d CRLDrawBackEdge::getPoint3D(const AcGePoint2d& vec)
{
	return AcGePoint3d(vec.x, vec.y, 0);
}

AcGePoint2d CRLDrawBackEdge::getPoint2D(const AcGePoint3d& vec)
{
	return AcGePoint2d(vec.x, vec.y);
}
