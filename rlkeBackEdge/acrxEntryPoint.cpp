// (C) Copyright 2002-2007 by Autodesk, Inc. 
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted, 
// provided that the above copyright notice appears in all copies and 
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting 
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC. 
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to 
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//

//-----------------------------------------------------------------------------
//----- acrxEntryPoint.cpp
//-----------------------------------------------------------------------------
#include "StdAfx.h"
#include "resource.h"
#include "RLDrawBackEdge.h"
#include "rlDocReactor.h"
#include "AcadRibPanelWrapper.h"

//-----------------------------------------------------------------------------
#define szRDS _RXST("rlkj")

//-----------------------------------------------------------------------------
//----- ObjectARX EntryPoint
class CrlkeBackEdgeApp : public AcRxArxApp {

private:
	CRLDocReactor* m_pDocReactor;
	CString m_strWrokSpace;
	CString m_strCuixPath;

public:
	CrlkeBackEdgeApp() : AcRxArxApp() {}

	virtual AcRx::AppRetCode On_kInitAppMsg (void *pkt) {
		// TODO: Load dependencies here

		// You *must* call On_kInitAppMsg here
		AcRx::AppRetCode retCode =AcRxArxApp::On_kInitAppMsg (pkt) ;
		
		// TODO: Add your initialization code here
		if (!CRLDrawBackEdge::InitBackEdge())
		{
			acutPrintf(_T("\n≥ı ºªØ±≥¿„ƒ£øÈ ß∞‹£°"));
		}
		frgGlobals::Gtol.setEqualPoint(rlTolerance::equal_point());
		frgGlobals::Gtol.setEqualVector(0.001);
		m_pDocReactor = new CRLDocReactor(true);
		resbuf rb;
		if (acedGetVar(_T("WSCURRENT"), &rb) == RTNORM) 
		{
			ASSERT(rb.restype == RTSTR);
			m_strWrokSpace = rb.resval.rstring;
			free(rb.resval.rstring);
		}
		if (acedGetVar(_T("MENUNAME"), &rb) == RTNORM) 
		{
			ASSERT(rb.restype == RTSTR);
			m_strCuixPath = rb.resval.rstring;
			free(rb.resval.rstring);
			m_strCuixPath += _T(".CUIX");
		}
		CString strIciux = _T("");
		CString strModulePath = _T("");
		CString strArxFullPath;
		TCHAR lpszFilePath[_MAX_PATH];
		if (::GetModuleFileName(_hdllInstance, lpszFilePath, _MAX_PATH) == 0)
			return (retCode);
		strArxFullPath = lpszFilePath;
		int nLength = strArxFullPath.ReverseFind('\\');
		strModulePath = strArxFullPath.Left(nLength);
		strIciux = strModulePath+_T("\\acad.CUIX");
		if (m_strCuixPath.CompareNoCase(strIciux) != 0 || m_strWrokSpace.CompareNoCase(_T("AutoCAD æ≠µ‰")) != 0)
		{
			CAcadRibPanelWrapper g_acadPanelWrap;
			if (!g_acadPanelWrap.RremoveAllMenus())
			{
				return (retCode);
			}
			if (!g_acadPanelWrap.LoadMenuFromCuix(strIciux, _T("AutoCAD æ≠µ‰")))
			{
				return (retCode);
			}
		}
		return (retCode) ;
	}

	virtual AcRx::AppRetCode On_kUnloadAppMsg (void *pkt) {
		// TODO: Add your code here
		AcRx::AppRetCode retCode = AcRxArxApp::On_kUnloadAppMsg(pkt);
		// You *must* call On_kUnloadAppMsg here
		// TODO: Unload dependencies here
		if (m_pDocReactor != NULL)
		{
			delete m_pDocReactor;
			m_pDocReactor = NULL;
		}
		unLoadArx();
		return (retCode) ;  
	}
         
	void unLoadArx()
	{
		CAcadRibPanelWrapper g_acadPanelWrap;
		if (!g_acadPanelWrap.RremoveAllMenus())
		{
			return;
		}
		if (!g_acadPanelWrap.LoadMenuFromCuix(m_strCuixPath, m_strWrokSpace))
		{
			return;
		}
	}

	virtual void RegisterServerComponents () {
	}


	static void rlkjBackEdge_BackEdge(void)
	{
		CRLDrawBackEdge drawBack;
		drawBack.drawBackEdge();    
	}

	static void rlkjBackEdge_BackEdgeAuto(void)
	{
		CRLDrawBackEdge drawBack;
		drawBack.AutodrawBackEdge();
	}

	static void rlkjBackEdge_BackEdgeAutoWall(void)
	{
		CRLDrawBackEdge drawBack;
		drawBack.AutodrawBackEdgeWall();
	}

	static void rlkjBackEdge_CreatExcel(void)
	{
		CRLDrawBackEdge drawBack;
		drawBack.CreatExcel();
	}

	static void rlkjBackEdge_rllock(void)
	{
		CRLDrawBackEdge drawBack;
		drawBack.rllock();
	}
} ;

//-----------------------------------------------------------------------------
IMPLEMENT_ARX_ENTRYPOINT(CrlkeBackEdgeApp)

ACED_ARXCOMMAND_ENTRY_AUTO(CrlkeBackEdgeApp, rlkjBackEdge, _BackEdge, BackEdge, ACRX_CMD_MODAL, NULL)

ACED_ARXCOMMAND_ENTRY_AUTO(CrlkeBackEdgeApp, rlkjBackEdge, _BackEdgeAuto, BackEdgeAuto, ACRX_CMD_MODAL, NULL)

ACED_ARXCOMMAND_ENTRY_AUTO(CrlkeBackEdgeApp, rlkjBackEdge, _CreatExcel, CreatExcel, ACRX_CMD_MODAL, NULL)

ACED_ARXCOMMAND_ENTRY_AUTO(CrlkeBackEdgeApp, rlkjBackEdge, _BackEdgeAutoWall, BackEdgeAutoWall, ACRX_CMD_MODAL, NULL)

ACED_ARXCOMMAND_ENTRY_AUTO(CrlkeBackEdgeApp, rlkjBackEdge, _rllock, _rllock, ACRX_CMD_MODAL | ACRX_CMD_USEPICKSET, NULL)