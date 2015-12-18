
#include "StdAfx.h"

#include "AcadRibPanelWrapper.h"

#include "axpnt3d.h"
#include "axlock.h"

bool CAcadRibPanelWrapper::m_bIsMenuLoaded = false;

CAcadRibPanelWrapper::CAcadRibPanelWrapper()
{
	m_pAcApplication = NULL;
	m_pAcMenuGroups  = NULL;
	InitMenuPanel();
}
CAcadRibPanelWrapper::~CAcadRibPanelWrapper()
{
	if (m_pAcMenuBar)
	{
		m_pAcMenuBar->Release();
	}
	if (m_pAcMenuGroups)
	{
		m_pAcMenuGroups->Release();
	}
	if (m_pAcApplication)
	{
		m_pAcApplication->Update();
		m_pAcApplication->Release();
	}
}

bool CAcadRibPanelWrapper::InitMenuPanel()
{
	TRY 
	{
		HRESULT hr = NOERROR;
		LPUNKNOWN lpUnk = NULL;

		LPDISPATCH pAcDisp = acedGetIDispatch(TRUE);//acedGetAcadWinApp()->GetIDispatch(TRUE);

		hr = pAcDisp->QueryInterface(AutoCAD::IID_IAcadApplication, (void**)&m_pAcApplication);
		pAcDisp->Release();

		if (FAILED(hr)){
			return false;
		}
		if (m_pAcApplication)
		{
			m_pAcApplication->put_Visible(true);
			hr = m_pAcApplication->get_MenuGroups(&m_pAcMenuGroups);
			hr = m_pAcApplication->get_MenuBar(&m_pAcMenuBar);
		}
		if (FAILED(hr))
		{
			return false;
		}
		
	}
	CATCH (COleDispatchException, e)
	{
		e->ReportError();
		e->Delete();
		return false;
	}
	END_CATCH
	return true;
}

bool CAcadRibPanelWrapper::RremoveAllMenus()
{
	TRY{

		if (m_pAcMenuGroups==NULL)
		{
			return false;
		}
		long lMenuCount;
		m_pAcMenuGroups->get_Count(&lMenuCount);
		/*int i, j*/;
		AutoCAD::IAcadMenuGroup *pAcMenuGroup = NULL;
		// 移除所有菜单
		while (lMenuCount > 0)
		{
			VARIANT index;
			VariantInit(&index);
			V_VT(&index) = VT_I4;
			V_I4(&index) = 0;
			m_pAcMenuGroups->Item(index, &pAcMenuGroup);
			VariantClear(&index);

			BSTR bstrName, bstrMenuFile;
			pAcMenuGroup->get_Name(&bstrName);
			pAcMenuGroup->get_MenuFileName(&bstrMenuFile);

			pAcMenuGroup->Unload();
			pAcMenuGroup->Release();
			pAcMenuGroup = NULL;
			m_pAcMenuGroups->get_Count(&lMenuCount);
		}
	}
	CATCH (COleDispatchException, e)
	{
		e->ReportError();
		e->Delete();
		return false;
	}
	END_CATCH
	return true;
}

int CAcadRibPanelWrapper::GetMenuCntByMenuBar()
{
	long menus;
	m_pAcMenuBar->get_Count(&menus);
	return menus;
}

bool CAcadRibPanelWrapper::CreatePopuMenu( CString strMenuTitle, const CArrayMenuPopuItem& menuItems)
{

	if (m_pAcMenuGroups==NULL)
	{
		return false;
	}

	if (menuItems.GetSize() ==0)
	{
		return false;
	}
	AutoCAD::IAcadMenuGroup * pAcMenuGroup = NULL;
	AutoCAD::IAcadPopupMenus *pPopUpMenus = NULL;
	AutoCAD::IAcadPopupMenu *pPopUpMenu= NULL;
	AutoCAD::IAcadPopupMenuItem *pPopUpMenuItem= NULL;

	TRY
	{
		HRESULT hr = NOERROR;
		BSTR bstrMenuTitle = ::SysAllocString(strMenuTitle);

		//找到第一个菜单组
		VARIANT index;
		VariantInit(&index);
		V_VT(&index) = VT_I4;
		V_I4(&index) = 0;
		m_pAcMenuGroups->Item(index, &pAcMenuGroup);
		VariantClear(&index);
		if (pAcMenuGroup == NULL)
		{
			::SysFreeString(bstrMenuTitle);
			return false;
		}
		//找到所有当前的菜单
		pAcMenuGroup->get_Menus(&pPopUpMenus);

		if (pPopUpMenus == NULL)
		{
			::SysFreeString(bstrMenuTitle);
			pAcMenuGroup->Release();
			return false;
		}

		if (!m_bIsMenuLoaded) 
		{
			//临时添加一个弹出菜单组
			long barCnt =0;
			m_pAcMenuBar->get_Count(&barCnt);
 			pPopUpMenus->Add(bstrMenuTitle, &pPopUpMenu);
			VARIANT indexBar;
			VariantInit(&indexBar);
			V_VT(&indexBar) = VT_I4;
			V_I4(&indexBar) = barCnt;
			VariantClear(&indexBar);
			pPopUpMenu->InsertInMenuBar(indexBar);

			if (pPopUpMenu != NULL) 
			{
				//设置当前菜单名
				pPopUpMenu->put_Name(bstrMenuTitle);

				for (int i=0; i<menuItems.GetSize(); ++i)
				{
					SPopuMenuItemStyle popuMenuItem = menuItems[i];

					if (popuMenuItem.isSeparator)
					{
						VARIANT indexKey;
						VariantInit(&indexKey);
						V_VT(&indexKey) = VT_I4;
						V_I4(&indexKey) = i;
						pPopUpMenu->AddSeparator(indexKey, &pPopUpMenuItem);
						VariantClear(&indexKey);
					}
					else
					{
						BSTR bstrMenuItemName = ::SysAllocString(popuMenuItem.strName);
						BSTR bstrMenuItemMacro = ::SysAllocString(popuMenuItem.strMacro);
						VARIANT indexKey;
						VariantInit(&indexKey);
						V_VT(&indexKey) = VT_I4;
						V_I4(&indexKey) = i;
						hr =pPopUpMenu->AddMenuItem(indexKey, bstrMenuItemName,bstrMenuItemMacro, &pPopUpMenuItem);
						SysFreeString(bstrMenuItemName);
						SysFreeString(bstrMenuItemMacro);
						VariantClear(&indexKey);
						if (FAILED(hr))
						{
							acutPrintf(_T("创建菜单失败--%s\n"),popuMenuItem.strName);
							continue;
						}

					}

					m_bIsMenuLoaded = true;
				}
				pPopUpMenu->Release();
				if (pPopUpMenuItem)
				{
					pPopUpMenuItem->Release();
				}
				
			}
			else {
				acutPrintf(_T("\nMenu not created."));
			}
		}
		else 
		{
			VARIANT indexRet;
			VariantInit(&indexRet);
			V_VT(&indexRet) = VT_BSTR;
			V_BSTR(&indexRet) = bstrMenuTitle;
			if(pPopUpMenus){
				pPopUpMenus->RemoveMenuFromMenuBar(indexRet);
			}
			VariantClear(&indexRet);
			m_bIsMenuLoaded = false;	
		}
		::SysFreeString(bstrMenuTitle);
		pPopUpMenus->Release();
	}
	CATCH (COleDispatchException, e)
	{
		e->ReportError();
		e->Delete();
		return false;
	}
	END_CATCH
	return true;
}

AutoCAD::IAcadMenuGroup * CAcadRibPanelWrapper::FindMenuGroupByName(const CString& menuName)
{

	if (m_pAcMenuGroups==NULL)
	{
		return NULL;
	}
	TRY{
		long lMenuCount;
		m_pAcMenuGroups->get_Count(&lMenuCount);
		/*int i, j;*/
		AutoCAD::IAcadMenuGroup *pAcMenuGroup = NULL;
		for (int i=0; i<lMenuCount; i++)
		{
			VARIANT index;
			VariantInit(&index);
			V_VT(&index) = VT_I4;
			V_I4(&index) = i;
			m_pAcMenuGroups->Item(index, &pAcMenuGroup);
			VariantClear(&index);

			BSTR bstrName;
			pAcMenuGroup->get_Name(&bstrName);
			if (CString(bstrName).CompareNoCase(menuName)==0)
			{
				return pAcMenuGroup;
			}
			pAcMenuGroup->Release();
			pAcMenuGroup = NULL;
		}
	}
	CATCH (COleDispatchException, e)
	{
		e->ReportError();
		e->Delete();
		return false;
	}
	END_CATCH
	return NULL;
}

bool CAcadRibPanelWrapper::LoadMenuFromCuix(const CString& strCuixPath, const CString& workSpaceName)
{
	TRY{
		AutoCAD::IAcadMenuGroup *pAcMenuGroup = NULL;
		HRESULT hr = NOERROR;
		// 加载图形库菜单
		if (m_pAcMenuGroups==NULL)
			return false;

		BSTR bstrCuixPath = strCuixPath.AllocSysString();

		VARIANT baseMenu;
		baseMenu.vt = VT_BOOL;
		baseMenu.boolVal = TRUE;
		hr = m_pAcMenuGroups->Load(bstrCuixPath, baseMenu, &pAcMenuGroup);

		if (FAILED(hr))
		{
			SysFreeString(bstrCuixPath);
			if (pAcMenuGroup)
			{
				pAcMenuGroup->Release();
				pAcMenuGroup = NULL;
			}
			return false;
		}
		
		SysFreeString(bstrCuixPath);

		pAcMenuGroup->Release();
		pAcMenuGroup = NULL;

		resbuf rb;
		acedGetVar(_T("WSCURRENT"), &rb);
		rb.restype = RTSTR;
		CString str = workSpaceName;
		rb.resval.rstring = str.GetBuffer();
		str.ReleaseBuffer();
		acedSetVar(_T("WSCURRENT"), &rb);
	}
	CATCH (COleDispatchException, e)
	{
		e->ReportError();
		e->Delete();
		return false;
	}
	END_CATCH
	return true;
}

//一组菜单, 根据菜单创建相应的工具条
void CAcadRibPanelWrapper::CreateToolBar(const CString& GroupName,const CArrayMenuPopuItem& menuItems)
{
	TRY{
		AutoCAD::IAcadToolbars *pAcToolbars = NULL;
		AutoCAD::IAcadToolbar *pAcToolbar = NULL;
		AutoCAD::IAcadToolbarItem *pAcToolbarItem = NULL;
		HRESULT hr;

		AutoCAD::IAcadMenuGroup *pAcMenuGroup = NULL;
		pAcMenuGroup = FindMenuGroupByName(GroupName);
		if (pAcMenuGroup == NULL)
		{
			return;
		}

		if (pAcMenuGroup)
		{
			BSTR bstrName, bstrMenuFile;
			pAcMenuGroup->get_Name(&bstrName);
			pAcMenuGroup->get_MenuFileName(&bstrMenuFile);

			pAcMenuGroup->get_Toolbars(&pAcToolbars);
			pAcMenuGroup->Release();
			pAcMenuGroup = NULL;

			BSTR bstrGroupName = ::SysAllocString(GroupName);

			VARIANT index;
			VariantInit(&index);
			V_VT(&index) = VT_BSTR;
			V_BSTR(&index) = bstrGroupName;

			hr = pAcToolbars->Item(index, &pAcToolbar);
			if (FAILED(hr) && !pAcToolbar)
			{
				pAcToolbars->Add(bstrGroupName, &pAcToolbar);

				COleVariant flyOutButton;
				int nLen = menuItems.GetSize();
				for (int i = 0; i < nLen; i++)
				{
					pAcToolbar->AddToolbarButton(index, ::SysAllocString(menuItems[i].strMacro), ::SysAllocString(menuItems[i].strName), ::SysAllocString(menuItems[i].strMacro), flyOutButton, &pAcToolbarItem);
					pAcToolbarItem->Release();
				}
				//新添加的按钮图标自行设定
				pAcToolbar->put_Visible(TRUE);
				pAcToolbar->Dock(AutoCAD::acToolbarDockTop);
				pAcToolbar->Release();
			}
			else
			{
				pAcToolbar->Release();
			}	

			SysFreeString(bstrGroupName);
			VariantClear(&index);
			pAcToolbars->Release();
		}
	}
	CATCH (COleDispatchException, e)
	{
		e->ReportError();
		e->Delete();
		return ;
	}
	END_CATCH
}

void CAcadRibPanelWrapper::BrushRib()
{
	acedGetAcadFrame()->Invalidate(TRUE);	

	CMDIFrameWnd*pMainFram=(CMDIFrameWnd*)acedGetAcadWinApp()->m_pMainWnd;	
	if(pMainFram)
	{
		CWnd* pWnd = pMainFram->GetActiveWindow();
		int nCnt =AfxGetApp()->GetOpenDocumentCount();
		while(pWnd){
			CString strTitle;
			pWnd->GetWindowText(strTitle);
			pWnd->Invalidate(TRUE);
			pWnd = pWnd->GetNextWindow();
		}

	}
}

void CAcadRibPanelWrapper::GetRibStyle()
{
}

void CAcadRibPanelWrapper::SetRibStyle()
{
}

void CAcadRibPanelWrapper::ModifyRibStyle()
{
}

#include "..\samples\com\AsdkMfcComSamp_dg\CAcadApplication.h"
#include "..\samples\com\AsdkMfcComSamp_dg\CAcadDocument.h"
#include "..\samples\com\AsdkMfcComSamp_dg\CAcadModelSpace.h"
#include "..\samples\com\AsdkMfcComSamp_dg\CAcadMenuBar.h"
#include "..\samples\com\AsdkMfcComSamp_dg\CAcadMenuGroup.h"
#include "..\samples\com\AsdkMfcComSamp_dg\CAcadMenuGroups.h"
#include "..\samples\com\AsdkMfcComSamp_dg\CAcadPopupMenu.h"
#include "..\samples\com\AsdkMfcComSamp_dg\CAcadPopupMenus.h"
#include "acadi.h"

BOOL CAcadRibPanelWrapper::creatMenuOnCurren(CString strMenuTitle,const CArrayMenuPopuItem& menuItems)
{
	TRY
	{
		CAcadApplication IAcad(acedGetAcadWinApp()->GetIDispatch(TRUE));
		AcadMenuGroup *pAcMenuGroup = NULL;
		CAcadMenuBar IMenuBar(IAcad.get_MenuBar());

		long numberOfMenus;
		numberOfMenus = IMenuBar.get_Count();

		CAcadMenuGroups IMenuGroups(IAcad.get_MenuGroups());

		VARIANT index;
		VariantInit(&index);
		V_VT(&index) = VT_I4;
		V_I4(&index) = 0;

		CAcadMenuGroup IMenuGroup(IMenuGroups.Item(index));

		CAcadPopupMenus IPopUpMenus(IMenuGroup.get_Menus());

		VariantInit(&index);
		V_VT(&index) = VT_BSTR;
		V_BSTR(&index) = strMenuTitle.AllocSysString();

		IDispatch* pDisp=NULL;

		//see if the menu is already there
		TRY
		{
			pDisp = IPopUpMenus.Item(index); 
			pDisp->AddRef();
		} 
		CATCH(COleDispatchException,e)
		{
			
		}
		END_CATCH;
		if (pDisp==NULL) 
		{
			//create it
			CAcadPopupMenu IPopUpMenu(IPopUpMenus.Add(strMenuTitle));

			int nLen = menuItems.GetSize();
			int iTmp = 0;
			for (int i = 0; i < nLen; i++)
			{
				
				VariantInit(&index);
				V_VT(&index) = VT_I4;
				V_I4(&index) = iTmp;
				IPopUpMenu.AddMenuItem(index, menuItems[i].strName, menuItems[i].strMacro);
				if (menuItems[i].isSeparator)
				{
					VariantInit(&index);
					V_VT(&index) = VT_I4;
					V_I4(&index) = iTmp;
					IPopUpMenu.AddSeparator(index);
					iTmp++;
				}
			}
			pDisp = IPopUpMenu.m_lpDispatch;
			pDisp->AddRef();
		}

		CAcadPopupMenu IPopUpMenu(pDisp);
		if (!IPopUpMenu.get_OnMenuBar())
		{
			VariantInit(&index);
			V_VT(&index) = VT_I4;
			V_I4(&index) = numberOfMenus - 2;
			IPopUpMenu.InsertInMenuBar(index);
		}
		else
		{
			VariantInit(&index);
			V_VT(&index) = VT_BSTR;
			V_BSTR(&index) = strMenuTitle.AllocSysString();
			IPopUpMenus.RemoveMenuFromMenuBar(index);
			VariantClear(&index);
		}
		pDisp->Release();
	}
	CATCH(COleDispatchException,e)
	{
		e->ReportError();
		e->Delete();
	}
	END_CATCH;
	return TRUE;
}
