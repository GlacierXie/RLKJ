/********************************************************************
	created:	2015/10/09
	file base:	ZWMainRibPanelWrapper
	file ext:	h

	purpose:	增加ribon菜单的驱动
*********************************************************************/
#ifndef __CAcadRibPanelWrapper_H_
#define __CAcadRibPanelWrapper_H_

#include "tchar.h"
#include "resource.h"       // main symbols
#include <AtlBase.h>
#pragma warning( disable : 4278 )
#import "acax19ENU.tlb" no_implementation raw_interfaces_only named_guids
#pragma warning( default : 4278 )


//测试路径
/////////////////////////////////////////////////////////////////////////////
typedef struct _SPopuMenuItemStyle
{
	_SPopuMenuItemStyle():strName(_T("")),
						 strMacro(_T("")),
						 isSeparator(false)
	{}
	_SPopuMenuItemStyle( CString menuName, CString menuMacro,bool isSep= false)
	{
		strName = menuName;
		strMacro = menuMacro;
		isSeparator  = isSep;
	}

	_SPopuMenuItemStyle(const _SPopuMenuItemStyle& rhl)
	{
		*this = rhl;
	}

	_SPopuMenuItemStyle& operator =(const _SPopuMenuItemStyle&rhl )
	{
		strName = rhl.strName;
		strMacro = rhl.strMacro;
		isSeparator = rhl.isSeparator;
		return *this;
	}

	CString strName;					//命令名
	CString strMacro;					//宏名
	BOOL	isSeparator;				//是否是分割线

} SPopuMenuItemStyle;

typedef CArray<SPopuMenuItemStyle> CArrayMenuPopuItem;
class  CAcadRibPanelWrapper
{
public:
    CAcadRibPanelWrapper();
	~CAcadRibPanelWrapper();

public:
	//{@ 增加相关的操作

	//	@bref: 初始化相关环境
	bool InitMenuPanel();

	//	@bref: 移除所有菜单
	bool RremoveAllMenus();

	//	@bref: 加载菜单及工作空间
	bool LoadMenuFromCuix(const CString& strCuixPath, const CString& workSpaceName);

	//	@bref: 创建工具栏
	void CreateToolBar(const CString& GroupName, const CArrayMenuPopuItem& menuItems);

	//	@bref: 得到menuBar上菜单数量
	int GetMenuCntByMenuBar();

	//	@bref: 创建一个快捷菜单
	//			menuItems->MenuGroups:  命令命名
	//			menuItems->MenuMacros:  执行命令的命令宏，和ARX中手动定义的命令宏一致
	//	  返回: 是否已经创建及加载菜单
	bool CreatePopuMenu( CString strMenuTitle,const CArrayMenuPopuItem& menuItems);

	BOOL creatMenuOnCurren(CString strMenuTitle, const CArrayMenuPopuItem& menuItems);

	void BrushRib();
	void GetRibStyle();
	void SetRibStyle();
	void ModifyRibStyle();
	//@}

protected:
	AutoCAD::IAcadMenuGroup * FindMenuGroupByName(const CString& menuName);
private:
    //short mNumber;
	AutoCAD::IAcadApplication *m_pAcApplication;
	AutoCAD::IAcadMenuGroups  *m_pAcMenuGroups ;
	AutoCAD::IAcadMenuBar	  *m_pAcMenuBar;

	static bool m_bIsMenuLoaded;
};

#endif
